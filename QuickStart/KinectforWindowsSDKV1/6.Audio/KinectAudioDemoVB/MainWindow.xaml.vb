Imports Microsoft.Speech.Recognition
Imports Microsoft.Speech.AudioFormat
Imports Microsoft.Kinect
Imports System.Windows.Threading
Imports System.Threading
Imports System.IO

'------------------------------------------------------------------------------
' <copyright file="MainWindow.xaml.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace KinectAudioDemo

	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window

		Private Const AngleChangeSmoothingFactor As Double = 0.35
		Private Const AcceptedSpeechPrefix As String = "Accepted_"
		Private Const RejectedSpeechPrefix As String = "Rejected_"

		Private Const WaveImageWidth As Integer = 500
		Private Const WaveImageHeight As Integer = 100

		Private ReadOnly redBrush As New SolidColorBrush(Colors.Red)
		Private ReadOnly greenBrush As New SolidColorBrush(Colors.Green)
		Private ReadOnly blueBrush As New SolidColorBrush(Colors.Blue)
		Private ReadOnly blackBrush As New SolidColorBrush(Colors.Black)

		Private ReadOnly bitmapWave As WriteableBitmap
		Private ReadOnly pixels() As Byte
		Private ReadOnly energyBuffer(WaveImageWidth - 1) As Double
		Private ReadOnly blackPixels(WaveImageWidth * WaveImageHeight - 1) As Byte
		Private ReadOnly fullImageRect As New Int32Rect(0, 0, WaveImageWidth, WaveImageHeight)

		Private kinect As KinectSensor
		Private angle As Double
		Private running As Boolean = True
		Private readyTimer As DispatcherTimer
		Private stream As EnergyCalculatingPassThroughStream
		Private speechRecognizer As SpeechRecognitionEngine

		Public Sub New()
			InitializeComponent()

			Dim colorList = New List(Of Color) From {Colors.Black, Colors.Green}
			Me.bitmapWave = New WriteableBitmap(WaveImageWidth, WaveImageHeight, 96, 96, PixelFormats.Indexed1, New BitmapPalette(colorList))

			Me.pixels = New Byte(WaveImageWidth - 1){}
			For i As Integer = 0 To Me.pixels.Length - 1
				Me.pixels(i) = &Hff
			Next i

			imgWav.Source = Me.bitmapWave

			AddHandler SensorChooser.KinectSensorChanged, AddressOf SensorChooserKinectSensorChanged
		End Sub

		Private Shared Function GetKinectRecognizer() As RecognizerInfo
			Dim matchingFunc As Func(Of RecognizerInfo, Boolean) = Function(r)
				Dim value As String
				r.AdditionalInfo.TryGetValue("Kinect", value)
				Return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) AndAlso "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase)
			End Function
			Return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault()
		End Function

		Private Sub SensorChooserKinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)
			Dim oldSensor As KinectSensor = TryCast(e.OldValue, KinectSensor)
			If oldSensor IsNot Nothing Then
				Me.UninitializeKinect()
			End If

			Dim newSensor As KinectSensor = TryCast(e.NewValue, KinectSensor)
			Me.kinect = newSensor

			' Only enable this checkbox if we have a sensor
			enableAec.IsEnabled = Me.kinect IsNot Nothing

			If newSensor IsNot Nothing Then
				Me.InitializeKinect()
			End If
		End Sub

		Private Sub InitializeKinect()
			Dim sensor = Me.kinect
			Me.speechRecognizer = Me.CreateSpeechRecognizer()
			Try
				sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30)
				sensor.Start()
			Catch e1 As Exception
				SensorChooser.AppConflictOccurred()
				Return
			End Try

			If Me.speechRecognizer IsNot Nothing AndAlso sensor IsNot Nothing Then
				' NOTE: Need to wait 4 seconds for device to be ready to stream audio right after initialization
				Me.readyTimer = New DispatcherTimer()
				AddHandler Me.readyTimer.Tick, AddressOf ReadyTimerTick
				Me.readyTimer.Interval = New TimeSpan(0, 0, 4)
				Me.readyTimer.Start()

				Me.ReportSpeechStatus("Initializing audio stream...")
				Me.UpdateInstructionsText(String.Empty)

				AddHandler Me.Closing, AddressOf MainWindowClosing
			End If

			Me.running = True
		End Sub

		Private Sub ReadyTimerTick(ByVal sender As Object, ByVal e As EventArgs)
			Me.Start()
			Me.ReportSpeechStatus("Ready to recognize speech!")
			Me.UpdateInstructionsText("Say: 'red', 'green' or 'blue'")
			Me.readyTimer.Stop()
			Me.readyTimer = Nothing
		End Sub

		Private Sub UninitializeKinect()
			Dim sensor = Me.kinect
			Me.running = False
			If Me.speechRecognizer IsNot Nothing AndAlso sensor IsNot Nothing Then
				sensor.AudioSource.Stop()
				sensor.Stop()
				Me.speechRecognizer.RecognizeAsyncCancel()
				Me.speechRecognizer.RecognizeAsyncStop()
			End If

			If Me.readyTimer IsNot Nothing Then
				Me.readyTimer.Stop()
				Me.readyTimer = Nothing
			End If
		End Sub

		Private Sub MainWindowClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
			Me.UninitializeKinect()
		End Sub

		Private Function CreateSpeechRecognizer() As SpeechRecognitionEngine
			Dim ri As RecognizerInfo = GetKinectRecognizer()
			If ri Is Nothing Then
				MessageBox.Show("There was a problem initializing Speech Recognition." & ControlChars.CrLf & "Ensure you have the Microsoft Speech SDK installed.", "Failed to load Speech SDK", MessageBoxButton.OK, MessageBoxImage.Error)
				Me.Close()
				Return Nothing
			End If

			Dim sre As SpeechRecognitionEngine
			Try
				sre = New SpeechRecognitionEngine(ri.Id)
			Catch
				MessageBox.Show("There was a problem initializing Speech Recognition." & ControlChars.CrLf & "Ensure you have the Microsoft Speech SDK installed and configured.", "Failed to load Speech SDK", MessageBoxButton.OK, MessageBoxImage.Error)
				Me.Close()
				Return Nothing
			End Try

			Dim grammar = New Choices()
			grammar.Add("red")
			grammar.Add("green")
			grammar.Add("blue")
			grammar.Add("Camera on")
			grammar.Add("Camera off")

			Dim gb = New GrammarBuilder With {.Culture = ri.Culture}
			gb.Append(grammar)

			' Create the actual Grammar instance, and then load it into the speech recognizer.
			Dim g = New Grammar(gb)

			sre.LoadGrammar(g)
			AddHandler sre.SpeechRecognized, AddressOf SreSpeechRecognized
			AddHandler sre.SpeechHypothesized, AddressOf SreSpeechHypothesized
			AddHandler sre.SpeechRecognitionRejected, AddressOf SreSpeechRecognitionRejected

			Return sre
		End Function

		Private Sub RejectSpeech(ByVal result As RecognitionResult)
			Dim status As String = "Rejected: " & (If(result Is Nothing, String.Empty, result.Text & " " & result.Confidence))
			Me.ReportSpeechStatus(status)

			Dispatcher.BeginInvoke(New Action(Sub() tbColor.Background = blackBrush), DispatcherPriority.Normal)
		End Sub

		Private Sub SreSpeechRecognitionRejected(ByVal sender As Object, ByVal e As SpeechRecognitionRejectedEventArgs)
			Me.RejectSpeech(e.Result)
		End Sub

		Private Sub SreSpeechHypothesized(ByVal sender As Object, ByVal e As SpeechHypothesizedEventArgs)
			Me.ReportSpeechStatus("Hypothesized: " & e.Result.Text & " " & e.Result.Confidence)
		End Sub

		Private Sub SreSpeechRecognized(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
			Dim brush As SolidColorBrush

			If e.Result.Confidence < 0.5 Then
				Me.RejectSpeech(e.Result)
				Return
			End If

			Select Case e.Result.Text.ToUpperInvariant()
				Case "RED"
					brush = Me.redBrush
				Case "GREEN"
					brush = Me.greenBrush
				Case "BLUE"
					brush = Me.blueBrush
				Case "CAMERA ON"
					Process.Start("notepad.exe")
					Me.kinectColorViewer1.Visibility = Visibility.Visible
					brush = Me.blackBrush
				Case "CAMERA OFF"
					Me.kinectColorViewer1.Visibility = Visibility.Hidden
					brush = Me.blackBrush
				Case Else
					brush = Me.blackBrush
			End Select

			Dim status As String = "Recognized: " & e.Result.Text & " " & e.Result.Confidence
			Me.ReportSpeechStatus(status)

			Dispatcher.BeginInvoke(New Action(Sub() tbColor.Background = brush), DispatcherPriority.Normal)
		End Sub

		Private Sub ReportSpeechStatus(ByVal status As String)
			Dispatcher.BeginInvoke(New Action(Sub() tbSpeechStatus.Text = status), DispatcherPriority.Normal)
		End Sub

		Private Sub UpdateInstructionsText(ByVal instructions As String)
			Dispatcher.BeginInvoke(New Action(Sub() tbColor.Text = instructions), DispatcherPriority.Normal)
		End Sub

		Private Sub Start()
			Dim audioSource = Me.kinect.AudioSource
			audioSource.BeamAngleMode = BeamAngleMode.Adaptive
			Dim kinectStream = audioSource.Start()
			Me.stream = New EnergyCalculatingPassThroughStream(kinectStream)
			Me.speechRecognizer.SetInputToAudioStream(Me.stream, New SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, Nothing))
			Me.speechRecognizer.RecognizeAsync(RecognizeMode.Multiple)
			Dim t = New Thread(AddressOf Me.PollSoundSourceLocalization)
			t.Start()
		End Sub

		Private Sub EnableAecChecked(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Dim enableAecCheckBox As CheckBox = CType(sender, CheckBox)
			If enableAecCheckBox.IsChecked IsNot Nothing Then
				Me.kinect.AudioSource.EchoCancellationMode = If(enableAecCheckBox.IsChecked.Value, EchoCancellationMode.CancellationAndSuppression, EchoCancellationMode.None)
			End If
		End Sub

		Private Sub PollSoundSourceLocalization()
			Do While Me.running
				Dim audioSource = Me.kinect.AudioSource
				If audioSource.SoundSourceAngleConfidence > 0.5 Then
					' Smooth the change in angle
					Dim a As Double = AngleChangeSmoothingFactor * audioSource.SoundSourceAngleConfidence
					Me.angle = ((1 - a) * Me.angle) + (a * audioSource.SoundSourceAngle)

					Dispatcher.BeginInvoke(New Action(Sub() rotTx.Angle = -angle), DispatcherPriority.Normal)
				End If

				Dispatcher.BeginInvoke(New Action(Sub()
					clipConf.Rect = New Rect(0, 0, 100 + (600 * audioSource.SoundSourceAngleConfidence), 50)
					Dim sConf As String = String.Format("Conf: {0:0.00}", audioSource.SoundSourceAngleConfidence)
					tbConf.Text = sConf
					stream.GetEnergy(energyBuffer)
					Me.bitmapWave.WritePixels(fullImageRect, blackPixels, WaveImageWidth, 0)
					For i As Integer = 1 To energyBuffer.Length - 1
						Dim energy As Integer = CInt(Fix(energyBuffer(i) * 5))
						Dim r As New Int32Rect(i, (WaveImageHeight \ 2) - energy, 1, 2 * energy)
						Me.bitmapWave.WritePixels(r, pixels, 1, 0)
					Next i
				End Sub), DispatcherPriority.Normal)

				Thread.Sleep(50)
			Loop
		End Sub

		Private Class EnergyCalculatingPassThroughStream
			Inherits Stream
			Private Const SamplesPerPixel As Integer = 10

			Private ReadOnly energy(WaveImageWidth - 1) As Double
			Private ReadOnly syncRoot As New Object()
			Private ReadOnly baseStream As Stream

			Private index As Integer
			Private sampleCount As Integer
            Private avgSample As Double

			Public Sub New(ByVal stream As Stream)
				Me.baseStream = stream
			End Sub

			Public Overrides ReadOnly Property Length() As Long
				Get
					Return Me.baseStream.Length
				End Get
			End Property

			Public Overrides Property Position() As Long
				Get
					Return Me.baseStream.Position
				End Get
				Set(ByVal value As Long)
					Me.baseStream.Position = value
				End Set
			End Property

			Public Overrides ReadOnly Property CanRead() As Boolean
				Get
					Return Me.baseStream.CanRead
				End Get
			End Property

			Public Overrides ReadOnly Property CanSeek() As Boolean
				Get
					Return Me.baseStream.CanSeek
				End Get
			End Property

			Public Overrides ReadOnly Property CanWrite() As Boolean
				Get
					Return Me.baseStream.CanWrite
				End Get
			End Property

			Public Overrides Sub Flush()
				Me.baseStream.Flush()
			End Sub

			Public Sub GetEnergy(ByVal energyBuffer() As Double)
				SyncLock Me.syncRoot
					Dim energyIndex As Integer = Me.index
					For i As Integer = 0 To Me.energy.Length - 1
						energyBuffer(i) = Me.energy(energyIndex)
						energyIndex += 1
						If energyIndex >= Me.energy.Length Then
							energyIndex = 0
						End If
					Next i
				End SyncLock
			End Sub

			Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
				Dim retVal As Integer = Me.baseStream.Read(buffer, offset, count)
				Const A As Double = 0.3
				SyncLock Me.syncRoot
					For i As Integer = 0 To retVal - 1 Step 2
						Dim sample As Short = BitConverter.ToInt16(buffer, i + offset)

                        Me.avgSample += CInt(sample) * CInt(sample)
						Me.sampleCount += 1

						If Me.sampleCount = SamplesPerPixel Then
							Me.avgSample /= SamplesPerPixel

							Me.energy(Me.index) =.2 + ((Me.avgSample * 11) / (Integer.MaxValue / 2))
							Me.energy(Me.index) = If(Me.energy(Me.index) > 10, 10, Me.energy(Me.index))

							If Me.index > 0 Then
								Me.energy(Me.index) = (Me.energy(Me.index) * A) + ((1 - A) * Me.energy(Me.index - 1))
							End If

							Me.index += 1
							If Me.index >= Me.energy.Length Then
								Me.index = 0
							End If

							Me.avgSample = 0
							Me.sampleCount = 0
						End If
					Next i
				End SyncLock

				Return retVal
			End Function

			Public Overrides Function Seek(ByVal offset As Long, ByVal origin As SeekOrigin) As Long
				Return Me.baseStream.Seek(offset, origin)
			End Function

			Public Overrides Sub SetLength(ByVal value As Long)
				Me.baseStream.SetLength(value)
			End Sub

			Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
				Me.baseStream.Write(buffer, offset, count)
			End Sub
		End Class
	End Class
End Namespace
