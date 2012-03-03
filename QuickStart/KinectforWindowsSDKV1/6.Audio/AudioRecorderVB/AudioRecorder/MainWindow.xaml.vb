' (c) Copyright Microsoft Corporation.
' This source is subject to the Microsoft Public License (Ms-PL).
' Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
' All other rights reserved.
Imports System.IO
Imports System.Text
Imports System.Threading

Imports Microsoft.Kinect

Namespace AudioRecorder
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()

			AddHandler FinishedRecording, AddressOf MainWindow_FinishedRecording
		End Sub

		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			AddHandler kinectSensorChooser1.KinectSensorChanged, AddressOf kinectSensorChooser1_KinectSensorChanged
		End Sub

		Private Sub kinectSensorChooser1_KinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)
			Dim old = CType(e.OldValue, KinectSensor)

			If old IsNot Nothing Then
				old.Stop()
				old.AudioSource.Stop()
			End If

			Dim sensor As KinectSensor = CType(e.NewValue, KinectSensor)
			sensor.Start()

		End Sub

		Private _amountOfTimeToRecord As Double
		Private _lastRecordedFileName As String
		Private Event FinishedRecording As RoutedEventHandler

		Private Sub RecordButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			RecordButton.IsEnabled = False
			PlayButton.IsEnabled = False
			_amountOfTimeToRecord = RecordForTimeSpan.Value
			_lastRecordedFileName = Date.Now.ToString("yyyyMMddHHmmss") & "_wav.wav"

            '//Start recording audio on new thread
            Dim t As New Thread(AddressOf RecordAudio)
            t.Start(kinectSensorChooser1.Kinect)

			'You can also Record audio synchronously but it will "freeze" the UI 
			'RecordAudio(kinectSensorChooser1.Kinect); 


		End Sub

		Private Sub PlayButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			If (Not String.IsNullOrEmpty(_lastRecordedFileName)) AndAlso File.Exists(_lastRecordedFileName) Then
				audioPlayer.Source = New Uri(_lastRecordedFileName, UriKind.RelativeOrAbsolute)
				audioPlayer.LoadedBehavior = MediaState.Play
				audioPlayer.UnloadedBehavior = MediaState.Close
			End If
		End Sub

		Private Sub MainWindow_FinishedRecording(ByVal sender As Object, ByVal e As RoutedEventArgs)
			'This is only required if recording on a separate thread to ensure that enabling the buttons
			'happens on the UI thread
			Dispatcher.BeginInvoke(New ThreadStart(AddressOf ReenableButtons))

			'use this if recording on the same thread
			'ReenableButtons(); 
		End Sub

		Private Sub ReenableButtons()
			RecordButton.IsEnabled = True
			PlayButton.IsEnabled = True
		End Sub


		Private Sub RecordAudio(ByVal kinectSensor As Object)

			Dim _sensor As KinectSensor = CType(kinectSensor, KinectSensor)
			RecordAudio(_sensor)
		End Sub

		Private Sub RecordAudio(ByVal kinectSensor As KinectSensor)

			If kinectSensor Is Nothing Then
				Return
			End If

			Dim recordingLength As Integer = CInt(Fix(_amountOfTimeToRecord)) * 2 * 16000
			Dim buffer(1023) As Byte


			Using _fileStream As New FileStream(_lastRecordedFileName, FileMode.Create)
				WriteWavHeader(_fileStream, recordingLength)

				'Start capturing audio                               
				Using audioStream As Stream = kinectSensor.AudioSource.Start()
					'Simply copy the data from the stream down to the file
					Dim count As Integer, totalCount As Integer = 0
					count = audioStream.Read(buffer, 0, buffer.Length)
					Do While count > 0 AndAlso totalCount < recordingLength
						_fileStream.Write(buffer, 0, count)
						totalCount += count
						count = audioStream.Read(buffer, 0, buffer.Length)
					Loop
				End Using
			End Using


			RaiseEvent FinishedRecording(Nothing, Nothing)
		End Sub


		''' <summary>
		''' A bare bones WAV file header writer
		''' </summary>        
		Private Shared Sub WriteWavHeader(ByVal stream As Stream, ByVal dataLength As Integer)
			'We need to use a memory stream because the BinaryWriter will close the underlying stream when it is closed
			Using memStream = New MemoryStream(64)
				Dim cbFormat As Integer = 18 'sizeof(WAVEFORMATEX)
				Dim format As New WAVEFORMATEX() With {.wFormatTag = 1, .nChannels = 1, .nSamplesPerSec = 16000, .nAvgBytesPerSec = 32000, .nBlockAlign = 2, .wBitsPerSample = 16, .cbSize = 0}

				Using bw = New BinaryWriter(memStream)
					'RIFF header
					WriteString(memStream, "RIFF")
					bw.Write(dataLength + cbFormat + 4) 'File size - 8
					WriteString(memStream, "WAVE")
					WriteString(memStream, "fmt ")
					bw.Write(cbFormat)

					'WAVEFORMATEX
					bw.Write(format.wFormatTag)
					bw.Write(format.nChannels)
					bw.Write(format.nSamplesPerSec)
					bw.Write(format.nAvgBytesPerSec)
					bw.Write(format.nBlockAlign)
					bw.Write(format.wBitsPerSample)
					bw.Write(format.cbSize)

					'data header
					WriteString(memStream, "data")
					bw.Write(dataLength)
					memStream.WriteTo(stream)
				End Using
			End Using
		End Sub

		Private Shared Sub WriteString(ByVal stream As Stream, ByVal s As String)
			Dim bytes() As Byte = Encoding.ASCII.GetBytes(s)
			stream.Write(bytes, 0, bytes.Length)
		End Sub

		Private Structure WAVEFORMATEX
			Public wFormatTag As UShort
			Public nChannels As UShort
			Public nSamplesPerSec As UInteger
			Public nAvgBytesPerSec As UInteger
			Public nBlockAlign As UShort
			Public wBitsPerSample As UShort
			Public cbSize As UShort
		End Structure

        Private Sub StopKinect(ByVal sensor As KinectSensor)
            If sensor IsNot Nothing Then
                If sensor.IsRunning Then
                    'stop sensor
                    sensor.Stop()

                    'stop audio if not null
                    If sensor.AudioSource IsNot Nothing Then
                        sensor.AudioSource.Stop()
                    End If
                End If
            End If
        End Sub


        Private Sub Window_Closing(sender As System.Object, e As System.ComponentModel.CancelEventArgs)
            StopKinect(kinectSensorChooser1.Kinect)
        End Sub
    End Class
End Namespace
