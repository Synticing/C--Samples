' (c) Copyright Microsoft Corporation.
' This source is subject to the Microsoft Public License (Ms-PL).
' Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
' All other rights reserved.

Imports System.Text
Imports Microsoft.Kinect

Namespace WorkingWithDepthData
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
		End Sub

        Private Const MaxDepthDistance As Single = 4095 ' max value returned
		Private Const MinDepthDistance As Single = 850 ' min value returned
		Private Const MaxDepthDistanceOffset As Single = MaxDepthDistance - MinDepthDistance

		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			AddHandler kinectSensorChooser1.KinectSensorChanged, AddressOf kinectSensorChooser1_KinectSensorChanged

		End Sub

		Private Sub kinectSensorChooser1_KinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)

			Dim oldSensor = CType(e.OldValue, KinectSensor)

			'stop the old sensor
			If oldSensor IsNot Nothing Then
				oldSensor.Stop()
				oldSensor.AudioSource.Stop()
			End If

			'get the new sensor
			Dim newSensor = CType(e.NewValue, KinectSensor)
			If newSensor Is Nothing Then
				Return
			End If

			'turn on features that you need
			newSensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30)
			newSensor.SkeletonStream.Enable()

			'sign up for events if you want to get at API directly
			AddHandler newSensor.AllFramesReady, AddressOf newSensor_AllFramesReady


			Try
				newSensor.Start()
			Catch e1 As System.IO.IOException
				'this happens if another app is using the Kinect
				kinectSensorChooser1.AppConflictOccurred()
			End Try
		End Sub

		Private Sub newSensor_AllFramesReady(ByVal sender As Object, ByVal e As AllFramesReadyEventArgs)

			Using depthFrame As DepthImageFrame = e.OpenDepthImageFrame()
				If depthFrame Is Nothing Then
					Return
				End If

				Dim pixels() As Byte = GenerateColoredBytes(depthFrame)

				'number of bytes per row width * 4 (B,G,R,Empty)
				Dim stride As Integer = depthFrame.Width * 4

				'create image
				image1.Source = BitmapSource.Create(depthFrame.Width, depthFrame.Height, 96, 96, PixelFormats.Bgr32, Nothing, pixels, stride)

			End Using
		End Sub


		Private Function GenerateColoredBytes(ByVal depthFrame As DepthImageFrame) As Byte()

			'get the raw data from kinect with the depth for every pixel
			Dim rawDepthData(depthFrame.PixelDataLength - 1) As Short
			depthFrame.CopyPixelDataTo(rawDepthData)

			'use depthFrame to create the image to display on-screen
			'depthFrame contains color information for all pixels in image
			'Height x Width x 4 (Red, Green, Blue, empty byte)
			Dim pixels(depthFrame.Height * depthFrame.Width * 4 - 1) As Byte

			'Bgr32  - Blue, Green, Red, empty byte
			'Bgra32 - Blue, Green, Red, transparency 
			'You must set transparency for Bgra as .NET defaults a byte to 0 = fully transparent

			'hardcoded locations to Blue, Green, Red (BGR) index positions       
			Const BlueIndex As Integer = 0
			Const GreenIndex As Integer = 1
			Const RedIndex As Integer = 2


			'loop through all distances
			'pick a RGB color based on distance
			Dim depthIndex As Integer = 0
			Dim colorIndex As Integer = 0
			Do While depthIndex < rawDepthData.Length AndAlso colorIndex < pixels.Length
				'get the player (requires skeleton tracking enabled for values)
                Dim player As Integer = CInt(rawDepthData(depthIndex)) And DepthImageFrame.PlayerIndexBitmask

				'gets the depth value
                Dim depth As Integer = CInt(rawDepthData(depthIndex)) >> DepthImageFrame.PlayerIndexBitmaskWidth

				'.9M or 2.95'
				If depth <= 900 Then
					'we are very close
					pixels(colorIndex + BlueIndex) = 255
					pixels(colorIndex + GreenIndex) = 0
					pixels(colorIndex + RedIndex) = 0

				' .9M - 2M or 2.95' - 6.56'
				ElseIf depth > 900 AndAlso depth < 2000 Then
					'we are a bit further away
					pixels(colorIndex + BlueIndex) = 0
					pixels(colorIndex + GreenIndex) = 255
					pixels(colorIndex + RedIndex) = 0
				' 2M+ or 6.56'+
				ElseIf depth > 2000 Then
					'we are the farthest
					pixels(colorIndex + BlueIndex) = 0
					pixels(colorIndex + GreenIndex) = 0
					pixels(colorIndex + RedIndex) = 255
				End If


				'//equal coloring for monochromatic histogram
				Dim intensity As Byte = CalculateIntensityFromDepth(depth)
				pixels(colorIndex + BlueIndex) = intensity
				pixels(colorIndex + GreenIndex) = intensity
				pixels(colorIndex + RedIndex) = intensity


				'Color all players "gold"
				If player > 0 Then
					pixels(colorIndex + BlueIndex) = Colors.Gold.B
					pixels(colorIndex + GreenIndex) = Colors.Gold.G
					pixels(colorIndex + RedIndex) = Colors.Gold.R
				End If

				depthIndex += 1
				colorIndex += 4
			Loop


			Return pixels
		End Function


		Public Shared Function CalculateIntensityFromDepth(ByVal distance As Integer) As Byte
			'formula for calculating monochrome intensity for histogram
			Return CByte(255 - (255 * Math.Max(distance - MinDepthDistance, 0) / (MaxDepthDistanceOffset)))
		End Function




        Private Sub Window_Closing(sender As System.Object, e As System.ComponentModel.CancelEventArgs)
            StopKinect(kinectSensorChooser1.Kinect)
        End Sub

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

    End Class

End Namespace

