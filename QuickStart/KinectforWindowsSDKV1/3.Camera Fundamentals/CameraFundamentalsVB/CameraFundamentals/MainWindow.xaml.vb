' (c) Copyright Microsoft Corporation.
' This source is subject to the Microsoft Public License (Ms-PL).
' Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
' All other rights reserved.


Imports System.Text
Imports Microsoft.Kinect


Namespace CameraFundamentals
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
		End Sub


		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			'sign up for the event
			AddHandler kinectSensorChooser1.KinectSensorChanged, AddressOf kinectSensorChooser1_KinectSensorChanged


		End Sub

		Private Sub kinectSensorChooser1_KinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)

			Dim oldSensor = CType(e.OldValue, KinectSensor)

			'stop the old sensor
			If oldSensor IsNot Nothing Then
				StopKinect(oldSensor)
			End If

			'get the new sensor
			Dim newSensor = CType(e.NewValue, KinectSensor)
			If newSensor Is Nothing Then
				Return
			End If

			'newSensor.AllFramesReady += new EventHandler<AllFramesReadyEventArgs>(newSensor_AllFramesReady);

			'turn on features that you need
			newSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30)
			newSensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30)
			newSensor.SkeletonStream.Enable()

			Try
				newSensor.Start()
			Catch e1 As System.IO.IOException
				'this happens if another app is using the Kinect
				kinectSensorChooser1.AppConflictOccurred()
			End Try
		End Sub


		'this event fires when Color/Depth/Skeleton are synchronized
		Private Sub newSensor_AllFramesReady(ByVal sender As Object, ByVal e As AllFramesReadyEventArgs)
			'using (ColorImageFrame colorFrame = e.OpenColorImageFrame())
			'{
			'    if (colorFrame == null)
			'    {
			'        return;
			'    }

			'    byte[] pixels = new byte[colorFrame.PixelDataLength];
			'    colorFrame.CopyPixelDataTo(pixels);

			'    int stride = colorFrame.Width * 4;
			'    image1.Source =
			'        BitmapSource.Create(colorFrame.Width, colorFrame.Height,
			'        96, 96, PixelFormats.Bgr32, null, pixels, stride); 




			'}
		End Sub

		Private Sub Window_Closed(ByVal sender As Object, ByVal e As EventArgs)
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

        Private Sub Window_Closing(sender As System.Object, e As System.ComponentModel.CancelEventArgs)
            StopKinect(kinectSensorChooser1.Kinect)
        End Sub
    End Class
End Namespace
