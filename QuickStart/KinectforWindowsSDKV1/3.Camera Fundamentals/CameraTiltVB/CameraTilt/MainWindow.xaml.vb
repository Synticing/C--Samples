' (c) Copyright Microsoft Corporation.
' This source is subject to the Microsoft Public License (Ms-PL).
' Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
' All other rights reserved.

Imports System.Text
Imports Microsoft.Kinect

Namespace CameraTilt
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
		End Sub





		Private Sub button1_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			button1.IsEnabled = False

			'Set angle to slider1 value
			If kinectSensorChooser1.Kinect IsNot Nothing AndAlso kinectSensorChooser1.Kinect.IsRunning Then
				kinectSensorChooser1.Kinect.ElevationAngle = CInt(Fix(slider1.Value))
				lblCurrentAngle.Content = kinectSensorChooser1.Kinect.ElevationAngle
			End If

			System.Threading.Thread.Sleep(New TimeSpan(hours:= 0, minutes:= 0, seconds:= 1))
			button1.IsEnabled = True
		End Sub

		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			AddHandler kinectSensorChooser1.KinectSensorChanged, AddressOf kinectSensorChooser1_KinectSensorChanged
		End Sub

		Private Sub kinectSensorChooser1_KinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)
			Dim old = CType(e.OldValue, KinectSensor)
			Dim sensor = CType(e.NewValue, KinectSensor)

			sensor.Start()
			lblCurrentAngle.Content = kinectSensorChooser1.Kinect.ElevationAngle.ToString()

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

        Private Sub Window_Closing(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
            StopKinect(kinectSensorChooser1.Kinect)
        End Sub
    End Class
End Namespace
