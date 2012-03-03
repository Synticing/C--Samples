' (c) Copyright Microsoft Corporation.
' This source is subject to the Microsoft Public License (Ms-PL).
' Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
' All other rights reserved.

Imports System.Text
Imports Microsoft.Kinect
Imports Coding4Fun.Kinect.Wpf


Namespace SettingUpDevEnvironment
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		'this is a *template* app that includes library references only. It does not do anything but use SensorChooser
		'to manage Kinect state

		Public Sub New()
			InitializeComponent()
		End Sub

		'//Kinect 
		'KinectSensor _sensor;

		Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			'SetupKinectManually(); 
			AddHandler kinectSensorChooser1.KinectSensorChanged, AddressOf kinectSensorChooser1_KinectSensorChanged

		End Sub


		'private void SetupKinectManually()
		'{
		'    if (KinectSensor.KinectSensors.Count > 0)
		'    {
		'        //use first Kinect
		'        _sensor = KinectSensor.KinectSensors[0];

		'        MessageBox.Show("Kinect Status = " + _sensor.Status.ToString()); 
		'    }
		'    else
		'    {
		'        MessageBox.Show("No Kinects are connected"); 
		'    }


		'}


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

		Private Sub kinectSensorChooser1_KinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)
			Dim oldSensor As KinectSensor = CType(e.OldValue, KinectSensor)

			StopKinect(oldSensor)

			Dim newSensor As KinectSensor = CType(e.NewValue, KinectSensor)

			If newSensor Is Nothing Then
				Return
			End If

			'register for event and enable Kinect sensor features you want
			AddHandler newSensor.AllFramesReady, AddressOf newSensor_AllFramesReady
			newSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30)
			newSensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30)
			newSensor.SkeletonStream.Enable()


			Try
				newSensor.Start()
			Catch e1 As System.IO.IOException
				'another app is using Kinect
				kinectSensorChooser1.AppConflictOccurred()
			End Try



		End Sub

		Private Sub newSensor_AllFramesReady(ByVal sender As Object, ByVal e As AllFramesReadyEventArgs)
			'put code here!
			'put code here!
			'put code here!
		End Sub

		Private Sub Window_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
			StopKinect(kinectSensorChooser1.Kinect)
		End Sub




	End Class
End Namespace

