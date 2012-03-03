// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Public License (Ms-PL).
// Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Coding4Fun.Kinect.Wpf; 


namespace SettingUpDevEnvironment
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //this is a *template* app that includes library references only. It does not do anything but use SensorChooser
        //to manage Kinect state

        public MainWindow()
        {
            InitializeComponent();
        }

        ////Kinect 
        //KinectSensor _sensor;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //SetupKinectManually(); 
            kinectSensorChooser1.KinectSensorChanged +=new DependencyPropertyChangedEventHandler(kinectSensorChooser1_KinectSensorChanged);

        }


        //private void SetupKinectManually()
        //{
        //    if (KinectSensor.KinectSensors.Count > 0)
        //    {
        //        //use first Kinect
        //        _sensor = KinectSensor.KinectSensors[0];

        //        MessageBox.Show("Kinect Status = " + _sensor.Status.ToString()); 
        //    }
        //    else
        //    {
        //        MessageBox.Show("No Kinects are connected"); 
        //    }


        //}


        private void StopKinect(KinectSensor sensor)
        {
            if (sensor != null)
            {
                if (sensor.IsRunning)
                {
                    //stop sensor 
                    sensor.Stop();

                    //stop audio if not null
                    if (sensor.AudioSource != null)
                    {
                        sensor.AudioSource.Stop();
                    }
                    
                    
                }
            }
        }

        private void kinectSensorChooser1_KinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            KinectSensor oldSensor = (KinectSensor)e.OldValue;

            StopKinect(oldSensor); 

            KinectSensor newSensor = (KinectSensor)e.NewValue;

            if (newSensor == null)
            {
                return; 
            }

            //register for event and enable Kinect sensor features you want
            newSensor.AllFramesReady += new EventHandler<AllFramesReadyEventArgs>(newSensor_AllFramesReady);
            newSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            newSensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30); 
            newSensor.SkeletonStream.Enable();


            try
            {
                newSensor.Start(); 
            }
            catch (System.IO.IOException)
            {
                //another app is using Kinect
                kinectSensorChooser1.AppConflictOccurred();                 
            }



        }

        void newSensor_AllFramesReady(object sender, AllFramesReadyEventArgs e)
        {
            //put code here!
            //put code here!
            //put code here!
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopKinect(kinectSensorChooser1.Kinect); 
        }




    }
}

