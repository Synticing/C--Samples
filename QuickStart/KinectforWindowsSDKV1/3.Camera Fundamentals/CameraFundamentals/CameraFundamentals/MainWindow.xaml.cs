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
using System.Diagnostics;


namespace CameraFundamentals
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
      

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //sign up for the event
            kinectSensorChooser1.KinectSensorChanged += new DependencyPropertyChangedEventHandler(kinectSensorChooser1_KinectSensorChanged);

            
        }

        void kinectSensorChooser1_KinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {

            var oldSensor = (KinectSensor)e.OldValue;

            //stop the old sensor
            if (oldSensor != null)
            {
                StopKinect(oldSensor); 
            }

            //get the new sensor
            var newSensor = (KinectSensor)e.NewValue;
            if (newSensor == null)
            {
                return; 
            }

            //newSensor.AllFramesReady += new EventHandler<AllFramesReadyEventArgs>(newSensor_AllFramesReady);

            //turn on features that you need
            newSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);                
            newSensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30);
            newSensor.SkeletonStream.Enable();

            try
            {
                newSensor.Start();                
            }
            catch (System.IO.IOException)
            {
                //this happens if another app is using the Kinect
                kinectSensorChooser1.AppConflictOccurred();
            }
        }


        //this event fires when Color/Depth/Skeleton are synchronized
        void newSensor_AllFramesReady(object sender, AllFramesReadyEventArgs e)
        {
            //using (ColorImageFrame colorFrame = e.OpenColorImageFrame())
            //{
            //    if (colorFrame == null)
            //    {
            //        return;
            //    }

            //    byte[] pixels = new byte[colorFrame.PixelDataLength];
            //    colorFrame.CopyPixelDataTo(pixels);

            //    int stride = colorFrame.Width * 4;
            //    image1.Source =
            //        BitmapSource.Create(colorFrame.Width, colorFrame.Height,
            //        96, 96, PixelFormats.Bgr32, null, pixels, stride); 




            //}
        }


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

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopKinect(kinectSensorChooser1.Kinect); 
        } 
    }
}
