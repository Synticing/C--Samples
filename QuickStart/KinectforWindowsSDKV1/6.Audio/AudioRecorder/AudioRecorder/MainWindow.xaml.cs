// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Public License (Ms-PL).
// Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
// All other rights reserved.
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Kinect;

namespace AudioRecorder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            FinishedRecording += new RoutedEventHandler(MainWindow_FinishedRecording);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            kinectSensorChooser1.KinectSensorChanged += new DependencyPropertyChangedEventHandler(kinectSensorChooser1_KinectSensorChanged);
        }

        void kinectSensorChooser1_KinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var old = (KinectSensor) e.OldValue;

            if (old != null)
            {
                old.Stop();
                old.AudioSource.Stop(); 
            }

            KinectSensor sensor = (KinectSensor)e.NewValue;
            sensor.Start(); 

        }

        double _amountOfTimeToRecord;
        string _lastRecordedFileName;
        private event RoutedEventHandler FinishedRecording;

        private void RecordButton_Click(object sender, RoutedEventArgs e)
        {
            RecordButton.IsEnabled = false;
            PlayButton.IsEnabled = false;
            _amountOfTimeToRecord = RecordForTimeSpan.Value; 
            _lastRecordedFileName = DateTime.Now.ToString("yyyyMMddHHmmss") + "_wav.wav";


            

            ////Start recording audio on new thread
            var t = new Thread(new ParameterizedThreadStart((RecordAudio)));
            t.Start(kinectSensorChooser1.Kinect); 

            //You can also Record audio synchronously but it will "freeze" the UI 
            //RecordAudio(kinectSensorChooser1.Kinect); 


        }

        private void PlayButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastRecordedFileName) && File.Exists(_lastRecordedFileName))
            {
                audioPlayer.Source = new Uri(_lastRecordedFileName, UriKind.RelativeOrAbsolute);
                audioPlayer.LoadedBehavior = MediaState.Play;
                audioPlayer.UnloadedBehavior = MediaState.Close;
            }
        }

        void MainWindow_FinishedRecording(object sender, RoutedEventArgs e)
        {
            //This is only required if recording on a separate thread to ensure that enabling the buttons
            //happens on the UI thread
            Dispatcher.BeginInvoke(new ThreadStart(ReenableButtons));

            //use this if recording on the same thread
            //ReenableButtons(); 
        }

        private void ReenableButtons()
        {
            RecordButton.IsEnabled = true;
            PlayButton.IsEnabled = true;
        }


        private void RecordAudio(object kinectSensor)
        {

            KinectSensor _sensor = (KinectSensor)kinectSensor;
            RecordAudio(_sensor); 
        }

        private void RecordAudio(KinectSensor kinectSensor)
        {

            if (kinectSensor == null)
            {
                return;
            }            
           
            int recordingLength = (int) _amountOfTimeToRecord * 2 * 16000;
            byte[] buffer = new byte[1024];                
                

            using (FileStream _fileStream  = new FileStream(_lastRecordedFileName, FileMode.Create))
            {
                WriteWavHeader(_fileStream, recordingLength);

                //Start capturing audio                               
                using (Stream audioStream = kinectSensor.AudioSource.Start())
                {
                    //Simply copy the data from the stream down to the file
                    int count, totalCount = 0;
                    while ((count = audioStream.Read(buffer, 0, buffer.Length)) > 0 && totalCount < recordingLength)
                    {
                        _fileStream.Write(buffer, 0, count);
                        totalCount += count;
                    }
                }
            }


            if (FinishedRecording != null)
            { 
                FinishedRecording(null, null);
            }
    }
        

        /// <summary>
        /// A bare bones WAV file header writer
        /// </summary>        
        static void WriteWavHeader(Stream stream, int dataLength)
        {
            //We need to use a memory stream because the BinaryWriter will close the underlying stream when it is closed
            using (var memStream = new MemoryStream(64))
            {
                int cbFormat = 18; //sizeof(WAVEFORMATEX)
                WAVEFORMATEX format = new WAVEFORMATEX()
                {
                    wFormatTag = 1,
                    nChannels = 1,
                    nSamplesPerSec = 16000,
                    nAvgBytesPerSec = 32000,
                    nBlockAlign = 2,
                    wBitsPerSample = 16,
                    cbSize = 0
                };

                using (var bw = new BinaryWriter(memStream))
                {
                    //RIFF header
                    WriteString(memStream, "RIFF");
                    bw.Write(dataLength + cbFormat + 4); //File size - 8
                    WriteString(memStream, "WAVE");
                    WriteString(memStream, "fmt ");
                    bw.Write(cbFormat);

                    //WAVEFORMATEX
                    bw.Write(format.wFormatTag);
                    bw.Write(format.nChannels);
                    bw.Write(format.nSamplesPerSec);
                    bw.Write(format.nAvgBytesPerSec);
                    bw.Write(format.nBlockAlign);
                    bw.Write(format.wBitsPerSample);
                    bw.Write(format.cbSize);

                    //data header
                    WriteString(memStream, "data");
                    bw.Write(dataLength);
                    memStream.WriteTo(stream);
                }
            }
        }

        static void WriteString(Stream stream, string s)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(s);
            stream.Write(bytes, 0, bytes.Length);
        }

        struct WAVEFORMATEX
        {
            public ushort wFormatTag;
            public ushort nChannels;
            public uint nSamplesPerSec;
            public uint nAvgBytesPerSec;
            public ushort nBlockAlign;
            public ushort wBitsPerSample;
            public ushort cbSize;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopKinect(kinectSensorChooser1.Kinect);
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

    }
}
