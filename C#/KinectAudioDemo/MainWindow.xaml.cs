﻿//------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace KinectAudioDemo
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Threading;
    using Microsoft.Kinect;
    using Microsoft.Speech.AudioFormat;
    using Microsoft.Speech.Recognition;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const double AngleChangeSmoothingFactor = 0.35;
        private const string AcceptedSpeechPrefix = "Accepted_";
        private const string RejectedSpeechPrefix = "Rejected_";

        private const int WaveImageWidth = 500;
        private const int WaveImageHeight = 100;

        private readonly SolidColorBrush redBrush = new SolidColorBrush(Colors.Red);
        private readonly SolidColorBrush greenBrush = new SolidColorBrush(Colors.Green);
        private readonly SolidColorBrush blueBrush = new SolidColorBrush(Colors.Blue);
        private readonly SolidColorBrush blackBrush = new SolidColorBrush(Colors.Black);

        private readonly WriteableBitmap bitmapWave;
        private readonly byte[] pixels;
        private readonly double[] energyBuffer = new double[WaveImageWidth];
        private readonly byte[] blackPixels = new byte[WaveImageWidth * WaveImageHeight];
        private readonly Int32Rect fullImageRect = new Int32Rect(0, 0, WaveImageWidth, WaveImageHeight);

        private KinectSensor kinect;
        private double angle;
        private bool running = true;
        private DispatcherTimer readyTimer;
        private EnergyCalculatingPassThroughStream stream;
        private SpeechRecognitionEngine speechRecognizer;

        public MainWindow()
        {
            InitializeComponent();

            var colorList = new List<Color> { Colors.Black, Colors.Green };
            this.bitmapWave = new WriteableBitmap(WaveImageWidth, WaveImageHeight, 96, 96, PixelFormats.Indexed1, new BitmapPalette(colorList));

            this.pixels = new byte[WaveImageWidth];
            for (int i = 0; i < this.pixels.Length; i++)
            {
                this.pixels[i] = 0xff;
            }

            imgWav.Source = this.bitmapWave;

            SensorChooser.KinectSensorChanged += this.SensorChooserKinectSensorChanged;
        }

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        private void SensorChooserKinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            KinectSensor oldSensor = e.OldValue as KinectSensor;
            if (oldSensor != null)
            {
                this.UninitializeKinect();
            }

            KinectSensor newSensor = e.NewValue as KinectSensor;
            this.kinect = newSensor;

            // Only enable this checkbox if we have a sensor
            enableAec.IsEnabled = this.kinect != null;

            if (newSensor != null)
            {
                this.InitializeKinect();
            }
        }

        private void InitializeKinect()
        {
            var sensor = this.kinect;
            this.speechRecognizer = this.CreateSpeechRecognizer();
            try
            {
                sensor.Start();
            }
            catch (Exception)
            {
                SensorChooser.AppConflictOccurred();
                return;
            }

            if (this.speechRecognizer != null && sensor != null)
            {
                // NOTE: Need to wait 4 seconds for device to be ready to stream audio right after initialization
                this.readyTimer = new DispatcherTimer();
                this.readyTimer.Tick += this.ReadyTimerTick;
                this.readyTimer.Interval = new TimeSpan(0, 0, 4);
                this.readyTimer.Start();

                this.ReportSpeechStatus("Initializing audio stream...");
                this.UpdateInstructionsText(string.Empty);
                
                this.Closing += this.MainWindowClosing;
            }

            this.running = true;
        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.Start();
            this.ReportSpeechStatus("Ready to recognize speech!");
            this.UpdateInstructionsText("Say: 'red', 'green' or 'blue'");
            this.readyTimer.Stop();
            this.readyTimer = null;
        }
        
        private void UninitializeKinect()
        {
            var sensor = this.kinect;
            this.running = false;
            if (this.speechRecognizer != null && sensor != null)
            {
                sensor.AudioSource.Stop();
                sensor.Stop();
                this.speechRecognizer.RecognizeAsyncCancel();
                this.speechRecognizer.RecognizeAsyncStop();
            }

            if (this.readyTimer != null)
            {
                this.readyTimer.Stop();
                this.readyTimer = null;
            }
        }

        private void MainWindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.UninitializeKinect();
        }

        private SpeechRecognitionEngine CreateSpeechRecognizer()
        {
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Close();
                return null;
            }

            SpeechRecognitionEngine sre;
            try
            {
                sre = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Close();
                return null;
            }

            var colors = new Choices();
            colors.Add("red");
            colors.Add("green");
            colors.Add("blue");

            var gb = new GrammarBuilder { Culture = ri.Culture };
            gb.Append(colors);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            sre.LoadGrammar(g);
            sre.SpeechRecognized += this.SreSpeechRecognized;
            sre.SpeechHypothesized += this.SreSpeechHypothesized;
            sre.SpeechRecognitionRejected += this.SreSpeechRecognitionRejected;

            return sre;
        }

        private void RejectSpeech(RecognitionResult result)
        {
            string status = "Rejected: " + (result == null ? string.Empty : result.Text + " " + result.Confidence);
            this.ReportSpeechStatus(status);

            Dispatcher.BeginInvoke(new Action(() => { tbColor.Background = blackBrush; }), DispatcherPriority.Normal);
        }

        private void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            this.RejectSpeech(e.Result);
        }

        private void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            this.ReportSpeechStatus("Hypothesized: " + e.Result.Text + " " + e.Result.Confidence);
        }

        private void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            SolidColorBrush brush;

            if (e.Result.Confidence < 0.7)
            {
                this.RejectSpeech(e.Result);
                return;
            }

            switch (e.Result.Text.ToUpperInvariant())
            {
                case "RED":
                    brush = this.redBrush;
                    break;
                case "GREEN":
                    brush = this.greenBrush;
                    break;
                case "BLUE":
                    brush = this.blueBrush;
                    break;
                default:
                    brush = this.blackBrush;
                    break;
            }

            string status = "Recognized: " + e.Result.Text + " " + e.Result.Confidence;
            this.ReportSpeechStatus(status);

            Dispatcher.BeginInvoke(new Action(() => { tbColor.Background = brush; }), DispatcherPriority.Normal);
        }

        private void ReportSpeechStatus(string status)
        {
            Dispatcher.BeginInvoke(new Action(() => { tbSpeechStatus.Text = status; }), DispatcherPriority.Normal);
        }

        private void UpdateInstructionsText(string instructions)
        {
            Dispatcher.BeginInvoke(new Action(() => { tbColor.Text = instructions; }), DispatcherPriority.Normal);
        }

        private void Start()
        {
            var audioSource = this.kinect.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();
            this.stream = new EnergyCalculatingPassThroughStream(kinectStream);
            this.speechRecognizer.SetInputToAudioStream(
                this.stream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            this.speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
            var t = new Thread(this.PollSoundSourceLocalization);
            t.Start();
        }

        private void EnableAecChecked(object sender, RoutedEventArgs e)
        {
            CheckBox enableAecCheckBox = (CheckBox)sender;
            if (enableAecCheckBox.IsChecked != null)
            {
                this.kinect.AudioSource.EchoCancellationMode = enableAecCheckBox.IsChecked.Value
                                                             ? EchoCancellationMode.CancellationAndSuppression
                                                             : EchoCancellationMode.None;
            }
        }

        private void PollSoundSourceLocalization()
        {
            while (this.running)
            {
                var audioSource = this.kinect.AudioSource;
                if (audioSource.SoundSourceAngleConfidence > 0.5)
                {
                    // Smooth the change in angle
                    double a = AngleChangeSmoothingFactor * audioSource.SoundSourceAngleConfidence;
                    this.angle = ((1 - a) * this.angle) + (a * audioSource.SoundSourceAngle);

                    Dispatcher.BeginInvoke(new Action(() => { rotTx.Angle = -angle; }), DispatcherPriority.Normal);
                }

                Dispatcher.BeginInvoke(
                    new Action(
                        () =>
                        {
                            clipConf.Rect = new Rect(
                                0, 0, 100 + (600 * audioSource.SoundSourceAngleConfidence), 50);
                            string sConf = string.Format("Conf: {0:0.00}", audioSource.SoundSourceAngleConfidence);
                            tbConf.Text = sConf;

                            stream.GetEnergy(energyBuffer);
                            this.bitmapWave.WritePixels(fullImageRect, blackPixels, WaveImageWidth, 0);

                            for (int i = 1; i < energyBuffer.Length; i++)
                            {
                                int energy = (int)(energyBuffer[i] * 5);
                                Int32Rect r = new Int32Rect(i, (WaveImageHeight / 2) - energy, 1, 2 * energy);
                                this.bitmapWave.WritePixels(r, pixels, 1, 0);
                            }
                        }),
                    DispatcherPriority.Normal);

                Thread.Sleep(50);
            }
        }

        private class EnergyCalculatingPassThroughStream : Stream
        {
            private const int SamplesPerPixel = 10;

            private readonly double[] energy = new double[WaveImageWidth];
            private readonly object syncRoot = new object();
            private readonly Stream baseStream;

            private int index;
            private int sampleCount;
            private double avgSample;

            public EnergyCalculatingPassThroughStream(Stream stream)
            {
                this.baseStream = stream;
            }

            public override long Length
            {
                get { return this.baseStream.Length; }
            }

            public override long Position
            {
                get { return this.baseStream.Position; }
                set { this.baseStream.Position = value; }
            }

            public override bool CanRead
            {
                get { return this.baseStream.CanRead; }
            }

            public override bool CanSeek
            {
                get { return this.baseStream.CanSeek; }
            }

            public override bool CanWrite
            {
                get { return this.baseStream.CanWrite; }
            }

            public override void Flush()
            {
                this.baseStream.Flush();
            }

            public void GetEnergy(double[] energyBuffer)
            {
                lock (this.syncRoot)
                {
                    int energyIndex = this.index;
                    for (int i = 0; i < this.energy.Length; i++)
                    {
                        energyBuffer[i] = this.energy[energyIndex];
                        energyIndex++;
                        if (energyIndex >= this.energy.Length)
                        {
                            energyIndex = 0;
                        }
                    }
                }
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                int retVal = this.baseStream.Read(buffer, offset, count);
                const double A = 0.3;
                lock (this.syncRoot)
                {
                    for (int i = 0; i < retVal; i += 2)
                    {
                        short sample = BitConverter.ToInt16(buffer, i + offset);
                        this.avgSample += sample * sample;
                        this.sampleCount++;

                        if (this.sampleCount == SamplesPerPixel)
                        {
                            this.avgSample /= SamplesPerPixel;

                            this.energy[this.index] = .2 + ((this.avgSample * 11) / (int.MaxValue / 2));
                            this.energy[this.index] = this.energy[this.index] > 10 ? 10 : this.energy[this.index];

                            if (this.index > 0)
                            {
                                this.energy[this.index] = (this.energy[this.index] * A) + ((1 - A) * this.energy[this.index - 1]);
                            }

                            this.index++;
                            if (this.index >= this.energy.Length)
                            {
                                this.index = 0;
                            }

                            this.avgSample = 0;
                            this.sampleCount = 0;
                        }
                    }
                }

                return retVal;
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                return this.baseStream.Seek(offset, origin);
            }

            public override void SetLength(long value)
            {
                this.baseStream.SetLength(value);
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                this.baseStream.Write(buffer, offset, count);
            }
        }
    }
}
