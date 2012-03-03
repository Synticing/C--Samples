//------------------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

// This module provides sample code used to demonstrate the use
// of the KinectAudioSource for audio capture and beam tracking

namespace RecordAudio
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using Microsoft.Kinect;

    public class Program
    {
        private const int RiffHeaderSize = 20;
        private const string RiffHeaderTag = "RIFF";
        private const int WaveformatExSize = 18; // native sizeof(WAVEFORMATEX)
        private const int DataHeaderSize = 8;
        private const string DataHeaderTag = "data";
        private const int FullHeaderSize = RiffHeaderSize + WaveformatExSize + DataHeaderSize;

        public static void Main(string[] args)
        {
            var buffer = new byte[4096];
            int recordingLength = 0;
            const string OutputFileName = "out.wav";
            int numBeamChanged = 0;
            int numSoundSourceChanged = 0;
            
            // We need to run in high priority to avoid dropping samples 
            Thread.CurrentThread.Priority = ThreadPriority.Highest;

            // Obtain a KinectSensor if any are available
            KinectSensor sensor = (from sensorToCheck in KinectSensor.KinectSensors where sensorToCheck.Status == KinectStatus.Connected select sensorToCheck).FirstOrDefault();
            if (sensor == null)
            {
                Console.WriteLine(
                        "No Kinect sensors are attached to this computer or none of the ones that are\n" +
                        "attached are \"Connected\".\n" +
                        "Attach the KinectSensor and restart this application.\n" +
                        "If that doesn't work run SkeletonViewer-WPF to better understand the Status of\n" +
                        "the Kinect sensors.\n\n" +
                        "Press any key to continue.\n");

                // Give a chance for user to see console output before it is dismissed
                Console.ReadKey(true);
                return;
            }

            // Obtain the KinectAudioSource to do audio capture
            KinectAudioSource source = sensor.AudioSource;
            source.AutomaticGainControlEnabled = false;

            // We always need to start the sensor before we can start any sub-component
            try
            {
                sensor.Start();
            }
            catch (IOException)
            {
                Console.WriteLine(
                       "This application needs a Kinect for Windows sensor in order to function.\n" +
                       "However, another application is using the Kinect Sensor.\n\n" +
                       "Press any key to continue.\n");

                // Give a chance for user to see console output before it is dismissed
                Console.ReadKey(true);
                return;
            }

            // Register for beam tracking and sound source change notifications
            source.BeamAngleChanged += delegate { ++numBeamChanged; };
            source.SoundSourceAngleChanged += delegate { ++numSoundSourceChanged; };

            // NOTE: Need to wait 4 seconds for device to be ready right after initialization
            int wait = 4;
            while (wait > 0)
            {
                Console.Write("Device will be ready for recording in {0} second(s).\r", wait--);
                Thread.Sleep(1000);
            }

            Console.WriteLine("Device is ready. Press any key to start recording.");
            Console.ReadKey();

            using (var fileStream = new FileStream(OutputFileName, FileMode.Create))
            {
                // FXCop note: This try/finally block may look strange, but it is
                // the recommended way to correctly dispose a stream that is used
                // by a writer to avoid the stream from being double disposed.
                // For more information see FXCop rule: CA2202
                FileStream logStream = null;
                try
                {
                    logStream = new FileStream("samples.log", FileMode.Create);
                    using (var sampleStream = new StreamWriter(logStream))
                    {
                        logStream = null;
                        WriteWavHeader(fileStream);

                        WriteStatus(OutputFileName, source, numBeamChanged, numSoundSourceChanged);

                        // Start capturing audio                               
                        using (var audioStream = source.Start())
                        {
                            // Simply copy the data from the stream down to the file
                            int count;
                            bool readStream = true;
                            while (readStream && ((count = audioStream.Read(buffer, 0, buffer.Length)) > 0))
                            {
                                for (int i = 0; i < buffer.Length; i += 2)
                                {
                                    short sample = (short)(buffer[i] | (buffer[i + 1] << 8));
                                    sampleStream.WriteLine(sample);
                                }

                                fileStream.Write(buffer, 0, count);
                                recordingLength += count;

                                WriteStatus(OutputFileName, source, numBeamChanged, numSoundSourceChanged);

                                if (Console.KeyAvailable)
                                {
                                    ConsoleKeyInfo keyInfo = Console.ReadKey(true);
                                    switch (keyInfo.Key)
                                    {
                                        case ConsoleKey.S:
                                            readStream = false;
                                            break;

                                        case ConsoleKey.G:
                                            source.AutomaticGainControlEnabled = !source.AutomaticGainControlEnabled;
                                            break;
                                    }
                                }
                            }
                        }

                        UpdateDataLength(fileStream, recordingLength);
                    }
                }
                finally 
                {
                    if (logStream != null)
                    {
                        logStream.Dispose();
                    }
                }
            }

            // And Stop it
            sensor.Stop();
        }

        // This is the main console output that gets cleared and re-output in a loop to give the impression of a
        // status UI that gets updated in-place.
        private static void WriteStatus(string fileName, KinectAudioSource source, int numBeamChanged, int numSoundSourceChanged)
        {
            Console.Clear();
            Console.WriteLine("Recording to file {0}\nPress 's' to stop recording.\n", Path.GetFullPath(fileName));
            Console.WriteLine(
                "Automatic gain control is {0}. Press 'g' to {1}.\n",
                source.AutomaticGainControlEnabled ? "enabled" : "disabled",
                source.AutomaticGainControlEnabled ? "disable" : "enable");
            Console.WriteLine("Beam direction (degrees): {0:00.000}", source.BeamAngle);
            Console.WriteLine("\t\tChanged {0} times so far.\n", numBeamChanged);
            Console.WriteLine("Sound source direction (degrees): {0:00.000}\t\tConfidence: {1:0.000}", source.SoundSourceAngle, source.SoundSourceAngleConfidence);
            Console.WriteLine("\t\tChanged {0} times so far.", numSoundSourceChanged);
        }

        /// <summary>
        /// A bare bones WAV file header writer
        /// </summary>        
        private static void WriteWavHeader(Stream stream)
        {
            // Data length to be fixed up later
            int dataLength = 0;

            // We need to use a memory stream because the BinaryWriter will close the underlying stream when it is closed
            MemoryStream memStream = null;
            BinaryWriter bw = null;

            // FXCop note: This try/finally block may look strange, but it is
            // the recommended way to correctly dispose a stream that is used
            // by a writer to avoid the stream from being double disposed.
            // For more information see FXCop rule: CA2202
            try
            {
                memStream = new MemoryStream(64);

                WAVEFORMATEX format = new WAVEFORMATEX
                            {
                                FormatTag = 1,
                                Channels = 1,
                                SamplesPerSec = 16000,
                                AvgBytesPerSec = 32000,
                                BlockAlign = 2,
                                BitsPerSample = 16,
                                Size = 0
                            };

                bw = new BinaryWriter(memStream);

                // RIFF header
                WriteHeaderString(memStream, RiffHeaderTag);
                bw.Write(dataLength + FullHeaderSize - 8); // File size - 8
                WriteHeaderString(memStream, "WAVE");
                WriteHeaderString(memStream, "fmt ");
                bw.Write(WaveformatExSize);

                // WAVEFORMATEX
                bw.Write(format.FormatTag);
                bw.Write(format.Channels);
                bw.Write(format.SamplesPerSec);
                bw.Write(format.AvgBytesPerSec);
                bw.Write(format.BlockAlign);
                bw.Write(format.BitsPerSample);
                bw.Write(format.Size);

                // data header
                WriteHeaderString(memStream, DataHeaderTag);
                bw.Write(dataLength);
                memStream.WriteTo(stream);
            }
            finally
            {
                if (bw != null)
                {
                    memStream = null;
                    bw.Dispose();
                }

                if (memStream != null)
                {
                    memStream.Dispose();
                }
            }
        }

        private static void UpdateDataLength(Stream stream, int dataLength)
        {
            using (var bw = new BinaryWriter(stream))
            {
                // Write file size - 8 to riff header
                bw.Seek(RiffHeaderTag.Length, SeekOrigin.Begin);
                bw.Write(dataLength + FullHeaderSize - 8);

                // Write data size to data header
                bw.Seek(FullHeaderSize - 4, SeekOrigin.Begin);
                bw.Write(dataLength);
            }
        }

        private static void WriteHeaderString(Stream stream, string s)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(s);
            Debug.Assert(bytes.Length == s.Length, "The bytes and the string should be the same length.");
            stream.Write(bytes, 0, bytes.Length);
        }

        private struct WAVEFORMATEX
        {
            public ushort FormatTag;
            public ushort Channels;
            public uint SamplesPerSec;
            public uint AvgBytesPerSec;
            public ushort BlockAlign;
            public ushort BitsPerSample;
            public ushort Size;
        }
    }
}
