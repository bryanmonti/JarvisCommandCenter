#region Using
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Kinect;
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
using System.Windows.Forms;
using System.Media;
#endregion
namespace Speech
{ //Pissing people off with #regions since Microsoft Visual 2008.
    public class Program
    {
        public static int tskid = 0;
        public static System.IO.Ports.SerialPort port;
        public static void Main(string[] args)
        {

            //string baud_string;
            string name;

            #region Kinect Finding
            // Obtain a KinectSensor if any are available
            KinectSensor sensor = (from sensorToCheck in KinectSensor.KinectSensors where sensorToCheck.Status == KinectStatus.Connected select sensorToCheck).FirstOrDefault();
            #endregion

            #region Kinect Checking
            if (sensor == null)
            {
                Console.WriteLine(
                        "No Kinect sensors are attached to this computer or none of the ones that are\n" +
                        "attached are \"Connected\".\n" +
                    //"Attach the KinectSensor and restart this application.\n" +
                    //"If that doesn't work run SkeletonViewer-WPF to better understand the Status of\n" +
                    //"the Kinect sensors.\n\n" +
                        "Press any key to continue.\n");

                // Give a chance for user to see console output before it is dismissed
                Console.ReadKey(true);
                return;
            }
            #endregion

            #region Port Checking + Counting
            System.IO.Ports.SerialPort.GetPortNames().Count(); //counts available ports (set this as a name somewhere)
            #endregion

            #region Activates Kinect Sensor
            sensor.Start();
            #endregion

            #region Obtains KinectAudioSource
            KinectAudioSource source = sensor.AudioSource;
            source.EchoCancellationMode = EchoCancellationMode.None; // No AEC :(
            source.AutomaticGainControlEnabled = false; // Important to turn this off for speech recognition
            #endregion

            #region Check for Audio SDK
            if (GetKinectRecognizer() == null)
            {
                Console.WriteLine("Could not find Kinect speech recognizer! You should probably install the Audio SDK for Kinect (released by Microsoft). Download here: http://www.microsoft.com/en-us/download/details.aspx?id=27226"); //Put a download link here to get the audio sdk --DONE
                return;
            }
            #endregion

            #region Writes Options
            Console.WriteLine("Enter your parameters to begin");
            Console.WriteLine(" ");
            Console.WriteLine("If no ports are displayed below, please check your connection to the serial device");
            Console.WriteLine("Available ports:");
            #endregion

            #region Available Port Printing
            if (System.IO.Ports.SerialPort.GetPortNames().Count() >= 0)
            {
                foreach (string p in System.IO.Ports.SerialPort.GetPortNames())
                {
                    Console.WriteLine(p);
                }
            }
            else
            {
                Console.WriteLine("No Ports are available. Press any key to quit!");
                Console.ReadLine();
                return; //Quit
            }
            #endregion

            #region Gets Port Name + Baud
            Console.WriteLine("Port Name:");
            name = Console.ReadLine();
            Console.WriteLine(" \n");
            Console.WriteLine("Baud rate:\n" +
                               "1. 300\n" +
                               "2. 1200\n" +
                               "3. 2400\n" +
                               "4. 4800\n" +
                               "5. 9600\n" +
                               "6. 14400\n" +
                               "7. 19200\n" +
                               "8. 28800\n" +
                               "9. 38400\n" +
                               "10. 57600\n" +
                               "11. 115200\n");
            /*baud_string = Console.ReadLine();
            int baud = int.Parse(baud_string); //Somewhat rigged*/
            //Console.WriteLine("You selected {0} as your baud rate\n", baud);
            int baud = GetBaudRate();

            Console.WriteLine(" ");
            Console.WriteLine("Beginning Serial...");
            BeginSerial(baud, name);
            port.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(port_DataReceived);
            port.Open();

            #endregion

            int wait = 5;
            while (wait > -1)//stops printing at 0 seconds
            {
                Console.Write("Device will be ready for speech recognition in {0} second(s).\r", wait--);//overwrite last printed statement
                Thread.Sleep(1000);
            }

            using (var sre = new SpeechRecognitionEngine(GetKinectRecognizer().Id))
            {
                var commands = new Choices();
                //POSSIBLY: Change commands and all of this shit(being the setup for it, if statments and all) to a simple XML file read my System.Xml.Linq; **http://stackoverflow.com/questions/10762715/c-sharp-parsing-specific-xml**
                //TODO: Reformat as well as organize this entire thing by subject, etc: Broswer stuff then inside broswre stuff is reddit, fb, gmail. Computer apps = steam, task manager
                #region commands
                commands.Add("pull up the weather");//DONE
                commands.Add("Open task manager");//DONE
                commands.Add("Ha gay");//DONE
                commands.Add("Play good feeling radio");//DONE
                commands.Add("Open reddit");//DONE
                commands.Add("Close chrome");//DONE
                commands.Add("Close task manager");//DONE
                commands.Add("Close reddit");
                commands.Add("Play pandora");//DONE
                commands.Add("Play good feeling radio");//DONE
                commands.Add("Play dead mouse radio");//DONE
                commands.Add("Boom");//DONE
                commands.Add("Sleep");//DONE-- Take a nap, Jarvis
                commands.Add("Shut down");//Program this
                commands.Add("Open my email");
                commands.Add("Open steam");//Also add launching games individually, Program this as well
                commands.Add("Open facebook");//Program this, also you may have to split the word
                commands.Add("battle stations");//Program this
                commands.Add("cable porn");//Program this
                commands.Add("ask science");//Program this
                commands.Add("github");//Program this (Make it open the CL and web interface)
                commands.Add("Raise Volume");//Setup and increment setting (Even in XML you can set the rate!)
                commands.Add("What are the commands available?");//Prints the commands that are available (Maybe explain in short each one) Program this
                commands.Add("What's new?");//Pull up some news and shit form reddit
                commands.Add("Hey Jarvis open a new tab");
                commands.Add("Sticky me");//Program this
                commands.Add("Next song Jarvis");
                commands.Add("Open microsoft visual");
                commands.Add("Open Webstorm");
                commands.Add("Open Pycharm");
                
                #endregion

                var gb = new GrammarBuilder { Culture = GetKinectRecognizer().Culture };                                
                gb.Append(commands);

                var g = new Grammar(gb);

                sre.LoadGrammar(g);
                sre.SpeechRecognized += SreSpeechRecognized;
                sre.SpeechHypothesized += SreSpeechHypothesized;
                sre.SpeechRecognitionRejected += SreSpeechRecognitionRejected;

                using (Stream s = source.Start())
                {
                    sre.SetInputToAudioStream(s, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));

                    Console.WriteLine(" ");
                    Console.WriteLine("What would you like me to do?\n");

                    sre.RecognizeAsync(RecognizeMode.Multiple);
                    Console.ReadLine();
                    Console.WriteLine("Stopping everything...\n");
                    sre.RecognizeAsyncStop();
                }
            }

            sensor.Stop();
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

        private static void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs audio)
        {
            Console.WriteLine("\nSpeech not recognized");
            if (audio.Result != null)
            {
                Console.WriteLine("In that Speech rejected block\n");
                DumpRecordedAudio(audio.Result.Audio);
            }
        }

        private static void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            Console.Write("\rSpeech Hypothesized: \t{0}", e.Result.Text);
        }

        private static void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {//TODO: Change all these if's into "else if" or setup a switch/case system (Faster processing is going to become key
            if (e.Result.Confidence >= 0.65)
            {
                Console.WriteLine("\nSpeech Recognized: \t{0}\tConfidence:\t{1}", e.Result.Text, e.Result.Confidence);
                //Add a line here that sends recieved audio to serial
                
                switch(e.Result.Text) {
                    case "Open reddit":
                        Process reddit = new Process();

                        reddit.StartInfo.FileName = "chrome.exe";
                        reddit.StartInfo.Arguments = "www.reddit.com";

                        reddit.Start();

                        port.WriteLine(e.Result.Text);
                        //Setup a variable and assign the process Id to it and use that Id to close that specific app
                        //or use MainWindowTitle (A little sketchy
                        //Session Id could work as well
                        Console.WriteLine("Sent your shit through serial\n");
                        break;

                    case "Close reddit":
                        Console.WriteLine("One does not simply close reddit!");
                        foreach (Process p in System.Diagnostics.Process.GetProcessesByName("reddit: the front pgae of the internet"))
                        {
                            try
                            {
                                p.Kill();
                                //add the iD to close the correct application instead of the entire browser
                                p.WaitForExit(); // possibly with a timeout
                            }
                            catch (Win32Exception winException)
                            {
                                // process was terminating or can't be terminated - deal with it
                            }
                            catch (InvalidOperationException invalidException)
                            {
                                // process has already exited - might be able to let this one go
                            }
                        }
                        break;

                    case "Close chrome":
                        foreach (Process p in System.Diagnostics.Process.GetProcessesByName("chrome"))
                        {
                            try
                            {
                                p.Kill();
                                //add the iD to close the correct application instead of the entire browser
                                p.WaitForExit(); // possibly with a timeout
                            }
                            catch (Win32Exception winException)
                            {
                                // process was terminating or can't be terminated - deal with it
                            }
                            catch (InvalidOperationException invalidException)
                            {
                                // process has already exited - might be able to let this one go
                            }
                        }
                        port.WriteLine(e.Result.Text);
                        Console.WriteLine("Calm your tits, I sent it.\n");
                        break;

                    case "pull up the weather":
                        Process weather = new Process();

                        weather.StartInfo.FileName = "chrome.exe";
                        weather.StartInfo.Arguments = "http://www.wunderground.com/cgi-bin/findweather/hdfForecast?query=11720";

                        weather.Start();
                        break;

                    case "Open task manager":
                        Process taskmanager = new Process();

                        SoundPlayer simpleSound = new SoundPlayer(@"c:/Windows/Media/chimes.wav");
                        simpleSound.Play();

                        taskmanager.StartInfo.FileName = "taskmgr.exe";
                        taskmanager.Close();
                        taskmanager.Start();
                        System.Diagnostics.Process.GetProcessById(tskid);
                        Console.WriteLine(tskid);
                        break;

                    case "Close task manager":
                        foreach (Process p in System.Diagnostics.Process.GetProcessesByName("taskmanager"))//Doesn't work with for eachj because not plural
                        {
                            try
                            {
                                p.Kill();
                                p.WaitForExit(); // possibly with a timeout
                            }
                            catch (Win32Exception winException)
                            {
                                // process was terminating or can't be terminated - deal with it
                            }
                            catch (InvalidOperationException invalidException)
                            {
                                // process has already exited - might be able to let this one go
                            }
                        }
                        break;

                    case "Play pandora":
                        Process pandora = new Process();

                        pandora.StartInfo.FileName = "chrome.exe";
                        pandora.StartInfo.Arguments = "www.pandora.com";

                        pandora.Start();
                        break;

                    case "Play dead mouse radio":
                        Process deadmau5 = new Process();

                        deadmau5.StartInfo.FileName = "chrome.exe";
                        deadmau5.StartInfo.Arguments = "http://www.pandora.com/station/play/680055046103061995";

                        deadmau5.Start();
                        break;

                    case "Play good feeling radio":
                        Process goodfeeling = new Process();

                        goodfeeling.StartInfo.FileName = "chrome.exe";
                        goodfeeling.StartInfo.Arguments = "http://www.pandora.com/station/play/705069957837182443";

                        goodfeeling.Start();
                        break;

                    case "Ha gay":
                        Process gaaay = new Process();

                        gaaay.StartInfo.FileName = "chrome.exe";
                        gaaay.StartInfo.Arguments = "www.hahgay.com";

                        gaaay.Start();
                        break;

                    case "Boom":
                        Process headshot = new Process();

                        headshot.StartInfo.FileName = "chrome.exe";
                        headshot.StartInfo.Arguments = "http://www.youtube.com/watch?v=F2FMDV8yW9M#t=0m56.3s";

                        headshot.Start();
                        break;

                    case "Sleep":
                        Application.SetSuspendState(PowerState.Suspend, true, true);
                        break;

                    case "Shut down":
                        Process.Start("shutdown", "/s /t 0");
                        break;

                    case "Open my email":
                        Process gmail = new Process();

                        gmail.StartInfo.FileName = "chrome.exe";
                        gmail.StartInfo.Arguments = "http://www.gmail.com";

                        gmail.Start();
                        break;

                    case "Open castle learning":
                        Process castle = new Process();

                        castle.StartInfo.FileName = "chrome.exe";
                        castle.StartInfo.Arguments = "http://www.castlelearning.com";

                        castle.Start();
                        break;
                }
                //so on and so fourth
            }
            else
            {
                Console.WriteLine("\nSpeech Recognized but confidence was too low: \t{0}", e.Result.Confidence);
                DumpRecordedAudio(e.Result.Audio); //deletes extra audio after being analyzed
            }
        }

        private static void DumpRecordedAudio(RecognizedAudio audio)
        {
            if (audio == null)
            {
                return;
            }

            int fileId = 0;
            string filename;
            while (File.Exists((filename = "RetainedAudio_" + fileId + ".wav")))
            {
                fileId++;
            }

            Console.WriteLine("\nWriting file: {0}", filename);
            using (var file = new FileStream(filename, System.IO.FileMode.CreateNew))
            {
                audio.WriteToWaveStream(file);
            }
        }

        static int GetBaudRate()
        {
            try
            {
                return int.Parse(Console.ReadLine());
            }
            catch
            {
                Console.WriteLine("Invalid integer.  Please try again:");
                return GetBaudRate();
            }
        }

        static void port_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            for (int i = 0; i < (10000 * port.BytesToRead) / port.BaudRate; i++)//not sure if this is needed (Check later)
                ;	 //Delay a bit for the serial to catch up
            Console.Write(port.ReadExisting());
            Console.WriteLine("");
            Console.WriteLine("> ");
        }

        static void BeginSerial(int baud, string name)
        {
            port = new SerialPort(name, baud);
        }
    }
}
