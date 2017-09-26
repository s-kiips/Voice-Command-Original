using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Net;
//using System.Net.Mail;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace VC
{
    public partial class Form2 : Form
    {
        RegistryKey reg = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
        SpeechRecognitionEngine rec = new SpeechRecognitionEngine();
        SpeechSynthesizer s = new SpeechSynthesizer();
        bool bb = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
        Boolean wake = true;
        string temp;
        string condition;
        Choices list = new Choices();
        private string count;
        string ProcWindow;
        Random rnd = new Random();
        [DllImport("user32")]
        public static extern void LockWorkStation();

        public Form2()
        {
            reg.SetValue("My App", Application.ExecutablePath.ToString());
            rec.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"D:\commands final.txt")))));
            //SpeechRecognitionEngine rec = new SpeechRecognitionEngine();
            list.Add(new string[] { "how are you" });
            
            Grammar gr = new Grammar(new GrammarBuilder(list));
            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeachRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch {
                return;
            }
            s.Speak("Welcome to Voice Command Application");
            InitializeComponent();
            //SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        }

       //------------------------------------XML TO ABSTRACT WEATHER--------------------------------------
        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='kathmandu')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");
            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;

                if (input == "temp")
                {
                    return temp;
                }

                if (input == "cond")
                {
                    return condition;
                }
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }

        //------------------------------------------- COMMANDS-------------------------------------------
        private void rec_SpeachRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string r = e.Result.Text;

            //-------------------------------------SLEEP AND WAKE--------------------------------------
            if (r == "wake")
            {
                Console.Beep();
                s.Speak("i am awake, what are your  commands");
                wake = true;
                label4.Text = "State : Awake";
                
            }
            if (r == "sleep")
            {
                Console.Beep();
                //Thread.Sleep(10);
                s.Speak("sleep mode activated");

                wake = false;

                label4.Text = "State : Sleep";

            }
            if (wake == true)
            {
                //-------------------------NORMAL CONVERSATION------------------------------------------------

                if (r == "hi")
                {
                    s.Speak("hello");
                }

                if (r == "hello AID" || r=="hello")
                {
                    DateTime timenow = DateTime.Now;
                    timenow = DateTime.Now;
                    if (timenow.Hour >= 5 && timenow.Hour < 12)
                    { s.Speak("Good morning how can i help you "); }
                    if (timenow.Hour >= 12 && timenow.Hour < 18)
                    { s.Speak("Good afternoon how may i help you"); }
                    if (timenow.Hour >= 18 && timenow.Hour < 24)
                    { s.Speak("Good evening how can i help you"); }
                    if (timenow.Hour < 5)
                    { s.Speak("hello, it's getting late"); }
                    //s.Speak("hello, how can i help you.");
                }

                if (r == "how are you")
                {
                    s.Speak("I'm good, thanks for asking");
                }

                if (r == "who are you")
                {
                    s.Speak("I am an A I, artificial intelligent being, A I D represents Assistant and Intelligent Device");
                }

                if (r == "what can you do")
                {
                    s.Speak("What would you like me to do?");
                }

                if (r == "who am i")
                {
                    s.Speak("I cant predict that, I'm still under development");
                }

                if (r == "sorry")
                {
                    s.Speak("It's ok no problem");
                }

                if (r == "thank you")
                {
                    s.Speak("You're welcome");
                }

                if (r == "is everything okay")
                {
                    s.Speak("Yes, why do you ask?");
                }

                if (r == "never mind")
                {
                    s.Speak("Alright then");
                }

                if (r == "tell me a joke")
                {
                    s.Speak("i dont know any, sorry");
                }
                               
                if (r == "your rude" || r== "rude" || r == "thats rude")
                {
                    s.Speak("Sorry, I didn't mean to hurt your feelings");
                }

                if (r == "thats mean" || r == "your mean")
                {
                    s.Speak("I was only joking");
                }

                
                if (r == "thats not funny" || r == "your not funny")
                {
                    s.Speak("neither are you");
                }

               
                if (r == "i need help")
                {
                    s.Speak("considering the fact that your talking to a computer about your problems I'd say you do need help");
                }

                if (r == "are you insulting me")
                {
                    s.Speak("NO, you misunderstood me.");
                }

                //--------------------------------------------------------NORMAL CONVERSATION END-------------------------------------

                //---------------------------------------------------------WEBSITES---------------------------------------------------
                if (r == "i want to search the web")
                {
                    s.Speak("opening the web now");
                    System.Diagnostics.Process.Start("chrome.exe");
                }
                if (r == "close chrome")
                {
                    s.Speak("closing chrome");
                    ProcWindow = "chrome";
                    StopWindow();
                }


                if (r == "open google")
                {
                    s.Speak("opening google in google chrome");
                    System.Diagnostics.Process.Start("http://google.com");

                }

                if (r == "open youtube")
                {
                    s.Speak("opening youtube");
                    System.Diagnostics.Process.Start("http://youtube.com");
                }

                if (r == "open facebook")
                {
                    s.Speak("opening facebook.");
                    System.Diagnostics.Process.Start("http://facebook.com");
                }

                if (r == "open gmail")
                {
                    s.Speak("opening your gmail.");
                    System.Diagnostics.Process.Start("http://gmail.com");
                }
                //--------------------------------------------------------------WEBSITES END--------------------------------------------


                
               //--------------------------------------------------------MINIMIZE-------------------------------------------------------
                if (r == "out of the way")
                {
                    if (WindowState == FormWindowState.Normal || WindowState == FormWindowState.Maximized)
                    {
                        WindowState = FormWindowState.Minimized;
                        s.Speak("My apologies");
                    }
                }

                if (r == "come back")
                {
                    if (this.WindowState == FormWindowState.Minimized)
                    {
                        s.Speak("Jojo lappa");
                        this.WindowState = FormWindowState.Normal;
                    }
                }

                if (r == "change window" || r == "switch window")
                {
                    s.Speak("ok");
                    SendKeys.Send("%{TAB " + count + "}");
                    count += 1;
                }

                
                if (r == "goodbye")
                {
                    if (WindowState == FormWindowState.Normal || WindowState == FormWindowState.Maximized)
                    {
                        WindowState = FormWindowState.Minimized;
                        s.Speak("ka wona ");
                    }
                }

                /* //-------------------------------------------------------------WINDOWS MAX-------------------------------------------
                if (r == "expand" || r == "enlarge" || r == "maxmize")
                 {
                     s.Speak("expanding");
                     FormBorderStyle = FormBorderStyle.None;
                     WindowState = FormWindowState.Maximized;
                     TopMost = true;

                 }
                 if (r == "exit fullscreen")
                 {
                     FormBorderStyle = FormBorderStyle.Sizable;
                     WindowState = FormWindowState.Normal;
                     TopMost = false;
                 }

                 //---------------------------------------------END OF WIMDOWs MAX-------------------------------------------*/

                //----------------------------------------------OPEN AND CLOSE APPLICATION------------------------------------
                if (r == "open explorer")
                {
                    System.Diagnostics.Process.Start(@"C:\Windows\explorer.exe");
                    s.Speak("Loading");
                }
                if (r == "close explorer")
                {
                    s.Speak("closing explorer");
                    foreach (Process proc in Process.GetProcessesByName("explorer"))
                    {

                        proc.Kill();
                    }                
                }
                if (r == "open excel")
                {
                    s.Speak("opening Excel");
                    Process.Start(@"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE");
                }

                if (r == "close excel")
                {
                    s.Speak("Closing excel");
                    ProcWindow = "excel";
                    StopWindow();
                }


                if (r == "open powerpoint")
                {
                    s.Speak("opening power point.");
                    System.Diagnostics.Process.Start(@"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE");
                }

                if (r == "close powerpoint")
                {
                    s.Speak("Closing powerpoint");
                    ProcWindow = "powerpnt";
                    StopWindow();
                }

                if (r == "open word")
                {
                    s.Speak("opening word.");
                    System.Diagnostics.Process.Start(@"C:\Program Files\Microsoft Office\Office16\winword.exe");
                }
                if (r == "close word")
                {
                    s.Speak("Closing word");
                    ProcWindow = "winword";
                    StopWindow();
                }

                if (r == "close chrome")
                {
                    s.Speak("closing chrome");
                    ProcWindow = "chrome";
                    StopWindow();
                }

                if (r == "open visual studio")
                {
                    s.Speak("opening Visual Studio");
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe");
                }


                if (r == "close visual studio")
                {
                    s.Speak("closing visual studio");
                    ProcWindow = "devenv";
                    StopWindow();
                }

                if (r == "open control panel")
                {
                    s.Speak("opening control panel");
                    System.Diagnostics.Process.Start(@"C:\Windows\System32\control.exe");
                }

                if (r == "open Calculator" || r == "calculator")
                {
                    s.Speak("opening calculator");
                    
                }
                if (r == "close calculator")
                {
                    ProcWindow = "calc";
                    StopWindow();
                }
                    if (r == "open notepad" || r == "i need to take notes")
                {
                    s.Speak("opening notepad");
                    System.Diagnostics.Process.Start(@"C:\Windows\notepad.exe");
                }
                if (r == "close notepad")
                {
                    s.Speak("closing notepad");
                    ProcWindow = "notepad";
                    StopWindow();
                }

                if (r == "open task manager")
                {
                    SendKeys.Send("^+{ESC}");
                }

                if (r == "open commands")
                {
                    s.Speak("Alright");
                    System.Diagnostics.Process.Start(@"D:\commands final.txt");
                }

                if (r == "open paint")
                {
                    s.Speak("opening paint");
                    Process.Start(@"C:\Windows\System32\mspaint.exe");
                }
                
                if (r == "close paint")
                {
                    s.Speak("closing paint");
                    ProcWindow = "mspaint";
                    StopWindow();
                }

                if (r == "my pictures")
                {
                    s.Speak("opening pictures");
                    Process.Start(@"C:\Users\ranji\Pictures");
                }

                //-------------------------------------------------------END OF OPEN CLOSE APP--------------------------------------------------------

                //-------------------------------------------------------MUSIC AND VIDEOS--------------------------------------------------
                if (r == "open video")
                {
                    s.Speak("opening your videos");
                    System.Diagnostics.Process.Start(@"C:\Users\ranji\Videos");
                }

                if (r == "enter")
                {
                    SendKeys.Send("{ENTER}");
                }


                if (r == "open vlc")
                {
                    s.Speak("opening vlc");
                    Process.Start(@"C:\Program Files (x86)\VideoLAN\VLC\vlc.EXE");
                }

                if (r == "close vlc")
                {
                    s.Speak("closing vlc");
                    ProcWindow = "vlc";
                    StopWindow();
                }
                                
                if (r == "next video")
                {
                    SendKeys.Send("{p}");
                }

                if (r == "last video")
                {
                    SendKeys.Send("{n}");
                }

                if (r == "play video" || r == "pause video")
                {
                    SendKeys.Send(" ");
                }

                if (r == "skip")
                {
                    SendKeys.Send("^{RIGHT}");
                }

                if (r == "video mute" || r == "unmute")
                {
                    SendKeys.Send("{m}");
                }

                if (r == "open player")
                {
                    s.Speak("opening windows media player");
                    Process.Start(@"C:\Program Files (x86)\Windows Media Player\wmplayer.EXE");
                }

                if (r == "play my playlist")
                {
                    s.Speak("opening your play list");
                    Process.Start(@"C:\Users\ranji\Music\Playlists\favourites.wpl");
                }

                if (r == "media close" || r == "close media player")
                {
                    s.Speak("clsoing media player");
                    ProcWindow = "wmplayer";
                    StopWindow();
            
                }

                if (r == "play this" || r == "pause")
                {
                    SendKeys.Send("^{p}");
                }

                if (r == "mute" || r == "sound")
                {
                    SendKeys.Send("{f7}");
                }

                if (r == "next")
                {
                    SendKeys.Send("^{f}");
                }
                if (r == "previous")
                {
                    SendKeys.Send("^{b}");
                }
                if (r == "shuffle")
                {
                    SendKeys.Send("^{h}");
                }
                if (r == "cancel shuffle")
                {
                    SendKeys.Send("^{h}");
                }
                if (r == "stop")
                {
                    SendKeys.Send("^{s}");
                }
                //------------------------------------------------------------END OF MUSIC AND VIDEO-----------------------------------------

                //-----------------------------------------------------------SELECT ALL, COPY, PASTE---------------------------------------
                if (r == "select all")
                {
                    SendKeys.Send("^{a}");
                }

                if (r == "copy")
                {
                    s.Speak("copy");
                    SendKeys.Send("^{c}");
                }

                if (r == "paste")
                {
                    SendKeys.Send("^{v}");
                }

                if (r == "cut this")
                {
                    SendKeys.Send("^{x}");
                }
                if (r == "delete this")
                {
                    SendKeys.Send("{delete}");
                 }
                //--------------------------------------------------------END OF COPY,PASTE,SELECT ALL-----------------------------------


                //-------------------------------------------------------EXIT CURRENT APPLICATION---------------------------------------
                if (r == "close program" || r== "close this")
                {
                    s.Speak("closing");
                    SendKeys.Send("%{F4}");
                }



               //-----------------------------------------------------DATE, TIME, WEATHER---------------------------------------

                if (r == "what time is it" || r == "whats the time")
                {
                    s.Speak("it is");
                    s.Speak(DateTime.Now.ToString("h:mm tt"));
                }
                if (r == "what day is today")
                {
                    s.Speak("today is" + DateTime.Today.ToString("dddd"));
                }
                    if( r == "whats the date" || r == "whats todays date")
                {
                    s.Speak("todays date is");
                    s.Speak(DateTime.Now.ToString("M/d/yyy"));
                }


                if (r == "whats the weather like")
                {
                    if (bb == true)
                    {
                        //MessageBox.Show("Internet connections are available");
                        s.Speak("the sky is" + GetWeather("cond"));
                    }
                    else
                    {
                        s.Speak("no internet connection");
                        //MessageBox.Show("Internet connections are not available");
                        //say("the sky is" + GetWeather("cond"));
                    }
                }

                if (r == "whats the temperature")
                {
                    if (bb == true)
                    {
                        s.Speak("it is" + GetWeather("temp") + "degree farenheit.");
                    }
                    else
                    {
                        s.Speak("no internet connection");
                    }
                }
                //----------------------------------------------------------END OF WEATHER,TIME,DATE-------------------------------

            if(r== "lock computer")
                {
                    s.Speak("locking computer");
                    System.Diagnostics.Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");
                }
            if(r== "shutdown computer")
                {
                    s.Speak("your computer will shutdown in one minute");
                    System.Diagnostics.Process.Start("Shutdown", "-s -t 60");
                }
                if(r== "abort shutdown")
                {
                    s.Speak("shutdown aborted");
                    System.Diagnostics.Process.Start("shutdown", "/a");
                }
                if(r== "restart computer")
                {
                    s.Speak("your computer will restart in one minute");
                    System.Diagnostics.Process.Start("Shutdown", "-r -t 60");
                }

            }
            textBox1.AppendText(r + "\n");
        }



        private void StopWindow()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName(ProcWindow);
            foreach (System.Diagnostics.Process proc in procs)
            {
                proc.CloseMainWindow();
            }
        }

        OpenFileDialog ofd = new OpenFileDialog();


        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            label1.Text = DateTime.Now.ToLongDateString();
            label2.Text = DateTime.Now.ToLongTimeString();

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongTimeString();
            timer1.Start();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"wmplayer.exe");
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"explorer.exe");
           
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.google.com/");
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
