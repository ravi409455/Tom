using System;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;


namespace speech2017
{
    public partial class Tom : Form
    {
        private readonly SpeechSynthesizer s = new SpeechSynthesizer();
        private readonly SpeechRecognitionEngine reco = new SpeechRecognitionEngine();
        private readonly Choices list = new Choices();
        private Boolean mode = true;
        private readonly string _name = "ravi";
        public static string GreetRandom()
        {
            var greetings = new string[5] { "Hi sir", "Welcome back Ravi", "Nice to see you again Ravi", "What is the plan today Ravi ", "Good to see you again" };
            Random r = new Random();
            return greetings[r.Next(5)];

        }
        private static void Restart()
        {
            Process.Start(@"C:\Users\Ravi\Documents\Visual Studio 2017\Projects\speech2017\speech2017\bin\Debug\speech2017.exe");
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        public Tom()
        {
            InitializeComponent();
            s.SelectVoiceByHints(VoiceGender.Male);
            Sayit(GreetRandom());
           list.Add(File.ReadAllLines(@"C:\Users\Ravi\Documents\Visual Studio 2017\Projects\speech2017\speech2017\commands.txt"));
            Grammar gm = new Grammar(new GrammarBuilder(list));
            //DictationGrammar gm = new DictationGrammar();
            try
            {
                //serialPort1.Open();
                reco.RequestRecognizerUpdate();
                reco.LoadGrammar(gm);
                reco.SpeechRecognized += Reco_SpeechRecognized;
                reco.SetInputToDefaultAudioDevice();
                reco.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch
            {
                MessageBox.Show(@"Arduino is not connected. Some features might not work properly.");
            }


        }

        private void Sayit(string v)
        {
            s.SpeakAsync(v);
            textBox2.AppendText(v + "\n");
        }

        private void Reco_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string a = e.Result.Text;
            textBox1.AppendText(a + "\n");
            
            if (a == "sleep")
            {
                Sayit("OK sir");
                label2.Text = @"Mode : Sleep";
                mode = false;
            }
            if (a == "wake")
            {
                Sayit("Good Morning sir");
                label2.Text = @"Mode : Awake";
                mode = true;
            }

            //if (search == true)
            //{
            //    Process.Start("https://www.google.co.in/search?dcr=0&source=hp&ei=9Tg-Wo7bOoTsvgT55JjICw&q=" + a + "&oq=" + a + "&gs_l=psy-ab.3...2474.3785.0.3958.0.0.0.0.0.0.0.0..0.0....0...1c.1.64.psy-ab..0.0.0....0.4opvkPS0c_E");
            //    search = false;
            //}

            if (mode == true )// && search == false)
            {
                // if (a == "change my name")
                //{
                //  sayit("What is your name");
                //File.WriteAllText(@":\users\ravi\documents\visual studio 2015\Projects\speech\speech\name.txt", a);
                //}
                switch (a)
                {
                    case ("hi tom"):
                    case ("hello tom"):
                    case ("hello"):
                        Sayit("Hello Ravi");
                        break;

                    //case ("search for"):
                    //    search = true;
                    //    break;

                    case ("on kar"):
                        if (!serialPort1.IsOpen)
                        {
                            serialPort1.Open();
                        }
                        serialPort1.Write("1");
                        Sayit("Lights are turned On");
                        break;

                    case ("off kar"):
                        if (serialPort1.IsOpen)
                        {
                            serialPort1.Write("0");
                            serialPort1.Close();
                        }
                        Sayit("Lights are turned Off");
                        break;

                    case ("what is my name"):
                        Sayit("Your name is " + _name);
                        break;

                    case ("who is your father"):
                    case ("who created you"):
                        Sayit("I am created by Mister Ravi Kumar");
                        break;

                    case ("hi"):
                    case ("hi buddy"):
                    case ("hey"):
                        Sayit("Hey");
                        break;

                    case ("how are you tom"):
                    case ("how are you"):
                        Sayit("I am fine Sir, what about you");
                        break;

                    case ("who are you"):
                    case ("what is your name"):
                        Sayit("I am Tom, but there is no Jerry");
                        break;

                    case ("what do you do"):
                        Sayit("You can operate your computer and household appliances using me, just by giving your voice commands. You just say and I will do the task for you.");
                        break;

                    case ("show me your code"):
                        textBox2.AppendText("I am written using C# language. My code is this ");
                        s.Speak("I am written using CSharp language. My code is this ");
                        Process.Start(@"C:\Users\Ravi\Documents\Visual Studio 2017\Projects\speech2017\speech2017\tom's code.txt");
                        break;

                    case ("open folder"):
                        Sayit("Opening it");
                        Process.Start("explorer.exe");
                        break;

                    case ("open d drive"):
                        Sayit("Opening it");
                        Process.Start(@"D:\");
                        break;

                    case ("open google"):
                        Sayit("Opening Google");
                        Process.Start("https://www.google.co.in");
                        break;

                    case ("open calculator"):
                        Sayit("Opening Calculator");
                        Process.Start(@"D:\work-dev\Calci - Copy\index.html");
                        break;

                    case ("open notepad"):
                        Sayit("Opening Note pad");
                        Process.Start("notepad.exe");
                        break;

                    case ("open word"):
                        Sayit("Cool");
                        Process.Start(@"C:\Program Files\Microsoft Office\Office15\WINWORD.exe");
                        break;

                    case ("create a blank document"):
                        SendKeys.Send("^(n)");
                        break;

                    case ("open powerpoint"):
                        Sayit("Cool");
                        Process.Start(@"C:\Program Files\Microsoft Office\Office15\POWERPNT.exe");
                        break;

                    case ("open excel"):
                        Sayit("Cool");
                        Process.Start("C:\\Program Files\\Microsoft Office\\Office15\\EXCEL.exe");
                        break;

                    case ("open music"):
                        Sayit("Sure Sir");
                        Process.Start(@"C:\Users\Ravi\Desktop\Groove Music.lnk");
                        break;

                    case ("gaane chala"):
                        SendKeys.Send("^(p)");
                        break;

                    case ("play"):
                        SendKeys.Send("^(p)");
                        break;

                    case ("ruk"):
                        SendKeys.Send("^(p)");
                        break;

                    case ("pause"):
                        SendKeys.Send("^(p)");
                        break;

                    case ("next"):
                        SendKeys.Send("^(f)");
                        break;

                    case ("previous"):
                        SendKeys.Send("^(b)");
                        break;

                    case ("bye"):
                        textBox2.AppendText("Bye sir, Have a nice day");
                        s.Speak("Bye sir, Have a nice day");
                        
                        Application.Exit();
                        break;

                    case ("hide yourself"):
                        WindowState = FormWindowState.Minimized;
                        break;

                    case ("comeback"):
                        WindowState = FormWindowState.Normal;
                        break;

                    case ("fullscreen"):
                        WindowState = FormWindowState.Maximized;
                        break;

                    case ("close it"):
                        SendKeys.Send("%{F4}");
                        break;

                    case ("stop talking"):
                        s.SpeakAsyncCancelAll();
                        break;

                    //case ("open browser"):
                    //    sayit("Opening browser");
                    //    Form browser = new browser();
                    //    WindowState = FormWindowState.Minimized;
                    //    browser.Show();
                    //    break;

                    case ("do you love me"):
                        Sayit("Yes sir, I love you so much");
                        break;

                    case ("what is today"):
                        Sayit("Today is " + DateTime.Now.ToLongDateString());
                        break;

                    case ("what is the time"):
                        Sayit("Time is " + DateTime.Now.ToString(@"hh:mm tt"));
                        break;

                    case ("lock computer"):
                        Sayit("OK Sir");
                        Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");
                        break;

                    case ("restart"):
                        textBox2.AppendText("Restarting");
                        s.Speak("Restarting");
                        Restart();  
                        Application.Exit();
                        break;
                }
            }
        }

        private void Tom_Load(object sender, EventArgs e)
        {

        }
    }
}

