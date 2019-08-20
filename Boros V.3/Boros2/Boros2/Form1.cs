using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Speech;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Threading;
using Shell32;

namespace Boros2
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer ss = new SpeechSynthesizer();
        PromptBuilder pb = new PromptBuilder();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();
        Choices clist = new Choices();
        new string[] nums = new string[30];

        new List<string> holder = new List<string>();

        Dictionary<string, string> pathToName = new Dictionary<string, string>();
        Dictionary<string, List<int>> programIDS = new Dictionary<string, List<int>>();
        bool hold = false;
        bool checkingSure = false;
        int choises = 0;
        string question, toClose, cProgramName, path_Commands, path_Dict;
        Action<string> method;
        int pixelJump;

      
        CSV csv = new CSV();
        Cursor_Keyboard ck = new Cursor_Keyboard();



        public Form1()
        {
            InitializeComponent();
            InitializeVars();
            //Setup first Speech recognition engine
            NewSRE();

            ss.SpeakAsync("Hello, I am Boros");
            //MessageBox.Show(csv.getData());

        }

        private void NewSRE()
        {
            Grammar gr = new Grammar(new GrammarBuilder(clist));
            try
            {
                sre.RequestRecognizerUpdate();
                sre.LoadGrammar(gr);
                sre.SpeechRecognized += Sre_SpeechRecognized;
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void InitializeVars()
        {
            path_Commands = @Directory.GetCurrentDirectory() + "\\Commands.csv";
            path_Dict = @Directory.GetCurrentDirectory() + "\\Dictionary.csv";

            ProcessDirectory(@"C:\Users\" + Environment.UserName + "\\Desktop");
            UpdDictionarys();

            pixelJump = 80;
        }

        private void UpdDictionarys()
        {
            //Is dit nummers toevoegen?, moet beter
            for (int i = 0; i < nums.Length - 10; i++)
            {
                nums[i] = (i + 1).ToString();
                //listBox1.Items.Add(nums[i]);
            }
            for (int i = 0; i < 10; i++)
            {
                nums[i + (nums.Length - 10)] = "1";
                //listBox1.Items.Add(nums[i + (nums.Length - 10)]);
            }

            clist.Add(nums);
            //clist.Add(holder.ToArray());
            clist.Add(csv.getData(path_Commands));
            clist.Add(csv.getData(path_Dict));
            //clist.Add(new string[] { "hello", "how are you", "open chrome", "open music", "what time is it", //"close chrome",
            //"close music", "exit", "shut up please", "563", "show log", "hide log", "yes", "no", "borrows respond", "clear log", "open notepad","close notepad","test", "open word"});
        }

        private void ProcessDirectory(string dirPath)
        {
            string[] fileEntries = Directory.GetFiles(dirPath);
            //string[] files = new string[fileEntries.Count]; // hoeft niet

            //fix assap
            string[] sa_data = new string[fileEntries.Length];
            int z = 0;
            foreach (string filePath in fileEntries)
            {

                sa_data[z] = "open " + Path.GetFileName(filePath).ToLower().Replace(".exe", "").Replace(".url", "").Replace(".lnk", "").Replace(".txt", "");
                //data.AppendLine("open " + Path.GetFileName(filePath).ToLower().Replace(".exe", "").Replace(".url", "").Replace(".lnk", "").Replace(".txt", ""));
                z++;
            }


            for (int i = 0; i < fileEntries.Length; i++)
            {
                pathToName.Add(fileEntries[i], sa_data[i]); //maybe later open voor zetten
            }

            foreach (KeyValuePair<string, string> kp in pathToName)
            {
                listBox1.Items.Add(kp.Value.Replace("open ", ""));
            }

            csv.UpdateData(sa_data, path_Dict);
        }

        private void Sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (hold)
            {
                if (e.Result.Text.ToString() == "borrows respond")
                {
                    hold = false;
                    ss.SpeakAsync("Yes?");
                    listBox1.Items.Add("free to talk");
                }
            }
            else
            {
                if (choises != 0)
                {
                    for (int i = 0; i < choises; i++)
                    {
                        if (e.Result.Text.ToString() == nums[i])
                        {
                            ss.SpeakAsync("you chose number " + nums[i]);
                            CloseChoise(i);
                            choises = 0;
                        }
                        //else
                        //{
                        //    ss.SpeakAsync("please choose a number from 0 to " + choises.ToString()); // check of het woord in de dictionary staat
                        //}
                    }
                }
                else
                {
                    if (checkingSure)
                    {

                        if (e.Result.Text.ToString() == "yes")
                        {
                            ss.SpeakAsync(question);
                            method(toClose);
                        }
                        else { ss.SpeakAsync("Ok"); }

                        checkingSure = false;
                    }
                    else
                    {
                        if (e.Result.Text.ToString().Contains("open"))
                        {
                            listBox1.Items.Add("open caught");
                            foreach (KeyValuePair<string, string> kp in pathToName)
                            {
                                if (e.Result.Text.ToString() == kp.Value)
                                {
                                    OpenSomething(kp.Key, kp.Value);
                                }
                            }
                        }
                        else
                        {
                            if (e.Result.Text.ToString().Contains("close"))
                            {
                                foreach (var kp in pathToName)
                                {
                                    if (e.Result.Text.ToString() == kp.Value.Replace("open", "close"))
                                    {
                                        CloseSomething(kp.Key);
                                    }
                                }


                            }
                            else
                            {

                                switch (e.Result.Text.ToString())
                                {

                                    case "borrows":
                                        ss.SpeakAsync("Yes?");
                                        break;
                                    case "21":
                                        ss.SpeakAsync("Hi lars");
                                        break;
                                    case "how are you":
                                        ss.SpeakAsync("fine how are you?");
                                        break;
                                    case "what time is it":
                                        ss.SpeakAsync("The time is" + DateTime.Now.ToLongTimeString());
                                        break;
                                    //case "open chrome":
                                    //    //ss.SpeakAsync("Starting Chrome");
                                    //    //Process.Start("chrome");
                                    //    OpenSomething("chrome");
                                    //    break;
                                    //case "close chrome":

                                    //    ss.Speak("Are you sure?");
                                    //    question = "Closing all windows off Chrome";

                                    //    method = CloseSomething;
                                    //    toClose = "chrome";
                                    //    checkingSure = true;


                                    //    break;
                                    //case "open music":
                                    //    //ss.SpeakAsync("Opening music");
                                    //    //Process.Start(@"E:\4.muziek\yt");
                                    //    OpenSomething(@"E:\4.muziek\yt");
                                    //    break;
                                    //case "close music":
                                    //    ss.SpeakAsync("Closing music");
                                    //    CloseSomething(@"E:\4.muziek\yt");
                                    //    break;


                                    case "open notepad":
                                        OpenSomething("notepad", "notepad");
                                        break;
                                    //case "close notepad":
                                    //    CloseSomething("notepad");
                                    //    break;

                                    case "test":
                                        test();
                                        break;

                                    case "clear log":
                                        ss.SpeakAsync("Log cleared");
                                        listBox1.Items.Clear();
                                        break;
                                    case "show log":
                                        ss.SpeakAsync("Showing Log");
                                        this.WindowState = FormWindowState.Minimized;
                                        this.WindowState = FormWindowState.Normal;
                                        break;
                                    case "hide log":
                                        ss.SpeakAsync("Hiding Log");
                                        this.WindowState = FormWindowState.Minimized;
                                        break;
                                    case "shut up please":
                                        ss.SpeakAsync("On Hold");
                                        hold = true;
                                        break;

                                    case "left":
                                        ck.CursorMove(Cursor_Keyboard.dirct.left, pixelJump);
                                        break;
                                    case "right":
                                        ck.CursorMove(Cursor_Keyboard.dirct.right, pixelJump);
                                        break;
                                    case "up":
                                        ck.CursorMove(Cursor_Keyboard.dirct.up, pixelJump);
                                        break;
                                    case "down":
                                        ck.CursorMove(Cursor_Keyboard.dirct.down, pixelJump);
                                        break;

                                    case "click":
                                        ck.ckEvent(Cursor_Keyboard.ck_event.click);
                                        break;
                                    case "back":
                                        ck.ckEvent(Cursor_Keyboard.ck_event.back);
                                        break;
                                    case "enter":
                                        ck.ckEvent(Cursor_Keyboard.ck_event.enter);
                                        break;

                                    case "exit":
                                        //csv.EndStream();
                                        Application.Exit();
                                        break;
                                }
                            }
                        }


                    }
                }






                listBox1.Items.Add(e.Result.Text.ToString());
            }

        }
        public static string GetShortcutTargetFile(string shortcutFilename)
        {
            string pathOnly = System.IO.Path.GetDirectoryName(shortcutFilename);
            string filenameOnly = System.IO.Path.GetFileName(shortcutFilename);

            Shell shell = new Shell();
            Folder folder = shell.NameSpace(pathOnly);
            FolderItem folderItem = folder.ParseName(filenameOnly);
            if (folderItem != null)
            {
                Shell32.ShellLinkObject link = (Shell32.ShellLinkObject)folderItem.GetLink;
                return link.Path;
            }

            return string.Empty;
        }


        void test()
        {

        }

        bool drag = false;
        int Mx, My;

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            drag = true;
            Mx = Cursor.Position.X - this.Left;
            My = Cursor.Position.Y - this.Top;
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (drag)
            {
                Top = Cursor.Position.Y - My;
                Left = Cursor.Position.X - Mx;
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            drag = false;
        }



        void OpenSomething(string pPath, string pName)
        {
            Process p;
            if (pPath.Contains(".lnk"))
            {
                string RealPath = GetShortcutTargetFile(pPath);
                if (RealPath.Contains(" (x86)"))
                {
                    RealPath = RealPath.Replace(" (x86)", "");
                }
                p = Process.Start(RealPath);
            }
            else
            {
                p = Process.Start(pPath);
            }


            listBox1.Items.Add(pPath);
            pName = pName.Replace("open ", "");
            ss.SpeakAsync("Opening " + pName);

            clist.Add("close " + pName);
            sre.Dispose();
            sre = new SpeechRecognitionEngine();
            NewSRE();

            if (programIDS.ContainsKey(pPath))
            {

                programIDS[pPath].Add(p.Id);
            }
            else
            {

                programIDS.Add(pPath, new List<int> { p.Id });
            }


            //============================================================================
            //SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindows();

            //string filename;
            //ArrayList windows = new ArrayList();


            //foreach (SHDocVw.InternetExplorer ie in shellWindows)
            //{
            //    filename = Path.GetFileNameWithoutExtension(ie.FullName).ToLower();
            //    if (filename.Equals("explorer"))
            //    {
            //        //do something with the handle here
            //        MessageBox.Show(ie.HWND.ToString());
            //    }
            //}
        }


        void CloseChrome()
        {
            Process[] chromeInstances = Process.GetProcessesByName("chrome");

            foreach (Process p in chromeInstances)
                //p.Kill();
                listBox1.Items.Add(p.Handle);
        }

        private void CloseChoise(int choise)
        {
            Process p = Process.GetProcessById(programIDS[cProgramName][choise]);
            p.Kill();
            programIDS[cProgramName].Remove(programIDS[cProgramName][choise]);
            //listBox1.Items.Add(programIDS[cProgramName].Count);
        }

        void CloseSomething(string pPath)
        {

            if (programIDS.ContainsKey(pPath) == false || programIDS[pPath].Count == 0)
            {
                ss.SpeakAsync(pPath + " is not running at the moment");
            }
            else
            {
                if (programIDS[pPath].Count == 1)
                {
                    Process p = Process.GetProcessById(programIDS[pPath][0]);
                    p.Kill();
                    programIDS.Remove(pPath);
                }
                else
                {
                    ss.SpeakAsync("Witch of the " + programIDS[pPath].Count.ToString() + " would you like to close?");
                    cProgramName = pPath;
                    choises = programIDS[pPath].Count;
                }
            }




            //foreach (Process p in Process.GetProcessesByName("explorer"))
            //{
            //    //if (p.MainWindowTitle.Contains(@path))
            //    listBox1.Items.Add(p.Id.ToString());

            //    //p.Kill();
            //}
        }


    }
}
