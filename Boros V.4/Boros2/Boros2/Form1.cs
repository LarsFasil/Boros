﻿using System;
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
        new string[] nums = new string[100];

        new List<string> holder = new List<string>();

        Dictionary<string, string> pathToName = new Dictionary<string, string>();
        Dictionary<string, List<int>> programIDS = new Dictionary<string, List<int>>();
        bool hold = false;
        bool checkingSure, savePrev;
        int IntChoise;
        string prevCommand, toClose, cProgramName, path_Commands, path_Dict;

        Action<bool> method1;
        Action<Cursor_Keyboard.EnumOptions> method2;
        Action<int> method3;
        Action method4;

        bool param1;
        Cursor_Keyboard.EnumOptions param2;
        int param3;

        enum Mode { Normal, Cursor, Audio, Other };
        Mode mode;
        Mode prevMode;
        int pixelJump;
        int volumeJump;
        int wheelJump;
        bool selfMute;


        CSV csv = new CSV();


        public Form1()
        {

            InitializeComponent();
            InitializeVars();
            //Setup first Speech recognition engine
            NewSRE();

            Say("Hello, I am Boros");
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

            savePrev = false;
            IntChoise = 0;
            checkingSure = false;
            mode = Mode.Normal;
            pixelJump = 100;
            volumeJump = 20;
            wheelJump = 4;
            selfMute = false;
        }

        private void UpdDictionarys()
        {
            //Is dit nummers toevoegen?, moet beter
            for (int i = 0; i < nums.Length; i++)
            {
                nums[i] = (i).ToString();
                //listBox1.Items.Add(nums[i]);
            }
            //for (int i = 0; i < 10; i++)
            //{
            //    nums[i + (nums.Length - 10)] = "1";
            //    //listBox1.Items.Add(nums[i + (nums.Length - 10)]);
            //}

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
            modeSelect(e.Result.Text.ToString());
            if (savePrev)
            {
                prevCommand = e.Result.Text.ToString();
                savePrev = false;
            }
        }

        void modeSelect(string result)
        {
            if (result == "respond")
            {
                hold = false;
                Say("Yes?");
                listBox1.Items.Add("free to talk");
                return;
            }
            if (hold)
            {
                return;
            }
            else
            {
                listBox1.Items.Add(result);
            }
            if ((checkingSure == true || IntChoise > 0) && mode != Mode.Other)
            {
                prevMode = mode;
                mode = Mode.Other;
            }
            else if (CommonResults(result))
            {
                return;
            }

            switch (mode)
            {
                case Mode.Normal:
                    NormalMode(result);
                    break;
                case Mode.Cursor:
                    CursorMode(result);
                    break;
                case Mode.Audio:
                    AudioMode(result);
                    break;
                case Mode.Other:
                    OtherResults(result);
                    break;
            }
        }

        void NormalMode(string result)
        {
            if (result.Contains("open"))
            {
                foreach (KeyValuePair<string, string> kp in pathToName)
                {
                    if (result == kp.Value)
                    {
                        OpenSomething(kp.Key, kp.Value);
                        return;
                    }
                }
                Say("could not find that name in dictionary");
                return;
            }

            if (result.Contains("close"))
            {
                foreach (var kp in pathToName)
                {
                    if (result == kp.Value.Replace("open", "close"))
                    {
                        CloseSomething(kp.Key);
                        return;
                    }
                }
                Say("could not find that name in dictionary");
                return;
            }

            switch (result)
            {
                case "borrows":
                    Say("Yes?");
                    break;

                case "what time is it":
                    Say("The time is" + DateTime.Now.ToLongTimeString());
                    break;

                case "open notepad":
                    OpenSomething("notepad", "notepad");
                    break;
            }
        }
        void CursorMode(string result)
        {
            switch (result)
            {
                case "left":
                    Cursor_Keyboard.CursorMove(Cursor_Keyboard.dirct.left, pixelJump);
                    break;
                case "right":
                    Cursor_Keyboard.CursorMove(Cursor_Keyboard.dirct.right, pixelJump);
                    break;
                case "up":
                    Cursor_Keyboard.CursorMove(Cursor_Keyboard.dirct.up, pixelJump);
                    break;
                case "down":
                    Cursor_Keyboard.CursorMove(Cursor_Keyboard.dirct.down, pixelJump);
                    break;

                case "scroll up":
                    Cursor_Keyboard.MouseScroll(wheelJump, true);
                    break;
                case "scroll down":
                    Cursor_Keyboard.MouseScroll(wheelJump, false);
                    break;

                case "click":
                    Cursor_Keyboard.ckEvent(Cursor_Keyboard.EnumOptions.click);
                    break;
                case "back":
                    Confirm(Cursor_Keyboard.ckEvent, Cursor_Keyboard.EnumOptions.back);
                    break;
                case "enter":
                    Cursor_Keyboard.ckEvent(Cursor_Keyboard.EnumOptions.enter);
                    break;



                case "wheel jump":
                    SetupInt(100, ChangeWheel);
                    break;
                case "pixel jump":
                    SetupInt(100, ChangePixel);
                    break;
            }
        }
        void AudioMode(string result)
        {
            switch (result)
            {
                case "mute computer":
                    Cursor_Keyboard.Audio(true);
                    break;
                case "volume":
                    SetupInt(100, ChangeVolume);
                    break;
                case "up":
                    Cursor_Keyboard.Audio(volumeJump);
                    break;
                case "down":
                    Cursor_Keyboard.Audio(-volumeJump);
                    break;
                case "mute self":
                    selfMute = !selfMute;
                    break;
            }
        }
        void OtherResults(string result)
        {
            if (IntChoise > 0)
            {
                if (result == "nevermind")
                {
                    IntChoise = 0;
                    return;
                }
                for (int i = 0; i < IntChoise; i++)
                {
                    if (result == nums[i])
                    {
                        Say("you chose number " + nums[i]);
                        // voer actiet uit met i als parameter
                        IntChoiseAction(i);
                        IntChoise = 0;
                        mode = prevMode;
                        return;
                    }
                }
                Say("please choose a number from 0 to " + IntChoise.ToString()); // check of het woord in de dictionary staat
                return;
            }

            if (checkingSure)
            {
                if (result == "yes")
                {
                    if (method1 != null)
                    {
                        method1(param1);
                    }
                    if (method2 != null)
                    {
                        method2(param2);
                    }
                    if (method3 != null)
                    {
                        method3(param3);
                    }
                    if (method4 != null)
                    {
                        method4();
                    }
                    checkingSure = false;
                    mode = prevMode;
                    return;
                }
                if (result == "no")
                {
                    Say("ok, we will not " + prevCommand);

                    checkingSure = false;
                    mode = prevMode;
                    return;
                }
                Say("please say yes or no");
                Say("Would you like to " + prevCommand);
                return;
            }
        }
        bool CommonResults(string result)
        {
            bool x = true;
            switch (result)
            {
                case "clear log":
                    Say("Log cleared");
                    listBox1.Items.Clear();
                    break;
                case "show log":
                    Say("Showing Log");
                    this.WindowState = FormWindowState.Minimized;
                    this.WindowState = FormWindowState.Normal;
                    break;
                case "hide log":
                    Say("Hiding Log");
                    this.WindowState = FormWindowState.Minimized;
                    break;
                case "shut up":
                    Say("On Hold");
                    hold = true;
                    break;
                case "test":
                    test();
                    break;
                case "audio mode":
                    if (mode != Mode.Audio)
                    {
                        mode = Mode.Audio;
                        Say("Audio Mode on");
                    }
                    else
                    {
                        Say("Audio Mode is already on");
                    }
                    break;
                case "cursor mode":
                    if (mode != Mode.Cursor)
                    {
                        mode = Mode.Cursor;
                        Say("Cursor Mode on");
                    }
                    else
                    {
                        Say("Cursor Mode is already on");
                    }
                    break;
                case "normal mode":
                    if (mode != Mode.Normal)
                    {
                        mode = Mode.Normal;
                        Say("Normal Mode on");
                    }
                    else
                    {
                        Say("Normal Mode is already on");
                    }
                    break;
                case "exit":
                    Confirm(Application.Exit);
                    break;
                default:
                    x = false;
                    break;
            }
            return x;
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
            Say("Opening " + pName);

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

        #region confirm
        void Confirm(Action<bool> action, bool param)
        {
            resetMethods();
            method1 = action;
            param1 = param;
        }
        void Confirm(Action<Cursor_Keyboard.EnumOptions> action, Cursor_Keyboard.EnumOptions param)
        {
            resetMethods();
            method2 = action;
            param2 = param;
        }
        void Confirm(Action<int> action, int param)
        {
            resetMethods();
            method3 = action;
            param3 = param;
        }
        void Confirm(Action action)
        {
            resetMethods();
            method4 = action;
        }
        void resetMethods()
        {
            Say("are you sure you want to do that?");
            method1 = null;
            method2 = null;
            method3 = null;
            method4 = null;
            checkingSure = true;
            savePrev = true;
        }
        #endregion
        void ChangeWheel(int i)
        {
            wheelJump = i;
        }
        void ChangeVolume(int i)
        {
            volumeJump = i;
        }
        void ChangePixel(int i)
        {
            pixelJump = i * 10;
        }
        void IntChoiseAction(int i)
        {
            method3(i);
        }
        void SetupInt(int i, Action<int> action)
        {
            IntChoise = i;
            method3 = action;
            Say("choose a number between 0 and " + i.ToString());
        }

        void Say(string s)
        {
            if (!selfMute)
            {
                ss.SpeakAsync(s);
            }
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
                Say(pPath + " is not running at the moment");
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
                    Say("Witch of the " + programIDS[pPath].Count.ToString() + " would you like to close?");
                    cProgramName = pPath;
                    IntChoise = programIDS[pPath].Count;
                }
            }




            //foreach (Process p in Process.GetProcessesByName("explorer"))
            //{
            //    //if (p.MainWindowTitle.Contains(@path))
            //    listBox1.Items.Add(p.Id.ToString());

            //    //p.Kill();
            //}
        }
        #region windowDrag
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
        #endregion
    }
}
