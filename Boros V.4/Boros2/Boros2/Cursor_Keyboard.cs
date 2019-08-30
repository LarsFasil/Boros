using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

namespace Boros2
{
    class Cursor_Keyboard
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.Cdecl)]
        public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        public enum dirct { left, right, up, down };
        public enum EnumOptions { click, enter, back};


        public static void CursorMove(dirct x, int pxl)
        {
            Point p = new Point();
            switch (x)
            {
                case dirct.left:
                    p = new Point(Cursor.Position.X - pxl, Cursor.Position.Y);
                    break;
                case dirct.right:
                    p = new Point(Cursor.Position.X + pxl, Cursor.Position.Y);
                    break;
                case dirct.up:
                    p = new Point(Cursor.Position.X, Cursor.Position.Y - pxl);
                    break;
                case dirct.down:
                    p = new Point(Cursor.Position.X, Cursor.Position.Y + pxl);
                    break;
            }
            Cursor.Position = p;
        }

        public static void ckEvent(EnumOptions x)
        {
            switch (x)
            {
                case EnumOptions.click:
                    DoMouseClick();
                    break;
                case EnumOptions.enter:
                    SendKeys.Send("{ENTER}");
                    break;
                case EnumOptions.back:
                    //% = ALT, LEFT = left arrowkey
                    SendKeys.Send("%{LEFT}");
                    break;
            }
        }

        public static void Audio(int volume)
        {
            AudioManager.SetMasterVolume(AudioManager.GetMasterVolume() + volume);
        }
        public static void Audio(bool mute)
        { 
            AudioManager.ToggleMasterVolumeMute();
        }

        static void DoMouseClick()
        {
            //Call the imported function with the cursor's current position
            int X = Cursor.Position.X;
            int Y = Cursor.Position.Y;
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, X, Y, 0, 0);
        }
    }
}
