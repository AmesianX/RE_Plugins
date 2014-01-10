using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Runtime.InteropServices;

namespace gleeGraph
{
    static class Program
    {
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        //how is this better than optional arguments? this is bullshhiittttt
        public static string InputBox(string text) { return InputBox("", text, ""); }
        public static string InputBox(string text, string def){return InputBox("", text, def);}
        public static string InputBox(string title, string text, string def){
            return Interaction.InputBox(text, title, def, -1, -1);
        }

        public static void SetTopMost(IntPtr hwnd, bool OnTop)
        {
            int HWND_TOPMOST = -1;
            int NOT_HWND_TOPMOST = -2;
            const UInt32 SWP_NOSIZE = 0x0001;
            const UInt32 SWP_NOMOVE = 0x0002;
            const UInt32 SWP_SHOWWINDOW = 0x0040;
            int flag = OnTop ? HWND_TOPMOST : NOT_HWND_TOPMOST;
            SetWindowPos(hwnd, flag, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
