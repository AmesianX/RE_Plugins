using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace gleeGraph
{
    static class Program
    {

        //how is this better than optional arguments? this is bullshhiittttt
        public static string InputBox(string text) { return InputBox("", text, ""); }
        public static string InputBox(string text, string def){return InputBox("", text, def);}
        public static string InputBox(string title, string text, string def){
            return Interaction.InputBox(text, title, def, -1, -1);
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
