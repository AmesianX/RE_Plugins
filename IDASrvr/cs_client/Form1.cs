using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace cs_client
{
    public partial class Form1 : Form
    {
        private ida_client ida = null;

        public Form1()
        {
            InitializeComponent();
        }

        protected override void WndProc(ref Message m)
        {
            if (ida == null)
            {
                base.WndProc(ref m);
            }
            else
            {
                if (!ida.HandleWindowProc(ref m)) base.WndProc(ref m);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ida = new ida_client(this.Handle);

            if (!ida.FindIDAHwnd())
            {
                listBox1.Items.Add("IDA Server window not found...");
                return;
            }

            listBox1.Items.Add("Listener hwnd: " + this.Handle );
            listBox1.Items.Add("IDA hwnd: " + ida.IDA_HWND );
            listBox1.Items.Add("File: " + ida.LoadedFileName() ) ;
            listBox1.Items.Add("#Funcs: " + ida.FuncCount());

            int fStart = ida.FuncStart(1);
            listBox1.Items.Add("Func[1] start: " + fStart.ToString("X") );
            listBox1.Items.Add("Func[1] end: " + ida.FuncEnd(1).ToString("X") );
            listBox1.Items.Add("Disasm @ 0x" + fStart.ToString("X") + ": " + ida.GetAsm(fStart));

 
          
        }
    }
}