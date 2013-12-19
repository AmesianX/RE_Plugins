using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Glee.Drawing;
using System.Threading;
using System.IO;

namespace gleeGraph
{
    public partial class Form1 : Form
    {
        CGraph graph;
        Node selNode;
        ida_client ida = null;
        
        public void debugLog(string msg){
            lst.Items.Add(msg);
            lst.SelectedIndex = lst.Items.Count-1;
        }

        public Form1()
        {
            InitializeComponent();
            gViewer.ZoomFraction = .02; //zoom increment smaller for smooth scrolling..
            mnuPopup.MouseLeave += new EventHandler(mnuPopup_MouseLeave);
            graph = new CGraph(gViewer, lvNodes);
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
            gViewer.SelectionChanged += new EventHandler(gViewer_SelectionChanged);
            gViewer.MouseUp += new MouseEventHandler(gViewer_MouseUp);
            gViewer.MouseWheel += new MouseEventHandler(gViewer_MouseWheel);
            
            ida = new ida_client(this.Handle);

            if (!ida.FindIDAHwnd())
            {
                debugLog("IDA not found...");
            }
            else
            {
                debugLog("IDA hwnd: " + ida.IDA_HWND );
                debugLog("IDA File: " + Path.GetFileName(ida.LoadedFileName()));
            }

            
            string[] tmp = Environment.GetCommandLineArgs();
            string last = "c:\\lastGraph.txt";
            string f = "";

            if (false && System.Diagnostics.Debugger.IsAttached)
            {
                string testFile = Application.StartupPath + "\\test.txt";
                if (!File.Exists(testFile)) testFile = testFile.Replace("\\bin\\Debug", "");
                if (File.Exists(testFile)) graph.LoadFile(testFile);
            }
            else //load from command line or lastgraph if not..
            {
                for (int i = tmp.Length - 1; i > 0; i--)
                {
                    //MessageBox.Show(i + " " + tmp[i]);
                    if (File.Exists(tmp[i])) { f = tmp[i]; break; }
                }

                if (File.Exists(f))
                {
                    if (File.Exists(last)) File.Delete(last);
                    File.Copy(f, last);
                    debugLog("Loading " + f);
                    graph.LoadFile(f);
                }
                else if (File.Exists(last))
                {
                    debugLog("Loading last graph");
                    graph.LoadFile(last);
                }
            }
             

        }

        void mnuPopup_MouseLeave(object sender, EventArgs e)
        {
            mnuPopup.Hide();
        }

        void gViewer_MouseWheel(object sender, MouseEventArgs e)
        {
            if (e.Delta > 0) gViewer.ZoomInPressed(); else gViewer.ZoomOutPressed();
        }

        void gViewer_ZoomChanged(object sender, EventArgs e)
        {
            try { hScroll.Value = (int)gViewer.ZoomF; }
            catch (Exception ex) { }
        }

        void gViewer_SelectionChanged(object sender, EventArgs e)
        {
            gViewer_ZoomChanged(null, null);

            if(gViewer.SelectedObject == null)
            {
                if (selNode != null)
                {
                    selNode.Attr.LineWidth = 1;
                    selNode = null;     
                }
            }
            else if(gViewer.SelectedObject is Node)
            {
                if (selNode != null)
                {
                    //you could rig multiselect here by testing ctrl key and latter using linewidth as criteria..
                    selNode.Attr.LineWidth = 1;
                }
                selNode = (Node)gViewer.SelectedObject;
                selNode.Attr.LineWidth = 3;
                gViewer.Refresh();
                //debugLog("Selected node is " + selNode.Attr.Label.Trim());
            }

            gViewer.Refresh();
        }

        void gViewer_MouseUp(object sender, MouseEventArgs e)
        {
            //have to hot track with selectionChanged event or else select wont process in time for MouseUp...
            //a right click can even beat the selectionChanged event, so mouse over watch highlight, then rightclick :-\
            //must be a thread.invoke queing delay causing the problem in the glee control. 
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (selNode != null) mnuPopup.Show(Cursor.Position);
            }
            else
            {
                if (selNode != null) ida.jmpName(selNode.Attr.Label);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            try
            {
                gViewer.Width = this.Width - gViewer.Left - 20;
                gViewer.Height = this.Height - gViewer.Top - 40;
            }
            catch (Exception ex) { }
        }

        private void zoomAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gViewer.ShowBBox(gViewer.Graph.BBox);
        }

        private void lvNodes_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {

            if (selNode != null) selNode.Attr.LineWidth = 1;
            selNode = (Node)e.Item.Tag;
            selNode.Attr.LineWidth = 3;
            ida.jmpName(selNode.Attr.Label);
            //gViewer.ShowBBox(selNode.BBox); //zoom to node..
            gViewer.Invalidate();
        }

        private void loadGraphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(dlg.ShowDialog() != DialogResult.OK) return;
            debugLog("Loading " + Path.GetFileName(dlg.FileName));
            graph.LoadFile(dlg.FileName);
        }

        private void lst_DoubleClick(object sender, EventArgs e)
        {
            if (lst.SelectedIndex > 0)
            {
                MessageBox.Show(lst.SelectedItem.ToString());
            }
        }

        private void renameNodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(selNode == null) return;
            string oldName = selNode.Attr.Label.Trim();
            string newName = Program.InputBox("Enter new name", oldName);
            if (newName.Length == 0) return;
            if (ida.Rename(oldName, newName))
            {
                selNode.Attr.Label = "   " + newName + "   ";
                ListViewItem li = (ListViewItem)selNode.UserData;
                li.Text = newName;
            }
            else
            {
                debugLog("Fail rename " + oldName);
            }
        }

        private void removeNodesBelowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (selNode == null) return;
            List<Node> nodes = graph.NodesBelow(selNode);
            foreach (Node n in nodes)
            {
                n.Attr.AddStyle(Style.Invis); //doesnt look like there is any way to hide nodes..
            }
            gViewer.Invalidate();
        }

        private void prefixAllFunctionsBelowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (selNode == null) return;

            string prefix = Program.InputBox("Enter prefix to use");
            if (prefix.Length == 0) return;
            
            List<Node> nodes = graph.NodesBelow(selNode);
            foreach (Node n in nodes)
            {
                string newName = prefix + n.Attr.Label.Trim();
                if (ida.Rename(n.Attr.Label, newName ))
                {
                    n.Attr.Label = "   " + newName + "   ";
                    ListViewItem li = (ListViewItem)n.UserData;
                    li.Text = newName;
                }
                else
                {
                    debugLog("Fail rename " + n.Attr.Label.Trim() );
                }
            }
            gViewer.Invalidate();
            lvNodes.Refresh();
        }

        private void hScroll_Scroll(object sender, ScrollEventArgs e)
        {
            try
            {
                gViewer.ZoomF = hScroll.Value;
            }
            catch (Exception ex) { }

        }


    }
}
