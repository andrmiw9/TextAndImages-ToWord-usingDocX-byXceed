using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace WorkWithWord2_2_graphical_DocX_lib_
{
    public partial class Entering : Form
    {
        static public int Templates = 0;
        public Entering()
        {
            InitializeComponent();
            ConsoleTB.Enabled = false;
        }
        private void button1_Click(object sender, EventArgs e)
        {
        int numt = int.Parse(NumberofT.Text);
            if (numt <= 0)                                  //error checking
            {
                ConsoleTB.Text += "Error: " + numt + " templates can not be added!\n";
                ConsoleTB.Update();
                Thread.Sleep(2000);
                Environment.Exit(0);
            }
            else
            {
                Templates = numt;
                this.Hide();
                Form1 t = new Form1(Templates);
                t.Show();
            }
        }
    }
}
