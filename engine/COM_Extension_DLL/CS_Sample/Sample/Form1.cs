using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Sample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
        }

        public string ShowForm()
        {
            monthCalendar1.MaxSelectionCount = 1;
            this.ShowDialog();
            return monthCalendar1.SelectionRange.Start.ToShortDateString();
        }
    }
}
