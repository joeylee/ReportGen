using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportGen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExReport report = new ExReport();
            report.read(@"D:\Documents\Weekly Activity Report\2011\TEST\TCS-H002_SYS1G2T_주간 업무 보고서(박준형)_111230.xls");
            report.close();
        }
    }
}
