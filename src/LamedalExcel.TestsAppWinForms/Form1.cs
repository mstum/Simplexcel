﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LamedalExcel.TestsAppWinForms
{
    public sealed partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button_Excel_Click(object sender, EventArgs e)
        {
            var excel = new LamedalExcel_();
            excel.Excel_CreateFile();
        }
    }
}
