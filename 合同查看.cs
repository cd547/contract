﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 合同管理
{
    public partial class 合同查看 : Form
    {
        public string _url;
        public string _name;
        public 合同查看(string url, string name)
        {
            InitializeComponent();
            _url = url;
            _name = name;
        }

        private void 合同查看_Load(object sender, EventArgs e)
        {
            this.Text = _name;
            this.pdfDocumentViewer1.LoadFromFile(_url);
        }

        private void 合同查看_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
