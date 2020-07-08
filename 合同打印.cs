using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 合同管理
{
    public partial class 合同打印 : Form
    {
        public string _url;
        public string _name;
        public 合同打印(string url, string name)
        {
            InitializeComponent();
            _url = url;
            _name = name;
        }

        private void 合同打印_Load(object sender, EventArgs e)
        {
            try {
                if(_url!="")
                {
                    this.Text = _name;
                    this.pdfViewer1.LoadFromFile(_url);
                }
              
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.ToString());
            }
          
        }

        private void 合同打印_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
