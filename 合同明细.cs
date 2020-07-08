using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 合同管理
{
    public partial class 合同明细 : Form
    {
        public string _name;
        public 合同明细(string name)
        {
            InitializeComponent();
            _name = name;
        }


        public static bool IsMoney(string input)
        {
            string pattern = @"^\-{0,1}[0-9]{0,}\.{0,1}[0-9]{1,}$";
            return System.Text.RegularExpressions.Regex.IsMatch(input, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

            private void 合同明细_Load(object sender, EventArgs e)
        {
            this.label2.Text = _name;
        }


        static int contractreturned(string contract_name, string return_url, string return_time,decimal money, decimal money1)
        {
            int res = 0;
            int returnnum = 0;
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = "UPDATE contract_detail SET return=@return,returnurl=@returnurl,returntime=@returntime,money=@money,money1=@money1 WHERE contract_name='" + contract_name + "'";
                cmd.Parameters.Add("return", DbType.Int32).Value = 1;
                cmd.Parameters.Add("returnurl", DbType.String).Value = return_url;
                cmd.Parameters.Add("returntime", DbType.String).Value = return_time;
                cmd.Parameters.Add("money", DbType.Decimal).Value = money;
                cmd.Parameters.Add("money1", DbType.Decimal).Value = money1;
                returnnum = cmd.ExecuteNonQuery();
                //MessageBox.Show(cmd.CommandText);
            }
            cn.Close();
            return returnnum;
        }

        //合同归档
        private void button1_Click(object sender, EventArgs e)
        {
            if (IsMoney(this.textBox1.Text) && IsMoney(this.textBox2.Text))
            {
                MessageBox.Show("ok");
                DialogResult dr = openFileDialog1.ShowDialog();
                //获取所打开文件的文件名
                string openfilename = openFileDialog1.FileName;

                if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(openfilename))
                {
                    string pLocalFilePath = openfilename;//要复制的文件路径
                    string filename1 = System.IO.Path.GetFileName(pLocalFilePath);

                    string pSaveFilePath = "return/" + _name;//指定存储的路径
                                                                // MessageBox.Show(pSaveFilePath);
                    if (File.Exists(pLocalFilePath))//必须判断要复制的文件是否存在
                    {
                        File.Copy(pLocalFilePath, pSaveFilePath, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
                        DateTime dt = DateTime.Now;
                        contractreturned(_name, pSaveFilePath, dt.ToString(),Convert.ToDecimal(this.textBox1.Text), Convert.ToDecimal(this.textBox2.Text));
                       
                        //showcontractdetailtable();

                        //this.button11.Enabled = true;
                        //this.button12.Enabled = false;
                        //openurl = pSaveFilePath;
                    }

                }

            }
            else
            {
                MessageBox.Show("数字输入错误");
                return; 
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
