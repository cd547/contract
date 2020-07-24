using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Runtime.CompilerServices;
using System.IO;
using System.Text.RegularExpressions;

namespace 合同管理
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //backup();
            showcontractdetailtable("");
            showcontractname();
            panel1.Visible = false;
            this.button13.Enabled = false;
            this.button18.Enabled = false;

        }

        public void backup(string events)
        {
            if (File.Exists(@"contract.db"))
            {
                MessageBox.Show(events);
                string NewFileName = @"dbback\contract_"+ events+"_" + DateTime.Now.ToString("yyyyMMddHHmmss").ToString()+".db";
                   
                   File.Copy(@"contract.db", NewFileName);
            }
        }

        static void CreateDB()
        {
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            cn.Open();
            cn.Close();
        }

        static void CreateTable(string sqlstr)
        {
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = sqlstr;//"CREATE TABLE contract( contract_name TEXT NOT NULL,lastnumber INT ); ";
                //cmd.CommandText = "CREATE TABLE IF NOT EXISTS t1(id varchar(4),score int)";
                try { cmd.ExecuteNonQuery().ToString(); }
                catch (Exception exp) { MessageBox.Show(exp.ToString()); }
               
            }
            cn.Close();
        }

        static void DeleteTable(string tablename)
        {
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = "DROP TABLE IF EXISTS "+tablename;
                try { cmd.ExecuteNonQuery().ToString(); }
                catch (Exception exp) { MessageBox.Show(exp.ToString()); }
            }
            cn.Close();
        }

        static void ClearTable(string tablename)
        {
           
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = "DELETE FROM sqlite_sequence WHERE name = "+ tablename + ';';
                try { cmd.ExecuteNonQuery().ToString(); }
                catch (Exception exp) { MessageBox.Show(exp.ToString()); }
            }
            cn.Close();
        }

        static void InsertContractRow(string contract_name,int lastnumber)
        {
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = "INSERT INTO contract(contract_name,lastnumber) VALUES(@contract_name,@lastnumber)";
                cmd.Parameters.Add("contract_name", DbType.String).Value = contract_name;
                cmd.Parameters.Add("lastnumber", DbType.Int32).Value = lastnumber;
                cmd.ExecuteNonQuery();
            }
            cn.Close();
        }

        //contract_name TEXT NOT NULL,creattime VARCHAR(18) NOT NULL,url VARCHAR(255) NOT NULL,return BLOB NOT NULL,returnurl VARCHAR(255) NOT NULL ,returntime VARCHAR(18)
        static void InsertContractdetailRow(string contract_name, string creattime, string url,int returned, string returnurl,string returntime,string linkname)
        {
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                cmd.CommandText = "INSERT INTO contract_detail(contract_name,creattime,url,return,returnurl,returntime,linkname) VALUES(@contract_name,@creattime,@url,@return,@returnurl,@returntime,@linkname)";
                cmd.Parameters.Add("contract_name", DbType.String).Value = contract_name;
                cmd.Parameters.Add("creattime", DbType.String).Value = creattime;
                cmd.Parameters.Add("url", DbType.String).Value = url;
                cmd.Parameters.Add("return", DbType.Int32).Value = returned;
                cmd.Parameters.Add("returnurl", DbType.String).Value = returnurl;
                cmd.Parameters.Add("returntime", DbType.String).Value = returntime;
                cmd.Parameters.Add("linkname", DbType.String).Value = linkname;
                cmd.ExecuteNonQuery();
            }
            cn.Close();
        }

        static int selectmaxnum(string contract_name)
        {
            int res = 0;
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                //查询第1条记录，这个并不保险,rowid 并不是连续的，只是和当时插入有关
                cmd.CommandText = "SELECT lastnumber FROM contract WHERE contract_name='" + contract_name+"';";
                SQLiteDataReader sr = cmd.ExecuteReader();
                while (sr.Read())
                {
                    res=sr.GetInt32(0);
                }
                sr.Close();
            }
            cn.Close();
            return res;
        }

        static int searchcontract(string contract_name)
        {
            int res = 0;
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                //查询第1条记录，这个并不保险,rowid 并不是连续的，只是和当时插入有关
                cmd.CommandText = "SELECT count(*) FROM contract WHERE contract_name='" + contract_name + "';";
                SQLiteDataReader sr = cmd.ExecuteReader();
                while (sr.Read())
                {
                    res = sr.GetInt32(0);
                }
                sr.Close();
            }
            cn.Close();
            return res;
        }

        static int addmaxnum(string contract_name)
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
                //查询第1条记录，这个并不保险,rowid 并不是连续的，只是和当时插入有关
                cmd.CommandText = "SELECT lastnumber FROM contract WHERE contract_name='" + contract_name + "';";
                SQLiteDataReader sr = cmd.ExecuteReader();
                while (sr.Read())
                {
                    res = sr.GetInt32(0);
                }
                sr.Close();
                cmd.CommandText = "UPDATE contract SET lastnumber=@lastnumber WHERE contract_name='" + contract_name + "'";
                cmd.Parameters.Add("lastnumber", DbType.Int32).Value = res+1;
                returnnum=cmd.ExecuteNonQuery();

            }
            cn.Close();
            return returnnum;
        }

        static int contractreturned(string contract_name, string return_url, string return_time, decimal money, decimal money1)
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

        public int contractcanceled(string contract_name)
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
                cmd.CommandText = "UPDATE contract_detail SET cancel=@cancel WHERE contract_name='" + contract_name + "'";
                cmd.Parameters.Add("cancel", DbType.Int32).Value = 1;
                returnnum = cmd.ExecuteNonQuery();
                //MessageBox.Show(cmd.CommandText);
            }
            cn.Close();
            return returnnum;
        }

        public void showcontracttable()
        {
            DataTable dt = new DataTable("contract");
            DataColumn dc = dt.Columns.Add("contract_name", Type.GetType("System.String"));
            dc = dt.Columns.Add("lastnumber", Type.GetType("System.Int32"));


            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                //查询第1条记录，这个并不保险,rowid 并不是连续的，只是和当时插入有关
                cmd.CommandText = "SELECT * FROM contract ;";
                try
                {
                    SQLiteDataReader sr = cmd.ExecuteReader();
                    while (sr.Read())
                    {
                        if (sr.HasRows)
                        {
                            DataRow dr = dt.NewRow();
                            dr["contract_name"] = sr.GetString(0);
                            dr["lastnumber"] = sr.GetInt32(1);
                            dt.Rows.Add(dr);
                        }

                    }
                    sr.Close();
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.ToString());
                }
               
            }
            cn.Close();
            this.dataGridView1.DataSource = dt;

        }
        //获取合同模板名称
        public void showcontractname()
        {
            DataTable dt = new DataTable("contract");
            DataColumn dc = dt.Columns.Add("contract_name", Type.GetType("System.String"));

            DataRow dr = dt.NewRow();
            dr["contract_name"] = "";
            dt.Rows.Add(dr);

            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                //查询第1条记录，这个并不保险,rowid 并不是连续的，只是和当时插入有关
                cmd.CommandText = "SELECT contract_name FROM contract ;";
                try
                {
                    SQLiteDataReader sr = cmd.ExecuteReader();
                    while (sr.Read())
                    {
                        if (sr.HasRows)
                        {
                            dr = dt.NewRow();
                            dr["contract_name"] = sr.GetString(0);
                            dt.Rows.Add(dr);
                        }

                    }
                    sr.Close();
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.ToString());
                }

            }
            cn.Close();
            this.comboBox1.DataSource = dt;
            comboBox1.ValueMember = "contract_name";
            comboBox1.DisplayMember = "contract_name";

        }

        public void showcontractdetailtable(string searchtxt)
        {
            //contract_name TEXT NOT NULL,creattime VARCHAR(18) NOT NULL,url VARCHAR(255) NOT NULL,return BLOB NOT NULL,returnurl VARCHAR(255) NOT NULL ,returntime VARCHAR(18)

            DataTable dt = new DataTable("contract_detail");
            DataColumn dc = dt.Columns.Add("合同名称", Type.GetType("System.String"));
            dc = dt.Columns.Add("创建时间", Type.GetType("System.String"));
            dc = dt.Columns.Add("原始路径", Type.GetType("System.String"));
            dc = dt.Columns.Add("是否归档", Type.GetType("System.Int32"));
            dc = dt.Columns.Add("归档路径", Type.GetType("System.String"));
            dc = dt.Columns.Add("归档时间", Type.GetType("System.String"));
            dc = dt.Columns.Add("合同金额", Type.GetType("System.Decimal"));
            dc = dt.Columns.Add("已交金额", Type.GetType("System.Decimal"));
            dc = dt.Columns.Add("取消", Type.GetType("System.Int32"));
            dc = dt.Columns.Add("关联文件", Type.GetType("System.String"));
            string path = @"contract.db";
            SQLiteConnection cn = new SQLiteConnection("data source=" + path);
            if (cn.State != System.Data.ConnectionState.Open)
            {
                cn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = cn;
                //查询第1条记录，这个并不保险,rowid 并不是连续的，只是和当时插入有关
                cmd.CommandText = "";
                if (searchtxt == "")
                { cmd.CommandText = "SELECT * FROM contract_detail WHERE cancel=0 ORDER BY creattime DESC;"; }
                else
                {
                    cmd.CommandText = "SELECT * FROM contract_detail " + searchtxt+ " and cancel=0 ORDER BY creattime DESC;";
                }
                try
                {
                    SQLiteDataReader sr = cmd.ExecuteReader();
                    while (sr.Read())
                    {
                        if (sr.HasRows)
                        {
                            DataRow dr = dt.NewRow();
                            dr["合同名称"] = sr.GetString(0);
                            dr["创建时间"] = sr.GetString(1);
                            dr["原始路径"] = sr.GetString(2);
                            dr["是否归档"] = sr.GetInt32(3);

                            dr["归档路径"] = sr.GetValue(4);
                            dr["归档时间"] = sr.GetValue(5);
                            dr["合同金额"] = sr.GetValue(6);
                            dr["已交金额"] = sr.GetValue(7);
                            dr["取消"] = sr.GetInt32(8);
                            dr["关联文件"] = sr.GetString(9);
                            dt.Rows.Add(dr);
                        }

                    }
                    sr.Close();
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.ToString());
                }
                
            }
            cn.Close();
           
            this.dataGridView2.DataSource = dt;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateDB();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            CreateTable("CREATE TABLE IF NOT EXISTS contract( contract_name TEXT NOT NULL,lastnumber INT ); ");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(this.comboBox3.Text!="")
            {
                int res = searchcontract(this.comboBox3.Text);
                if (res == 0)
                {
                    InsertContractRow(this.comboBox3.Text, 0);
                }
                else
                {
                    MessageBox.Show("合同模板" + this.comboBox3.Text + "已经存在");
                }
            }
            showcontractname();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            DeleteTable("contract");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            this.label1.Text = selectmaxnum(this.comboBox1.Text).ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            showcontracttable();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
           
            if (this.comboBox2.Text == "")
            {
                MessageBox.Show("请选择合同类型");
                return;
            }
            backup("创建"+ this.comboBox1.Text + "合同");
            DialogResult dr = MessageBox.Show("你确定要创建"+ this.comboBox1.Text + "合同吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.OK)
            {
                if (this.comboBox1.Text == "嘉兴爱阁家政服务有限公司月子护理服务协议" || this.comboBox1.Text == "月拇指母婴服务委托服务协议")
                {
                    //本合同
                    DateTime dt = DateTime.Now;

                    AsposeWordHelper helper = new AsposeWordHelper();

                    string templatePath = this.comboBox2.Text;  //模板路径
                    helper.OpenTempelte(templatePath); //打开模板文件
                    string FileName = "contractoutput/" + this.comboBox1.Text + ((202000 + dt.Month) * 10000 + selectmaxnum(this.comboBox1.Text) + 1).ToString() + ".pdf";//保存路径
                                                                                                                                                                          //MessageBox.Show(FileName);
                    string[] fieldNames = new string[] { "contractNum" };
                    object[] fieldValues = new object[] { ((202000 + dt.Month) * 10000 + selectmaxnum(this.comboBox1.Text) + 1).ToString() };
                    helper.Executefield(fieldNames, fieldValues);//域赋值
                    helper.SavePdf(FileName); //文件保存，保存为pdf
                    int effectrosnum = addmaxnum(this.comboBox1.Text);

                    this.label1.Text = selectmaxnum(this.comboBox1.Text).ToString();
                    //获取关联合同的名字
                    string FileName1 = "助邦母婴护理师委托服务合同" + ((202000 + dt.Month) * 10000 + selectmaxnum("助邦母婴护理师委托服务合同") + 1).ToString() + ".pdf";

                    InsertContractdetailRow(FileName.Split('/')[1], dt.ToString(), FileName, 0, null, null, FileName1);
                    //关联合同
                    templatePath = @"temple\助邦母婴护理师委托服务合同\助邦母婴护理师委托服务合同.docx";
                    helper.OpenTempelte(templatePath); //打开模板文件
                    string[] fieldNames1 = new string[] { "contractNum", "contractNum1" };
                    object[] fieldValues1 = new object[] { ((202000 + dt.Month) * 10000 + selectmaxnum("助邦母婴护理师委托服务合同") + 1).ToString(), FileName.Split('/')[1] };
                    helper.Executefield(fieldNames1, fieldValues1);//域赋值
                    helper.SavePdf("contractoutput/"+FileName1); //文件保存，保存为pdf
                    addmaxnum("助邦母婴护理师委托服务合同");
                    InsertContractdetailRow(FileName1, dt.ToString(), "contractoutput/"+FileName1, 0, null, null, FileName.Split('/')[1]);
                }
                else {
                    //确定
                    DateTime dt = DateTime.Now;

                    AsposeWordHelper helper = new AsposeWordHelper();

                    string templatePath = this.comboBox2.Text;  //模板路径
                    helper.OpenTempelte(templatePath); //打开模板文件
                    string FileName = "contractoutput/" + this.comboBox1.Text + ((202000 + dt.Month) * 10000 + selectmaxnum(this.comboBox1.Text) + 1).ToString() + ".pdf";//保存路径
                                                                                                                                                                          //MessageBox.Show(FileName);
                    string[] fieldNames = new string[] { "contractNum" };
                    object[] fieldValues = new object[] { ((202000 + dt.Month) * 10000 + selectmaxnum(this.comboBox1.Text) + 1).ToString() };
                    helper.Executefield(fieldNames, fieldValues);//域赋值
                    helper.SavePdf(FileName); //文件保存，保存为pdf
                    int effectrosnum = addmaxnum(this.comboBox1.Text);

                    this.label1.Text = selectmaxnum(this.comboBox1.Text).ToString();
                    //
                    InsertContractdetailRow(FileName.Split('/')[1], dt.ToString(), FileName, 0, null, null," ");

                }
                string sqltxt = "";
                if (this.comboBox1.Text != "")
                {
                    sqltxt = "WHERE contract_name LIKE '" + this.comboBox1.Text + "%' ";
                }
                showcontractdetailtable(sqltxt);
                
            }
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CreateTable("CREATE TABLE contract_detail( contract_name TEXT NOT NULL,creattime VARCHAR(18) NOT NULL,url VARCHAR(255) NOT NULL,return BLOB NOT NULL,returnurl VARCHAR(255),returntime VARCHAR(18));");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //ClearTable();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            
                DeleteTable("contract_detail");
        }

        public string openurl = "";
        public string printurl = "";
        public string filename = "";
        public string link = "";

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
           
            if (e.ColumnIndex < 0 || e.RowIndex < 0) return;
            filename = "";
            printurl = "";
            openurl = "";
            this.dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;// 单击选中整行，枚举
            int n = this.dataGridView2.Columns.Count;
           

            filename = this.dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
            printurl = this.dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
            link= this.dataGridView2.Rows[e.RowIndex].Cells[9].Value.ToString();
            if (this.dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString() == "0")
            {
                //未归档
                this.button11.Enabled = false;
                this.button12.Enabled = true;
                this.button13.Enabled = true;
            }
            else {
                //归档
                this.button11.Enabled = true;
                this.button12.Enabled = false;
                this.button13.Enabled = false;
                openurl = this.dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                
            }
            if (link.Length > 0&&link!=" ")
            {
                this.button18.Enabled = true;
            }
            else
            {
                this.button18.Enabled =false;
            }

            this.label2.Text = "文件名："+ filename +" 打印路径："+ printurl+" 打开路径："+ openurl;

        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                new 合同查看(openurl, filename).ShowDialog();
            }
            catch (Exception exp)
            {
                MessageBox.Show("请选择单元格！");
            }

        }

        //归档
        private void button12_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            this.panel1.Width = this.Width;
            this.panel1.Height = this.Height;
            this.panel1.Location = new Point(0, 0);
          //  new 合同明细(filename).ShowDialog();
            /*
            DialogResult dr = openFileDialog1.ShowDialog();
            //获取所打开文件的文件名
            string openfilename = openFileDialog1.FileName;

            if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(openfilename))
            {
                string pLocalFilePath = openfilename;//要复制的文件路径
                string filename1 = System.IO.Path.GetFileName(pLocalFilePath);
               
                string pSaveFilePath = "return/"+ filename;//指定存储的路径
               // MessageBox.Show(pSaveFilePath);
                if (File.Exists(pLocalFilePath))//必须判断要复制的文件是否存在
                {
                    File.Copy(pLocalFilePath, pSaveFilePath, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
                    DateTime dt = DateTime.Now;
                    contractreturned(filename, pSaveFilePath, dt.ToString());
                    showcontractdetailtable();

                    this.button11.Enabled = true;
                    this.button12.Enabled = false;
                    openurl = pSaveFilePath;
                }

            }
            */
            
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                new 合同打印(printurl, filename).ShowDialog();
            }
            catch (Exception exp)
            { 
                MessageBox.Show("请选择单元格！");
            }
           
        }
       public int countcombobox = 0;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var name = "";
            if (countcombobox > 1)
            {
                name = this.comboBox1.SelectedValue.ToString();
                //MessageBox.Show(name);
            }
            countcombobox++;
            //月拇指母婴服务委托服务协议
            IList<string> list = new List<string>();

            var files = Directory.GetFiles("temple\\" + name, "*.docx");

            foreach (var file in files)
            {
                list.Add(file);
            }
            this.comboBox2.DataSource = null;
            this.comboBox2.DataSource = list;
            string sqltxt = "";
            if (this.comboBox1.Text != "")
            {
                sqltxt = "WHERE contract_name LIKE '" + this.comboBox1.Text + "%' ";
            }
            showcontractdetailtable(sqltxt);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this.comboBox2.Text);   
        }

        public static bool IsMoney(string input)
        {
            string pattern = @"^\-{0,1}[0-9]{0,}\.{0,1}[0-9]{1,}$";
            return Regex.IsMatch(input, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            if (IsMoney(this.textBox1.Text) && IsMoney(this.textBox2.Text))
            {
                DialogResult dr = openFileDialog1.ShowDialog();
                //获取所打开文件的文件名
                string openfilename = openFileDialog1.FileName;

                if (dr == System.Windows.Forms.DialogResult.OK && !string.IsNullOrEmpty(openfilename))
                {
                    backup("归档" + this.comboBox1.Text + "合同");
                    string pLocalFilePath = openfilename;//要复制的文件路径
                    string filename1 = System.IO.Path.GetFileName(pLocalFilePath);

                    string pSaveFilePath = "return/" + filename;//指定存储的路径
                                                             // MessageBox.Show(pSaveFilePath);
                    if (File.Exists(pLocalFilePath))//必须判断要复制的文件是否存在
                    {
                        File.Copy(pLocalFilePath, pSaveFilePath, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
                        DateTime dt = DateTime.Now;
                        contractreturned(filename, pSaveFilePath, dt.ToString(), Convert.ToDecimal(this.textBox1.Text), Convert.ToDecimal(this.textBox2.Text));

                        showcontractdetailtable("");

                        this.button11.Enabled = true;
                        this.button12.Enabled = false;
                        openurl = pSaveFilePath;
                        this.panel1.Visible = false;
                        this.textBox1.Text = "0";
                        this.textBox2.Text = "0";
                    }

                }

            }
            else
            {
                MessageBox.Show("数字输入错误");
                return;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            this.panel1.Visible = false;
        }
         public bool IsUnsign(string value)
        {
           return Regex.IsMatch(value, @"^\d*[.]?\d*$");
        }
    private void button16_Click(object sender, EventArgs e)
        {
            if (IsUnsign(this.textBox3.Text))
            {
                string sqltxt = "";
                if (this.comboBox1.Text == "")
                {
                    sqltxt = "WHERE contract_name LIKE '%" + this.textBox3.Text + "%' ";
                }
                else {
                    sqltxt = "WHERE contract_name LIKE '" + this.comboBox1.Text+this.textBox3.Text + "%' ";
                }
                showcontractdetailtable(sqltxt);
            }
            else
            {
                MessageBox.Show("合同号无效！");
            }
           
        }

        private void button17_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("你确定要取消" + filename + "合同吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.OK)
            {
                contractcanceled(filename);

                showcontractdetailtable("");
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            //
            MessageBox.Show(link);
            if (File.Exists(@"contractoutput/"+ link))
            {
                if (File.Exists(@"return/" + link))
                {
                    try
                    {
                        new 合同查看(@"return/" + link, link.Replace(".pdf","")).ShowDialog();
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("请选择单元格！");
                    }
                }
                else {
                    try
                    {
                        new 合同查看(@"contractoutput/" + link, link.Replace(".pdf", "")).ShowDialog();
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("请选择单元格！");
                    }
                }
            }
        }
    }
}
