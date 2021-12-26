using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Drawing.Printing;

namespace Myfactory
{
    

    public partial class FactoryForm : Form
    {
        public static string connStr = "server=Localhost;database=work;port=3306;username=root;password=123456789;charset=utf8";
        MySqlConnection conn = new MySqlConnection(connStr);     
        string[] FactoryValue = new string[13];//存dataGridView1的值
        string[] MoneyValue = new string[8];//存dataGridView3.4.8的值
        string[] NewMoneyValue = new string[13];//存dataGridView2的值
        string[] PrintMoneyValue = new string[8];//存dataGridView5的值
        string[] PrintFactoryValue = new string[13];//

        public FactoryForm()
        {
            InitializeComponent();
        }

        public static DataTable GetDataTable(string sql)
        {
            DataTable dt = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter(sql, connStr);
            adapter.Fill(dt);           
            return dt;
        }

        public void moneyData()
        {
            string sql;
            DateTime date = DateTime.Now;
            string monthDate = date.AddMonths(-1).ToString("yyyyMMdd");
            sql = "select * from work_hour where date >='" + monthDate + " '  ";
            dataGridView3.DataSource = GetDataTable(sql);

            string weekDate = date.AddDays(-7).ToString("yyyyMMdd");
            sql = "select * from work_hour where date >='" + weekDate + " '  ";
            dataGridView4.DataSource = GetDataTable(sql);

            sql = "select * from work_hour order by date DESC limit 300";
            dataGridView8.DataSource = GetDataTable(sql);


        }

        public void kindChange()
        {
            string sql;
            DateTime date = DateTime.Now;
            string monthDate = date.AddMonths(-1).ToString("yyyyMMdd");
            sql = "select * from work_hour where date >='" + monthDate + " '  ";
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView3.Rows[i].Cells[2].Value = "現金";

                }
                else
                {
                    dataGridView3.Rows[i].Cells[2].Value = "月結";
                }
            }
            string weekDate = date.AddDays(-7).ToString("yyyyMMdd");
            sql = "select * from work_hour where date >='" + weekDate + " '  ";
            dt = GetDataTable(sql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView4.Rows[i].Cells[2].Value = "現金";

                }
                else
                {
                    dataGridView4.Rows[i].Cells[2].Value = "月結";
                }
            }

            sql = "select * from work_hour order by date DESC limit 300";           
            dt = GetDataTable(sql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView8.Rows[i].Cells[2].Value = "現金";

                }
                else
                {
                    dataGridView8.Rows[i].Cells[2].Value = "月結";
                }
            }
        }

        private void FactoryForm_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'workhour.work_hour' 資料表。您可以視需要進行移動或移除。
            //this.work_hourTableAdapter.Fill(this.workhour.work_hour);
            // TODO: 這行程式碼會將資料載入 'workDataSet.factory' 資料表。您可以視需要進行移動或移除。
            this.factoryTableAdapter.Fill(this.workDataSet.factory);

        }



        private void dataGridView2_Layout(object sender, LayoutEventArgs e)//收入管理-查詢公司新增收入資料表連線
        {
            textBox27.Text = DateTime.Now.ToString("yyyyMMdd");
            string sql = "select * from factory";
            dataGridView2.DataSource = GetDataTable(sql); 
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].Visible = false;
            dataGridView2.Columns[11].Visible = false;            
            dataGridView2.Columns[2].HeaderText = "公司名稱";
            dataGridView2.Columns[3].HeaderText = "負責人";
            dataGridView2.Columns[4].HeaderText = "地址";
            dataGridView2.Columns[5].HeaderText = "工作地點";
            dataGridView2.Columns[6].HeaderText = "電話1";
            dataGridView2.Columns[7].HeaderText = "電話2";
            dataGridView2.Columns[8].HeaderText = "手機";
            dataGridView2.Columns[9].HeaderText = "出車種類";
            dataGridView2.Columns[10].HeaderText = "工作內容";
            dataGridView2.Columns[12].HeaderText = "備註";

        }

        private void FactoryForm_FormClosing(object sender, FormClosingEventArgs e)//程式關閉時
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)//公司列表查詢鍵
        {
            if (textBox1.Text.Trim().Length > 0)
            {
               
                string sql = "select * from factory where f_name like '%" + textBox1.Text + "%'  " +
                                                         "or name  like '%" + textBox1.Text + "%' "+
                                                         "or address  like '%" + textBox1.Text + "%' "+
                                                         "or work_place  like '%" + textBox1.Text + "%' " +
                                                         "or tel  like '%" + textBox1.Text + "%' " +
                                                         "or tel1  like '%" + textBox1.Text + "%' " +
                                                         "or hand  like '%" + textBox1.Text + "%' " +
                                                         "or work_kind  like '%" + textBox1.Text + "%' " +
                                                         "or work_add  like '%" + textBox1.Text + "%' " +
                                                         "or other  like '%" + textBox1.Text + "%' " ;  //模糊查詢公司名稱                
                dataGridView1.DataSource = GetDataTable(sql);
            }
            else
            {
                string sql = "select * from Factory";
                dataGridView1.DataSource = GetDataTable(sql);
            }

        }

        private void button4_Click(object sender, EventArgs e)//新增公司-新增鍵
        {
            int payKind;
            if (radioButton1.Checked == true)
            {
                 payKind = 1;
            }
            else
            {
                payKind = 3;
            }

            try
            {
                /* DataRow dr = ds.Tables[0].NewRow();
                 dr[0] = "萬萬";
                 ds.Tables[0].Rows.Add(dr);
                 bs.DataSource = ds.Tables[0];
                 dataGridView1.DataSource = bs;*/
                /* string st = "INSERT INTO factory (f_name)" + "VALUES('萬萬')";
                 MySqlCommand cmd = new MySqlCommand(st,conn);
                 cmd.Parameters.AddWithValue("f_name", "萬萬");
                 cmd.ExecuteNonQuery();*/


                /* MySqlDataAdapter da = new MySqlDataAdapter("select * From fatory", conn);
                 DataSet ds = new DataSet();
                 da.Fill(ds, "factory");
                 DataTable dt = ds.Tables["factory"];
                 dt.Constraints.Clear();
                 foreach (DataColumn dcA in dt.Columns)
                       {
                           dcA.AllowDBNull = true;
                       }*/

                /*dt.Rows.Add("+textBox3.Text+", "+textBox4.Text+", " + textBox5.Text+",
                            "+textBox6.Text+", " + textBox7.Text+", "+textBox8.Text+",
                            "+textBox9.Text+", " + textBox10.Text+", "+textBox11.Text+",
                            "+textBox12.Text+", " + textBox13.Text+", "+paykind+", " + textBox15+");*/
                //dt.Rows.Add("03557", "11", "萬萬", "11", "11", "11", "11", "11", "11", "11", "11", "1","11");

                /*dataGridView1.DataSource = null;
                dataGridView2.DataSource = null;
                string sql = "select * from factory";
                DataTable dt = GetDataTable(sql);
                DataRow dr = dt.NewRow();               
                dr[0] = textBox3.Text;
                dr[1] = textBox4.Text;
                dr[2] = textBox5.Text;
                dr[3] = textBox6.Text;
                dr[4] = textBox7.Text;
                dr[5] = textBox8.Text;
                dr[6] = textBox9.Text;
                dr[7] = textBox10.Text;
                dr[8] = textBox11.Text;
                dr[9] = textBox12.Text;
                dr[10] = textBox13.Text;
                dr[11] = payKind;
                dr[12] = textBox15.Text;
                dt.Rows.Add(dr);
                dataGridView1.DataSource =dt;
                dataGridView2.DataSource =dt;*/

                string sql = "INSERT INTO factory (id ,id_num, f_name, name, address, work_place, tel, tel1, hand, work_kind, work_add, pay_kind, other) " +
                " VALUES ('"
               + textBox3.Text + "','"
               + textBox4.Text.Trim() + "','"
               + textBox5.Text + "','"
               + textBox6.Text + "','"
               + richTextBox1.Text + "','"
               + richTextBox2.Text + "','"
               + textBox9.Text.Trim() + "','"
               + textBox10.Text.Trim() + "','"
               + textBox11.Text.Trim() + "','"
               + richTextBox3.Text + "','"
               + richTextBox4.Text + "','"
               + payKind + "','"
               + richTextBox5.Text + "')";

                GetDataTable(sql);

                sql = "select * from factory";
                dataGridView1.DataSource = GetDataTable(sql);
                dataGridView2.DataSource = GetDataTable(sql);

                MessageBox.Show("公司資料已新增", "提示", MessageBoxButtons.OK,MessageBoxIcon.Information) ;

                tabControl2.SelectedTab = tabPage4;

                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                richTextBox1.Text = "";
                richTextBox2.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                richTextBox5.Text = "";


            }
            
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)//收入管理-查詢鍵
        {
            if (textBox2.Text.Trim().Length > 0)
            {
                string sql = "select * from factory where f_name like '%" + textBox2.Text + "%'  " +
                                                          "or name  like '%" + textBox2.Text + "%' " +
                                                          "or address  like '%" + textBox2.Text + "%' " +
                                                          "or work_place  like '%" + textBox2.Text + "%' " +
                                                          "or tel  like '%" + textBox2.Text + "%' " +
                                                          "or tel1  like '%" + textBox2.Text + "%' " +
                                                          "or hand  like '%" + textBox2.Text + "%' " +
                                                          "or work_kind  like '%" + textBox2.Text + "%' " +
                                                          "or work_add  like '%" + textBox2.Text + "%' " +
                                                          "or other  like '%" + textBox2.Text + "%' ";  //模糊查詢公司名稱              
                dataGridView2.DataSource = GetDataTable(sql);
            }
            else
            {
                string sql = "select * from factory";
                dataGridView2.DataSource = GetDataTable(sql);

            }
        }

        

        private void tabPage5_Layout(object sender, LayoutEventArgs e)//新增公司頁面
        {
            string sql = "select * from factory order by id DESC limit 1";
            DataTable dt = GetDataTable(sql);

            int num = int.Parse(dt.Rows[0][0].ToString());
            num++;
            string id = num.ToString();
            if (id.Length == 1)
            {
                id = "0000" + id;
            }
            else if (id.Length == 2)
            {
                id = "000" + id;
            }
            else if (id.Length == 3)
            {
                id = "00" + id;
            }
            else if (id.Length == 4)
            {
                id = "0" + id;
            }
            else
            {
                id = id;
            }
            textBox3.Text = id;


        }


        private void dataGridView1_Layout(object sender, LayoutEventArgs e)//查詢公司頁面
        {
            string sql = "select * from factory  ";
            dataGridView1.DataSource = GetDataTable(sql);
        }

       private void button5_Click(object sender, EventArgs e) //修改公司-修改鍵
        {
            if (textBox26.Text.Length == 0)
            {
                MessageBox.Show("請先至 [查詢公司]頁面 選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                if (MessageBox.Show("確定要修改資料嗎?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    int payKind;
                    if (radioButton4.Checked == true)
                    {
                        payKind = 1;
                    }
                    else
                    {
                        payKind = 3;
                    }
                    try
                    {
                        string sql = "update factory set id_num='" + textBox25.Text + "'  " +
                                                        ",f_name='" + textBox24.Text + "'  " +
                                                        ",name='" + textBox23.Text + "' " +
                                                        ",address='" + richTextBox6.Text + "' " +
                                                        ",work_place='" + richTextBox7.Text + "' " +
                                                        ",tel='" + textBox20.Text + "' " +
                                                        ",tel1='" + textBox19.Text + "' " +
                                                        ",hand='" + textBox18.Text + "' " +
                                                        ",work_kind='" + richTextBox8.Text + "' " +
                                                        ",work_add='" + richTextBox9.Text + "' " +
                                                        ",pay_kind='" + payKind + "' " +
                                                        ",other='" + richTextBox10.Text + "' " +
                                                        "where id = '" + textBox26.Text + "' ";
                        GetDataTable(sql);

                        sql = "select * from factory";
                        dataGridView1.DataSource = GetDataTable(sql);
                        dataGridView2.DataSource = GetDataTable(sql);
                        MessageBox.Show("資料已修改!","提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        textBox26.Text = "";
                        textBox25.Text = "";
                        textBox24.Text = "";
                        textBox23.Text = "";
                        richTextBox6.Text = "";
                        richTextBox7.Text = "";
                        textBox20.Text = "";
                        textBox19.Text = "";
                        textBox18.Text = "";
                        richTextBox8.Text = "";
                        richTextBox9.Text = "";
                        richTextBox10.Text = "";
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)//修改公司-恢復鍵
        {
            textBox26.Text = FactoryValue[0];           
            textBox25.Text = FactoryValue[1];
            textBox24.Text = FactoryValue[2];
            textBox23.Text = FactoryValue[3];
            richTextBox6.Text = FactoryValue[4];
            richTextBox7.Text = FactoryValue[5];
            textBox20.Text = FactoryValue[6];
            textBox19.Text = FactoryValue[7];
            textBox18.Text = FactoryValue[8];
            richTextBox8.Text = FactoryValue[9];
            richTextBox9.Text = FactoryValue[10];
            richTextBox10.Text = FactoryValue[12];

            if (FactoryValue[11] == "3")       //月結判斷
            {
                radioButton3.Checked = true;
            }
            else
            {
                radioButton4.Checked = true;
            }

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < 13; i++)//讀取data,一次讀一格
            {
                FactoryValue[i] = dataGridView1.Rows[e.RowIndex].Cells[i].Value.ToString();
            }
            textBox26.Text = FactoryValue[0];           //將資料放到textBox
            textBox25.Text = FactoryValue[1];
            textBox24.Text = FactoryValue[2];
            textBox23.Text = FactoryValue[3];
            richTextBox6.Text = FactoryValue[4];
            richTextBox7.Text = FactoryValue[5];
            textBox20.Text = FactoryValue[6];
            textBox19.Text = FactoryValue[7];
            textBox18.Text = FactoryValue[8];
            richTextBox8.Text = FactoryValue[9];
            richTextBox9.Text = FactoryValue[10];
            richTextBox10.Text = FactoryValue[12];



            if (FactoryValue[11] == "3")       //月結判斷
            {
                radioButton3.Checked = true;
            }
            else
            {
                radioButton4.Checked = true;
            }

            tabControl2.SelectedTab = tabPage6;

        }
      
        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < 13; i++)//讀取data,一次讀一格
            {
                NewMoneyValue[i] = dataGridView2.Rows[e.RowIndex].Cells[i].Value.ToString();
            }
            textBox41.Text = textBox27.Text;
            textBox39.Text = NewMoneyValue[0];           //將資料放到textBox
            textBox38.Text = NewMoneyValue[1];
            textBox37.Text = NewMoneyValue[2];
            textBox36.Text = NewMoneyValue[3];
            richTextBox11.Text = NewMoneyValue[4];
            richTextBox12.Text = NewMoneyValue[5];
            textBox33.Text = NewMoneyValue[6];
            textBox32.Text = NewMoneyValue[7];
            textBox31.Text = NewMoneyValue[8];
            richTextBox13.Text = NewMoneyValue[10];
            richTextBox14.Text = NewMoneyValue[12];

            textBox40.Text = DateTime.Now.ToString("yyMMddHHmms");



            if (NewMoneyValue[11] == "3")       //月結判斷
            {
                radioButton5.Checked = true;
            }
            else
            {
                radioButton6.Checked = true;
            }

            tabControl3.SelectedTab = tabPage8;
        }
         
        private void button7_Click(object sender, EventArgs e)//新增收入-新增鍵
        {
            if (textBox39.Text.Length == 0)
            {
                MessageBox.Show("請先至 [查詢公司新增收入]頁面 選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox49.Text.Trim().Length == 0)
            {
                MessageBox.Show("請輸入付款金額", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox49.Text.Trim().All(char.IsDigit) == false)
            {
                MessageBox.Show("付款金額只能輸入數字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox41.Text.Trim().Length != 8)
            {
                MessageBox.Show("工作日期格式錯誤", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                int payKind;
                if (radioButton6.Checked)
                {
                    payKind = 2;
                }
                else
                {
                    payKind = 3;
                }

                try
                {
                    int intMoney = int.Parse(textBox49.Text.Trim());
                    string stMoney = intMoney.ToString();
                    ; string sql = "INSERT INTO work_hour (id , f_name, kind,time,carkind,money,date,other) " +
                        " VALUES ('"
                       + textBox39.Text.Trim() + "','"
                       + textBox37.Text.Trim() + "','"
                       + payKind + "','"
                       + textBox40.Text.Trim() + "','"
                       + textBox30.Text.Trim() + "','"
                       + stMoney + "','"
                       + textBox41.Text.Trim() + "','"
                       + richTextBox15.Text.Trim() + "')";

                    GetDataTable(sql);

                    moneyData();

                    kindChange();

                    sql = "select * from work_hour order by date DESC limit 500";
                    dataGridView5.DataSource = GetDataTable(sql);
                    DataTable dt = GetDataTable(sql);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        if (dt.Rows[i][2].ToString() == "3")
                        {
                            dataGridView5.Rows[i].Cells[2].Value = "現金";

                        }
                        else
                        {
                            dataGridView5.Rows[i].Cells[2].Value = "月結";
                        }
                    }

                    MessageBox.Show("公司資料已新增", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    textBox41.Text = "";
                    textBox39.Text = "";           //將資料放到textBox
                    textBox38.Text = "";
                    textBox37.Text = "";
                    textBox36.Text = "";
                    richTextBox11.Text = "";
                    richTextBox12.Text = "";
                    textBox33.Text = "";
                    textBox32.Text = "";
                    textBox31.Text = "";
                    richTextBox13.Text = "";
                    richTextBox14.Text = "";
                    textBox40.Text = "";
                    textBox30.Text = "";
                    richTextBox15.Text = "";
                    textBox49.Text = "";
                    

                    tabControl3.SelectedTab = tabPage9;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void button9_Click(object sender, EventArgs e) //收入修改-修改鍵
        {
            if (textBox44.Text.Length == 0)
            {
                MessageBox.Show("請先至 [收入列表]頁面 選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox45.Text.Trim().Length == 0)
            {
                MessageBox.Show("請輸入付款金額", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox45.Text.Trim().All(char.IsDigit) == false)
            {
                MessageBox.Show("付款金額只能輸入數字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (textBox47.Text.Trim().Length != 8)
            {
                MessageBox.Show("工作日期格式錯誤", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                
                if (MessageBox.Show("確定要修改資料嗎?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    int payKind;
                    if (radioButton8.Checked == true)
                    {
                        payKind = 2;
                    }
                    else
                    {
                        payKind = 3;
                    }
                    try
                    {
                        int intMoney = int.Parse(textBox45.Text.Trim());
                        string stMoney = intMoney.ToString();
                        string sql = "update work_hour set date='" + textBox47.Text + "'  " +
                                                        ",time='" + textBox48.Text + "'  " +
                                                        ",money='" + stMoney + "'  " +
                                                        ",carkind='" + textBox50.Text + "' " +
                                                        ",kind='" + payKind + "' " +
                                                        ",other='" + richTextBox16.Text + "' " +
                                                        "where id = '" + MoneyValue[1] + "' and f_name = '"+MoneyValue[0]+"' and time='" + MoneyValue[3] + "' " +
                                                        "and carkind ='" + MoneyValue[4] + "'  and money = '" + MoneyValue[5] + "'and date ='" + MoneyValue[6] + "'" +
                                                        "and other = '" + MoneyValue[7] + "' and kind = '" + MoneyValue[2] + "' ";
                        GetDataTable(sql);

                        moneyData();

                        kindChange();

                        MessageBox.Show("資料已修改!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        sql = "select * from work_hour order by date DESC limit 300 ";
                        dataGridView5.DataSource = GetDataTable(sql);
                        DataTable dt = GetDataTable(sql);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            if (dt.Rows[i][2].ToString() == "3")
                            {
                                dataGridView5.Rows[i].Cells[2].Value = "現金";

                            }
                            else
                            {
                                dataGridView5.Rows[i].Cells[2].Value = "月結";
                            }
                        }

                        textBox43.Text = "";           //將資料放到textBox
                        textBox44.Text = "";

                        textBox48.Text = "";
                        textBox50.Text = "";
                        textBox45.Text = "";
                        textBox47.Text = "";
                        richTextBox16.Text = "";

                        tabControl3.SelectedTab = tabPage15;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < 8; i++)//讀取data,一次讀一格
            {
                MoneyValue[i] = dataGridView4.Rows[e.RowIndex].Cells[i].Value.ToString();
            }
            textBox43.Text = MoneyValue[0];           //將資料放到textBox
            textBox44.Text = MoneyValue[1];

            textBox48.Text = MoneyValue[3];
            textBox50.Text = MoneyValue[4];
            textBox45.Text = MoneyValue[5];
            textBox47.Text = MoneyValue[6];
            richTextBox16.Text = MoneyValue[7];

            if (MoneyValue[2] == "3")       //月結判斷
            {
                radioButton7.Checked = true;
            }
            else
            {
                radioButton8.Checked = true;
            }

            tabControl3.SelectedTab = tabPage11;

        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < 8; i++)//讀取data,一次讀一格
            {
                MoneyValue[i] = dataGridView3.Rows[e.RowIndex].Cells[i].Value.ToString();
            }
            textBox43.Text = MoneyValue[0];           //將資料放到textBox
            textBox44.Text = MoneyValue[1];

            textBox48.Text = MoneyValue[3];
            textBox50.Text = MoneyValue[4];
            textBox45.Text = MoneyValue[5];
            textBox47.Text = MoneyValue[6];
            richTextBox16.Text = MoneyValue[7];

            if (MoneyValue[2] == "3")       //月結判斷
            {
                radioButton7.Checked = true;
            }
            else
            {
                radioButton8.Checked = true;
            }

            tabControl3.SelectedTab = tabPage11;

        }

        private void dataGridView3and4_Layout(object sender, LayoutEventArgs e)//收入列表頁面
        {
            moneyData();

            kindChange();
            textBox12.Text = DateTime.Now.ToString("yyyyMMdd");
        }

        private void button3_Click(object sender, EventArgs e)//新增公司-清除鍵
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            richTextBox5.Text = "";
        }

        private void button10_Click(object sender, EventArgs e)//收入修改-恢復鍵
        {
            textBox43.Text = MoneyValue[0];           
            textBox44.Text = MoneyValue[1];
            textBox48.Text = MoneyValue[3];
            textBox50.Text = MoneyValue[4];
            textBox45.Text = MoneyValue[5];
            textBox47.Text = MoneyValue[6];
            richTextBox16.Text = MoneyValue[7];

            if (MoneyValue[2] == "3")       //月結判斷
            {
                radioButton7.Checked = true;
            }
            else
            {
                radioButton8.Checked = true;
            }

        }

        
        private void dataGridView5_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < 8; i++)//讀取data,一次讀一格
            {
                PrintMoneyValue[i] = dataGridView5.Rows[e.RowIndex].Cells[i].Value.ToString();
            }
          
            string sql = "select * from factory where id = '"+ PrintMoneyValue[1] +"' ";
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < 13; i++)//讀取data,一次讀一格
            {
                PrintFactoryValue[i] = dt.Rows[0][i].ToString();
            }

            label60.Text = PrintFactoryValue[2];
            label61.Text = PrintFactoryValue[4];
            label66.Text = PrintFactoryValue[1];
            label64.Text = PrintFactoryValue[6];

            label78.Text = PrintFactoryValue[2];
            label76.Text = PrintFactoryValue[4];
            label74.Text = PrintFactoryValue[1];
            label72.Text = PrintFactoryValue[6];

            DateTime date = DateTime.Now;
            string ffmonth = date.AddMonths(-2).ToString("yyyyMM");

            ffmonth = ffmonth + "25";

            string fmonth = date.AddMonths(-1).ToString("yyyyMM");

            fmonth = fmonth + "25";


            sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date>='"+ffmonth+"' and date<= '"+fmonth+"' ";

            dt = GetDataTable(sql);
            dataGridView6.DataSource = dt;

            int money=0;

            for(int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString());
            }
            label69.Text = money.ToString();

            sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' ";

            dt = GetDataTable(sql);
            dataGridView7.DataSource = dt;

            money = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString());
            }
            label70.Text = money.ToString();

            tabControl4.SelectedTab = tabPage13;

        }

        private void tabPage12_Layout(object sender, LayoutEventArgs e)
        {
          
        }

        private void button11_Click(object sender, EventArgs e)//列印管理-查詢鍵
        {
            if (textBox51.Text.Trim().Length > 0)
            {
                string sql = "select * from work_hour where f_name like '%" + textBox51.Text + "%'  " +
                                                         "or kind  like '%" + textBox51.Text + "%' " +
                                                         "or carkind  like '%" + textBox51.Text + "%' " +
                                                         "or money  = '%" + textBox51.Text + "%' " +
                                                         "or date  like '%" + textBox51.Text + "%' " +                                                      
                                                         "or other  like '%" + textBox51.Text + "%' ";
                dataGridView5.DataSource = GetDataTable(sql);
            }
        }

        private void tabPage13_Layout(object sender, LayoutEventArgs e)
        {
        }

       
        private void button13_Click(object sender, EventArgs e)
        {
         

            PrintDocument pd = new PrintDocument();
            
            pd.PrintPage += new PrintPageEventHandler(PrintImage);
                    
            PrintDialog p = new PrintDialog();
            p.Document = pd;
            if (DialogResult.OK == p.ShowDialog()) //如果確認，將會覆蓋所有的打印參數設置
            {
                
                PageSetupDialog psd = new PageSetupDialog();
                psd.Document = pd;
                if (DialogResult.OK == psd.ShowDialog())
                {
                    //打印預覽
                    PrintPreviewDialog ppd = new PrintPreviewDialog();
                    ppd.Document = pd;
                    if (DialogResult.OK == ppd.ShowDialog())
                    {
                        pd.Print(); //打印
                    }

                }
            }
        }
        
        void PrintImage(object o , PrintPageEventArgs e)
        {
            int x = SystemInformation.WorkingArea.X;
            int y = SystemInformation.WorkingArea.Y;
            int width = this.Width;
            int height = this.Height;

            Rectangle bounds = new Rectangle(100, 0, 1169, 708);

            Bitmap img = new Bitmap(1200,800)  ;

            this.DrawToBitmap(img, bounds);
            Point p = new Point(0, 0);
            e.Graphics.DrawImage(img, p);
        }
        private void dataGridView7_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
           
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
          
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox40.Text = DateTime.Now.ToString("HHmmss");
            textBox30.Text = "";
            richTextBox15.Text = "";
            textBox49.Text = "";
            textBox41.Text = textBox27.Text;
        }

        private void dataGridView8_Layout(object sender, LayoutEventArgs e)
        {
            
            
        }

        private void dataGridView8_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < 8; i++)//讀取data,一次讀一格
            {
                MoneyValue[i] = dataGridView8.Rows[e.RowIndex].Cells[i].Value.ToString();
            }
            textBox43.Text = MoneyValue[0];           //將資料放到textBox
            textBox44.Text = MoneyValue[1];

            textBox48.Text = MoneyValue[3];
            textBox50.Text = MoneyValue[4];
            textBox45.Text = MoneyValue[5];
            textBox47.Text = MoneyValue[6];
            richTextBox16.Text = MoneyValue[7];

            if (MoneyValue[2] == "3")       //月結判斷
            {
                radioButton7.Checked = true;
            }
            else
            {
                radioButton8.Checked = true;
            }

            tabControl3.SelectedTab = tabPage11;
        }

        private void dataGridView5_Layout(object sender, LayoutEventArgs e)
        {
            string sql = "select * from work_hour order by date DESC  limit 300 ";
            dataGridView5.DataSource = GetDataTable(sql);
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView5.Rows[i].Cells[2].Value = "現金";

                }
                else
                {
                    dataGridView5.Rows[i].Cells[2].Value = "月結";
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            string sql;
            if (textBox7.Text.Trim().Length > 0)
            {
                if (checkBox1.Checked == false)
                {
                    sql = "select * from work_hour where f_name like '%" + textBox7.Text + "%'  " +
                                                           "or kind  like '%" + textBox7.Text + "%' " +
                                                           "or carkind  like '%" + textBox7.Text + "%' " +
                                                           "or money  = '" + textBox7.Text + "' " +
                                                           "or date  like '%" + textBox7.Text + "%' " +
                                                           "or other  like '%" + textBox7.Text + "%' ";
                    dataGridView8.DataSource = GetDataTable(sql);
                }
                else
                {
                    sql = "select * from work_hour where  f_name like '%" + textBox7.Text + "%'  " +
                                                           "or kind  like '%" + textBox7.Text + "%' " +
                                                           "or carkind  like '%" + textBox7.Text + "%' " +
                                                           "or money  = '" + textBox7.Text + "' " +
                                                           "or other  like '%" + textBox7.Text + "%' " +
                                                           "and date >= '" + textBox8.Text + "' and date <= '" + textBox12.Text + "'" ;


                    dataGridView8.DataSource = GetDataTable(sql);
                }
            }
            else
            {
                if (checkBox1.Checked == false)
                {
                    sql = "select * from work_hour order by date DESC";
                    dataGridView8.DataSource = GetDataTable(sql);
                }
                else
                {
                    sql = "select * from work_hour where date >='" + textBox8.Text + "' and date <='" + textBox12.Text + "'";
                                                          
                    dataGridView8.DataSource = GetDataTable(sql);
                }
            }
            DataTable dt = GetDataTable(sql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView8.Rows[i].Cells[2].Value = "現金";

                }
                else
                {
                    dataGridView8.Rows[i].Cells[2].Value = "月結";
                }
            }


        }

        private void dataGridView7_Layout(object sender, LayoutEventArgs e)
        {
            dataGridView6.DataSource = null;
            dataGridView7.DataSource = null;
        }
    }
}
