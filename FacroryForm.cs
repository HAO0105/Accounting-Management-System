using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace Myfactory
{


    public partial class FactoryForm : Form
    {
        public static string connStr = "server=Localhost;database=work;port=3306;username=root;password=123456789;charset=utf8;SSL Mode=None";
        MySqlConnection conn = new MySqlConnection(connStr);
        string[] FactoryValue = new string[13];//存dataGridView1的值
        string[] MoneyValue = new string[8];//存dataGridView3.4.8的值
        string[] NewMoneyValue = new string[13];//存dataGridView2的值
        string[] PrintMoneyValue = new string[8];//存dataGridView5的值
        string[] PrintFactoryValue = new string[13];//

        public FactoryForm()
        {
            InitializeComponent();
            GetPrintName();
        }

        public static DataTable GetDataTable(string sql)
        {
            DataTable dt = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter(sql, connStr);
            adapter.Fill(dt);
            return dt;
        }

        public void moneykindChange()
        {
            string sql = "";
            DateTime date = DateTime.Now;
            string monthDate = date.ToString("yyyyMM");
            monthDate += "01";
            if (radioButton9.Checked)
            {
                sql = "select * from work_hour where date >='" + monthDate + " '  order by date DESC ";
            }
            else if (radioButton10.Checked)
            {
                sql = "select * from work_hour where date >='" + monthDate + " '  and kind = 2 order by date DESC ";
            }
            else if (radioButton11.Checked)
            {
                sql = "select * from work_hour where date >='" + monthDate + " '  and kind = 3 order by date DESC ";
            }
            DataTable dt = GetDataTable(sql);
            dataGridView3.DataSource = dt;

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
            string exmonthDate = date.AddMonths(-1).ToString("yyyyMM");
            exmonthDate += "01";

            if (radioButton14.Checked)
            {
                sql = "select * from work_hour where date >='" + exmonthDate + " '  and date <='" + monthDate + " '  order by date DESC ";
            }
            else if (radioButton13.Checked)
            {
                sql = "select * from work_hour where date >='" + exmonthDate + " '  and date <='" + monthDate + " '  and kind = 2 order by date DESC ";
            }
            else if (radioButton12.Checked)
            {
                sql = "select * from work_hour where date >='" + exmonthDate + " '  and date <='" + monthDate + " '  and kind = 3 order by date DESC ";
            }

            dt = GetDataTable(sql);
            dataGridView4.DataSource = dt;
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
            dataGridView8.DataSource = dt;
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

                string sql = "select * from factory where f_name like '%" + textBox1.Text.Trim() + "%'  " +
                                                         "or name  like '%" + textBox1.Text.Trim() + "%' " +
                                                         "or address  like '%" + textBox1.Text.Trim() + "%' " +
                                                         "or work_place  like '%" + textBox1.Text.Trim() + "%' " +
                                                         "or tel  = '" + textBox1.Text.Trim() + "' " +
                                                         "or work_kind  like '%" + textBox1.Text.Trim() + "%' " +
                                                         "or work_add  like '%" + textBox1.Text.Trim() + "%' " +
                                                         "or other  like '%" + textBox1.Text.Trim() + "%' ";  //模糊查詢公司名稱                
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
            if (MessageBox.Show("確定要新增資料嗎?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

                    MessageBox.Show("公司資料已新增", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

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

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)//收入管理-查詢鍵
        {
            if (textBox2.Text.Trim().Length > 0)
            {
                string sql = "select * from factory where f_name like '%" + textBox2.Text.Trim() + "%'  " +
                                                          "or name  like '%" + textBox2.Text.Trim() + "%' " +
                                                          "or address  like '%" + textBox2.Text.Trim() + "%' " +
                                                          "or work_place  like '%" + textBox2.Text.Trim() + "%' " +
                                                          "or tel  = '" + textBox2.Text.Trim() + "' " +
                                                          "or work_kind  like '%" + textBox2.Text.Trim() + "%' " +
                                                          "or work_add  like '%" + textBox2.Text.Trim() + "%' " +
                                                          "or other  like '%" + textBox2.Text.Trim() + "%' ";  //模糊查詢公司名稱              
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
                        MessageBox.Show("資料已修改!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

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

                        tabControl2.SelectedTab = tabPage4;

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
            try
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
            catch
            {
                MessageBox.Show("請點選公司!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
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

                textBox40.Text = DateTime.Now.ToString("HHmmss");

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
            catch
            {
                MessageBox.Show("請點選公司!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            else if (MessageBox.Show("確定要新增資料嗎?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

                    moneykindChange();

                    sql = "select * from work_hour where kind = 2 order by date DESC limit 300";
                    dataGridView5.DataSource = GetDataTable(sql);
                    DataTable dt = GetDataTable(sql);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                            dataGridView5.Rows[i].Cells[2].Value = "月結";                       
                    }
                    if (textBox39.Text == MoneyValue[1])
                    {
                        sql = "select * from work_hour where id = '" + MoneyValue[1] + "' ";//歷年

                        dt = GetDataTable(sql);
                        dataGridView7.DataSource = dt;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            if (dt.Rows[i][2].ToString() == "3")
                            {
                                dataGridView7.Rows[i].Cells[0].Value = "現金";

                            }
                            else
                            {
                                dataGridView7.Rows[i].Cells[0].Value = "月結";
                            }
                        }


                        int money = 0;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
                        }
                        label70.Text = money.ToString();
                    }
                    if (MoneyValue[1] == textBox39.Text && printclick == false)
                    {
                        data348();
                    }
                    else if (PrintMoneyValue[1] == textBox39.Text && printclick == true)
                    {
                        printdata348();
                    }


                    MessageBox.Show("公司資料已新增", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    textBox41.Text = "";
                    textBox39.Text = "";
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

                    tabControl3.SelectedTab = tabPage15;
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
                MessageBox.Show("請先至 [查詢收入]頁面 選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                                                        "where id = '" + MoneyValue[1] + "' and f_name = '" + MoneyValue[0] + "' and time='" + MoneyValue[3] + "' " +
                                                        "and carkind ='" + MoneyValue[4] + "'  and money = '" + MoneyValue[5] + "'and date ='" + MoneyValue[6] + "'" +
                                                        "and other = '" + MoneyValue[7] + "'  ";
                        GetDataTable(sql);

                        moneykindChange();

                        MessageBox.Show("資料已修改!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        sql = "select * from work_hour where kind = 2 order by date DESC limit 300 ";
                        dataGridView5.DataSource = GetDataTable(sql);
                        DataTable dt = GetDataTable(sql);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                                dataGridView5.Rows[i].Cells[2].Value = "月結";                          
                        }

                        if (MoneyValue[1] == textBox44.Text && printclick == false)
                        {
                            data348();
                        }
                        else if (PrintMoneyValue[1] == textBox44.Text && printclick == true)
                        {
                            printdata348();
                        }

                        totaldata();

                        textBox43.Text = "";
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

        private void dataGridView3and4_Layout(object sender, LayoutEventArgs e)//收入列表頁面
        {
            moneykindChange();
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
            try
            {
                for (int i = 0; i < 8; i++)//讀取data,一次讀一格
                {
                    PrintMoneyValue[i] = dataGridView5.Rows[e.RowIndex].Cells[i].Value.ToString();
                }
                printdata348();
                tabControl4.SelectedTab = tabPage16;
            }
            catch
            {
                MessageBox.Show("請點選公司!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }

        private void button11_Click(object sender, EventArgs e)//列印管理-查詢鍵
        {
            string sql;
            if (textBox51.Text.Trim().Length > 0)
            {
                if (checkBox2.Checked == false)
                {
                    sql = "select * from work_hour where (f_name like '%" + textBox51.Text.Trim() + "%'  " +
                                                           "or kind  like '%" + textBox51.Text.Trim() + "%' " +
                                                           "or carkind  like '%" + textBox51.Text.Trim() + "%' " +
                                                           "or money  = '" + textBox51.Text.Trim() + "' " +
                                                           "or date  like '%" + textBox51.Text.Trim() + "%' " +
                                                           "or other  like '%" + textBox51.Text.Trim() + "%') and kind = 2 ";
                    dataGridView5.DataSource = GetDataTable(sql);
                }
                else
                {
                    sql = "select * from work_hour where  (f_name like '%" + textBox51.Text.Trim() + "%'  " +
                                                           "or kind  like '%" + textBox51.Text.Trim() + "%' " +
                                                           "or carkind  like '%" + textBox51.Text.Trim() + "%' " +
                                                           "or money  = '" + textBox51.Text.Trim() + "' " +
                                                           "or other  like '%" + textBox51.Text.Trim() + "%' ) " +
                                                           "and date >= '" + textBox15.Text + "' and date <= '" + textBox14.Text + "' and kind = 2 ";


                    dataGridView5.DataSource = GetDataTable(sql);
                }
            }
            else
            {
                if (checkBox2.Checked == false)
                {
                    sql = "select * from work_hour where kind = 2 order by date DESC";
                    dataGridView5.DataSource = GetDataTable(sql);
                }
                else
                {
                    sql = "select * from work_hour where date >='" + textBox15.Text + "' and date <='" + textBox14.Text + "' and kind = 2 ";

                    dataGridView5.DataSource = GetDataTable(sql);
                }
            }
            DataTable dt = GetDataTable(sql);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                    dataGridView5.Rows[i].Cells[2].Value = "月結";               
            }
        }
        public void print2(PrintPageEventHandler PP,string choose)
        {
            PrinterSettings ps = new PrinterSettings();
            PrintDocument pd = new PrintDocument();
            pd.PrinterSettings = ps;
            IEnumerable<PaperSize> paperSizes = ps.PaperSizes.Cast<PaperSize>();
            PaperSize sizeA4 = paperSizes.First<PaperSize>(size => size.Kind == PaperKind.A4);
            pd.DefaultPageSettings.PaperSize = sizeA4;
            try
            {
                Font printFont = new Font("細明體", 10);
                Font titleFont = new Font("細明體", 15);
                pd.PrinterSettings.PrinterName = choose ;
                pd.DocumentName = pd.PrinterSettings.MaximumCopies.ToString();

                //pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.pd_PrintPage);
                pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PP);
                pd.PrintController = new System.Drawing.Printing.StandardPrintController();

                PrintPreviewDialog ppd = new PrintPreviewDialog();
                ppd.Document = pd;
                PrintDialog p = new PrintDialog();
                p.Document = pd;

                if (DialogResult.OK == ppd.ShowDialog())
                {

                    pd.Print();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        string choosef30="";
        string choosef25 = "";
        string choose30 = "";
        string choose25 = "";
        private void GetPrintName()
        {
       
            PrintDocument print = new PrintDocument();
            string sDefault = print.PrinterSettings.PrinterName;//默認印機名

            label100.Text = sDefault;
            label148.Text = sDefault;
            label149.Text = sDefault;
            label150.Text = sDefault;
            choosef30 = sDefault;
            choosef25 = sDefault;
            choose30 = sDefault;
            choose25 = sDefault;

            foreach (string sPrint in PrinterSettings.InstalledPrinters)//取所有印機名稱
            {
                listBox1.Items.Add(sPrint);
                listBox2.Items.Add(sPrint);
                listBox3.Items.Add(sPrint);
                listBox4.Items.Add(sPrint);

                if (sPrint == sDefault)
                {
                    listBox1.SelectedIndex = listBox1.Items.IndexOf(sPrint);
                    listBox2.SelectedIndex = listBox2.Items.IndexOf(sPrint);
                    listBox3.SelectedIndex = listBox3.Items.IndexOf(sPrint);
                    listBox4.SelectedIndex = listBox4.Items.IndexOf(sPrint);
                }


            }
        }
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)//本月月結帳單30
        {
            e.Graphics.DrawString("立晟起重公司", new Font(new FontFamily("細明體"), 25), System.Drawing.Brushes.Black, 300, 30);
            e.Graphics.DrawString("新北市樹林區俊英街219巷32號", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 303, 70);
            e.Graphics.DrawString("電話: 8688-8181 , 8688-5252    傳真:8688-3852", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 250, 88);
            e.Graphics.DrawString(label124.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 315, 130);
            e.Graphics.DrawString("公司:" + label105.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 180);
            e.Graphics.DrawString("地址:" + label103.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 210);
            e.Graphics.DrawString("統一編號:" + label99.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 180);
            e.Graphics.DrawString("聯絡電話:" + label97.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 210);
            e.Graphics.DrawLine(Pens.Black, 8, 235, 810, 235);
            e.Graphics.DrawString("付款方式", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 100, 250);
            e.Graphics.DrawString("工作日期", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 350, 250);
            e.Graphics.DrawString("付款金額", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 600, 250);

            int x = 100, y = 280;

            string month30;//30
            string fmonth30;
            month30 = DateTime.Now.ToString("yyyyMM") + "31";
            fmonth30 = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "31";

            string sql;
            if (printclick == true)
            {
                sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date >'" + fmonth30 + "' and date<= '" + month30 + "' and kind = 2 ";
            }
            else
            {
                sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date >'" + fmonth30 + "' and date<= '" + month30 + "' and kind = 2 ";
            }
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0) continue;
                    if (j == 1) continue;
                    if (j == 3) continue;
                    if (j == 4) continue;
                    if (j == 7) continue;
                    if (j == 2)
                    {
                        if (dt.Rows[i][2].ToString() == "2")
                        {
                            e.Graphics.DrawString("月結", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                        else
                        {
                            e.Graphics.DrawString("現金", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                    }
                    else if (j == 5)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j + 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }
                    else if (j == 6)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j - 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }

                    x = x + 250;
                }
                x = 100;
                y = y + 30;
                e.Graphics.DrawLine(Pens.Black, 8, 1040, 810, 1040);
                e.Graphics.DrawString("總金額(未稅):" + label115.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1060);
                e.Graphics.DrawString("       稅額 : " + label112.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1090);
                e.Graphics.DrawString("總金額(含稅):" + label110.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1120);
                e.Graphics.DrawString("以上內容如有疑問，請於收到帳單五日內，與本公司聯絡，逾期恕不更改。" , new Font(new FontFamily("細明體"), 8), System.Drawing.Brushes.Black, 225, 1150);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox40.Text = DateTime.Now.ToString("HHmmss");
            textBox30.Text = "";
            richTextBox15.Text = "";
            textBox49.Text = "";
            textBox41.Text = textBox27.Text;
        }

        private void dataGridView5_Layout(object sender, LayoutEventArgs e)
        {
            textBox14.Text = DateTime.Now.ToString("yyyyMMdd");
            string sql = "select * from work_hour where kind = 2 order by date DESC  limit 300 ";
            dataGridView5.DataSource = GetDataTable(sql);
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                    dataGridView5.Rows[i].Cells[2].Value = "月結";               
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string sql;
            if (textBox7.Text.Trim().Length > 0)
            {
                if (checkBox1.Checked == false)
                {
                    sql = "select * from work_hour where f_name like '%" + textBox7.Text.Trim() + "%'  " +
                                                           "or kind  like '%" + textBox7.Text.Trim() + "%' " +
                                                           "or carkind  like '%" + textBox7.Text.Trim() + "%' " +
                                                           "or money  = '" + textBox7.Text.Trim() + "' " +
                                                           "or date  like '%" + textBox7.Text.Trim() + "%' " +
                                                           "or other  like '%" + textBox7.Text.Trim() + "%' ";
                    dataGridView8.DataSource = GetDataTable(sql);
                }
                else
                {
                    sql = "select * from work_hour where  (f_name like '%" + textBox7.Text.Trim() + "%'  " +
                                                           "or kind  like '%" + textBox7.Text.Trim() + "%' " +
                                                           "or carkind  like '%" + textBox7.Text.Trim() + "%' " +
                                                           "or money  = '" + textBox7.Text.Trim() + "' " +
                                                           "or other  like '%" + textBox7.Text.Trim() + "%' ) " +
                                                           "and date >= '" + textBox8.Text.Trim() + "' and date <= '" + textBox12.Text.Trim() + "'";

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
                    sql = "select * from work_hour where date >='" + textBox8.Text.Trim() + "' and date <='" + textBox12.Text.Trim() + "'";

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
        private void printDocument2_PrintPage(object sender, PrintPageEventArgs e)//上月收入列印30
        {
            e.Graphics.DrawString("立晟起重公司", new Font(new FontFamily("細明體"), 25), System.Drawing.Brushes.Black, 300, 30);
            e.Graphics.DrawString("新北市樹林區俊英街219巷32號", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 303, 70);
            e.Graphics.DrawString("電話: 8688-8181 , 8688-5252    傳真:8688-3852", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 250, 88);
            e.Graphics.DrawString(label106.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 315, 130);
            e.Graphics.DrawString("公司:" + label105.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 180);
            e.Graphics.DrawString("地址:" + label103.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 210);
            e.Graphics.DrawString("統一編號:" + label99.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 180);
            e.Graphics.DrawString("聯絡電話:" + label97.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 210);
            e.Graphics.DrawLine(Pens.Black, 8, 235, 810, 235);
            e.Graphics.DrawString("付款方式", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 100, 250);
            e.Graphics.DrawString("工作日期", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 350, 250);
            e.Graphics.DrawString("付款金額", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 600, 250);

            int x = 100, y = 280;

            string ffmonth30;//30
            string fmonth30;
            ffmonth30 = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "01";
            fmonth30 = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "31";

            string sql;
            if (printclick)
            {
                sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date >='" + ffmonth30 + "' and date<= '" + fmonth30 + "' and kind = 2 ";
            }
            else
            {
                sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date >='" + ffmonth30 + "' and date<= '" + fmonth30 + "' and kind = 2 ";
            }
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0) continue;
                    if (j == 1) continue;
                    if (j == 3) continue;
                    if (j == 4) continue;
                    if (j == 7) continue;
                    if (j == 2)
                    {
                        if (dt.Rows[i][2].ToString() == "2")
                        {
                            e.Graphics.DrawString("月結", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                        else
                        {
                            e.Graphics.DrawString("現金", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                    }
                    else if (j == 5)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j + 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }
                    else if (j == 6)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j - 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }

                    x = x + 250;
                }
                x = 100;
                y = y + 30;
                e.Graphics.DrawLine(Pens.Black, 8, 1040, 810, 1040);
                e.Graphics.DrawString("總金額(未稅):" + label95.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1060);
                e.Graphics.DrawString("       稅額 : " + label92.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1090);
                e.Graphics.DrawString("總金額(含稅):" + label90.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1120);
                e.Graphics.DrawString("以上內容如有疑問，請於收到帳單五日內，與本公司聯絡，逾期恕不更改。", new Font(new FontFamily("細明體"), 8), System.Drawing.Brushes.Black, 225, 1150);
            }
        }
        private void button15_Click(object sender, EventArgs e)
        {
            if (label105.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                print2(printDocument2_PrintPage,choosef30);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (label105.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                print2(printDocument3_PrintPage,choosef25);
            }
        }

        private void printDocument3_PrintPage(object sender, PrintPageEventArgs e)//上月收入列印25.26
        {
            e.Graphics.DrawString("立晟起重公司", new Font(new FontFamily("細明體"), 25), System.Drawing.Brushes.Black, 300, 30);
            e.Graphics.DrawString("新北市樹林區俊英街219巷32號", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 303, 70);
            e.Graphics.DrawString("電話: 8688-8181 , 8688-5252    傳真:8688-3852", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 250, 88);
            e.Graphics.DrawString(label59.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 270, 130);
            e.Graphics.DrawString("公司:" + label105.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 180);
            e.Graphics.DrawString("地址:" + label103.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 210);
            e.Graphics.DrawString("統一編號:" + label99.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 180);
            e.Graphics.DrawString("聯絡電話:" + label97.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 210);
            e.Graphics.DrawLine(Pens.Black, 8, 235, 810, 235);
            e.Graphics.DrawString("付款方式", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 100, 250);
            e.Graphics.DrawString("工作日期", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 350, 250);
            e.Graphics.DrawString("付款金額", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 600, 250);

            int x = 100, y = 280;

            string ffmonth25;//25.26
            string fmonth25;
            ffmonth25 = DateTime.Now.AddMonths(-2).ToString("yyyyMM") + "26";
            fmonth25 = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "25";

            string sql;
            if (printclick)
            {
                sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date >='" + ffmonth25 + "' and date<= '" + fmonth25 + "' and kind = 2 ";
            }
            else
            {
                sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date >='" + ffmonth25 + "' and date<= '" + fmonth25 + "' and kind = 2 ";
            }
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0) continue;
                    if (j == 1) continue;
                    if (j == 3) continue;
                    if (j == 4) continue;
                    if (j == 7) continue;


                    if (j == 2)
                    {
                        if (dt.Rows[i][2].ToString() == "2")
                        {
                            e.Graphics.DrawString("月結", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                        else
                        {
                            e.Graphics.DrawString("現金", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                    }
                    else if (j == 5)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j + 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }
                    else if (j == 6)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j - 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }

                    x = x + 250;

                }
                x = 100;
                y = y + 30;
                e.Graphics.DrawLine(Pens.Black, 8, 1040, 810, 1040);
                e.Graphics.DrawString("總金額(未稅):" + label69.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1060);
                e.Graphics.DrawString("       稅額 : " + label85.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1090);
                e.Graphics.DrawString("總金額(含稅):" + label88.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1120);
                e.Graphics.DrawString("以上內容如有疑問，請於收到帳單五日內，與本公司聯絡，逾期恕不更改。", new Font(new FontFamily("細明體"), 8), System.Drawing.Brushes.Black, 225, 1150);
            }
        }

        void totaldata()
        {
            string sql = "select * from factory where id = '" + MoneyValue[1] + "' ";//放歷年
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < 13; i++)//讀取data,一次讀一格
            {
                PrintFactoryValue[i] = dt.Rows[0][i].ToString();
            }

            label78.Text = PrintFactoryValue[2];//歷年
            label76.Text = PrintFactoryValue[4];
            label74.Text = PrintFactoryValue[1];
            label72.Text = PrintFactoryValue[6];

            sql = "select * from work_hour where id = '" + MoneyValue[1] + "' ";//歷年

            dt = GetDataTable(sql);
            dataGridView7.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView7.Rows[i].Cells[0].Value = "現金";
                }
                else
                {
                    dataGridView7.Rows[i].Cells[0].Value = "月結";
                }
            }

            int money = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label70.Text = money.ToString();
        }
        Boolean printclick;
        void data348()
        {
            printclick = false;
            string sql = "select * from factory where id = '" + MoneyValue[1] + "' ";
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < 13; i++)//讀取data,一次讀一格
            {
                PrintFactoryValue[i] = dt.Rows[0][i].ToString();
            }

            label60.Text = PrintFactoryValue[2];//上月月結25.26
            label61.Text = PrintFactoryValue[4];
            label66.Text = PrintFactoryValue[1];
            label64.Text = PrintFactoryValue[6];

            label105.Text = PrintFactoryValue[2];//上月月結30
            label103.Text = PrintFactoryValue[4];
            label99.Text = PrintFactoryValue[1];
            label97.Text = PrintFactoryValue[6];

            label123.Text = PrintFactoryValue[2];//本月月結30
            label121.Text = PrintFactoryValue[4];
            label119.Text = PrintFactoryValue[1];
            label117.Text = PrintFactoryValue[6];

            label141.Text = PrintFactoryValue[2];//本月月結25.26
            label139.Text = PrintFactoryValue[4];
            label137.Text = PrintFactoryValue[1];
            label135.Text = PrintFactoryValue[6];

            DateTime date = DateTime.Now;

            string ffmonth25;//25.26
            string fmonth25;
            string month25;

            ffmonth25 = date.AddMonths(-2).ToString("yyyyMM") + "26";
            fmonth25 = date.AddMonths(-1).ToString("yyyyMM") + "25";
            month25 = date.ToString("yyyyMM") + "25";

            sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date>='" + ffmonth25 + "' and date<= '" + fmonth25 + "' and kind = 2 ";//上月25.26

            dt = GetDataTable(sql);
            dataGridView6.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView6.Rows[i].Cells[0].Value = "現金";

                }
                else
                {
                    dataGridView6.Rows[i].Cells[0].Value = "月結";
                }
            }

            int money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString());
            }
            label69.Text = money.ToString();
            int tax = Convert.ToInt32(money * 0.05);
            label85.Text = "" + tax.ToString();
            int total = money + tax;
            label88.Text = total.ToString();

            string ffmonth30;//30
            string fmonth30;
            string month30;

            ffmonth30 = date.AddMonths(-1).ToString("yyyyMM") + "01";
            fmonth30 = date.AddMonths(-1).ToString("yyyyMM") + "31";
            month30 = date.ToString("yyyyMM") + "31";

            sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date >='" + ffmonth30 + "' and date<= '" + fmonth30 + "' and kind = 2 ";//上月30
            dt = GetDataTable(sql);
            dataGridView9.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView9.Rows[i].Cells[0].Value = "現金";

                }
                else
                {
                    dataGridView9.Rows[i].Cells[0].Value = "月結";
                }
            }

            money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label95.Text = money.ToString();
            tax = Convert.ToInt32(money * 0.05);
            label92.Text = "" + tax.ToString();
            total = money + tax;
            label90.Text = total.ToString();

            sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date>'" + fmonth25 + "' and date<= '" + month25 + "' and kind = 2 ";//本月25.26

            dt = GetDataTable(sql);
            dataGridView11.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView11.Rows[i].Cells[0].Value = "現金";

                }
                else
                {
                    dataGridView11.Rows[i].Cells[0].Value = "月結";
                }
            }

            money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label133.Text = money.ToString();
            tax = Convert.ToInt32(money * 0.05);
            label130.Text = "" + tax.ToString();
            total = money + tax;
            label128.Text = total.ToString();

            sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date >'" + fmonth30 + "' and date<= '" + month30 + "' and kind = 2 ";//本月30
            dt = GetDataTable(sql);
            dataGridView10.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView10.Rows[i].Cells[0].Value = "現金";

                }
                else
                {
                    dataGridView10.Rows[i].Cells[0].Value = "月結";
                }
            }

            money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label115.Text = money.ToString();
            tax = Convert.ToInt32(money * 0.05);
            label112.Text = "" + tax.ToString();
            total = money + tax;
            label110.Text = total.ToString();

        }
        void printdata348()
        {
            printclick = true;
            string sql = "select * from factory where id = '" + PrintMoneyValue[1] + "' ";
            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < 13; i++)//讀取data,一次讀一格
            {
                PrintFactoryValue[i] = dt.Rows[0][i].ToString();
            }

            label60.Text = PrintFactoryValue[2];//上月月結25.26
            label61.Text = PrintFactoryValue[4];
            label66.Text = PrintFactoryValue[1];
            label64.Text = PrintFactoryValue[6];

            label105.Text = PrintFactoryValue[2];//上月月結30
            label103.Text = PrintFactoryValue[4];
            label99.Text = PrintFactoryValue[1];
            label97.Text = PrintFactoryValue[6];

            label123.Text = PrintFactoryValue[2];//本月月結30
            label121.Text = PrintFactoryValue[4];
            label119.Text = PrintFactoryValue[1];
            label117.Text = PrintFactoryValue[6];

            label141.Text = PrintFactoryValue[2];//本月月結25.26
            label139.Text = PrintFactoryValue[4];
            label137.Text = PrintFactoryValue[1];
            label135.Text = PrintFactoryValue[6];

            DateTime date = DateTime.Now;
            string ffmonth25;//25.26
            string fmonth25;
            string month25;

            ffmonth25 = date.AddMonths(-2).ToString("yyyyMM") + "26";
            fmonth25 = date.AddMonths(-1).ToString("yyyyMM") + "25";
            month25 = date.ToString("yyyyMM") + "25";

            sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date>='" + ffmonth25 + "' and date<= '" + fmonth25 + "' and kind = 2 ";//上月25.26

            dt = GetDataTable(sql);
            dataGridView6.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView6.Rows[i].Cells[0].Value = "現金";
                }
                else
                {
                    dataGridView6.Rows[i].Cells[0].Value = "月結";
                }
            }

            int money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString());
            }
            label69.Text = money.ToString();
            int tax = Convert.ToInt32(money * 0.05);
            label85.Text = "" + tax.ToString();
            int total = money + tax;
            label88.Text = total.ToString();

            string ffmonth30;//30
            string fmonth30;
            string month30;

            ffmonth30 = date.AddMonths(-1).ToString("yyyyMM") + "01";
            fmonth30 = date.AddMonths(-1).ToString("yyyyMM") + "31";
            month30 = date.ToString("yyyyMM") + "31";

            sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date >='" + ffmonth30 + "' and date<= '" + fmonth30 + "' and kind = 2 ";//上月30
            dt = GetDataTable(sql);
            dataGridView9.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView9.Rows[i].Cells[0].Value = "現金";
                }
                else
                {
                    dataGridView9.Rows[i].Cells[0].Value = "月結";
                }
            }

            money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label95.Text = money.ToString();
            tax = Convert.ToInt32(money * 0.05);
            label92.Text = "" + tax.ToString();
            total = money + tax;
            label90.Text = total.ToString();

            sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date>'" + fmonth25 + "' and date<= '" + month25 + "' and kind = 2 ";//本月25.26

            dt = GetDataTable(sql);
            dataGridView11.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView11.Rows[i].Cells[0].Value = "現金";
                }
                else
                {
                    dataGridView11.Rows[i].Cells[0].Value = "月結";
                }
            }

            money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label133.Text = money.ToString();
            tax = Convert.ToInt32(money * 0.05);
            label130.Text = "" + tax.ToString();
            total = money + tax;
            label128.Text = total.ToString();

            sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date >'" + fmonth30 + "' and date<= '" + month30 + "' and kind = 2 ";//本月30
            dt = GetDataTable(sql);
            dataGridView10.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][2].ToString() == "3")
                {
                    dataGridView10.Rows[i].Cells[0].Value = "現金";
                }
                else
                {
                    dataGridView10.Rows[i].Cells[0].Value = "月結";
                }
            }

            money = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                money = money + int.Parse(dt.Rows[i][5].ToString().Trim());
            }
            label115.Text = money.ToString();
            tax = Convert.ToInt32(money * 0.05);
            label112.Text = "" + tax.ToString();
            total = money + tax;
            label110.Text = total.ToString();
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
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

                totaldata();
                tabControl3.SelectedTab = tabPage11;
            }
            catch
            {
                MessageBox.Show("請點選公司!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
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

                totaldata();
                tabControl3.SelectedTab = tabPage11;
            }
            catch
            {
                MessageBox.Show("請點選公司!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView8_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
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

                totaldata();

                tabControl3.SelectedTab = tabPage11;
            }
            catch
            {
                MessageBox.Show("請點選公司!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (label105.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                print2(printDocument1_PrintPage,choose30);
            }
        }

        private void printDocument4_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString("立晟起重公司", new Font(new FontFamily("細明體"), 25), System.Drawing.Brushes.Black, 300, 30);
            e.Graphics.DrawString("新北市樹林區俊英街219巷32號", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 303, 70);
            e.Graphics.DrawString("電話: 8688-8181 , 8688-5252    傳真:8688-3852", new Font(new FontFamily("細明體"), 11), System.Drawing.Brushes.Black, 250, 88);
            e.Graphics.DrawString(label142.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 270, 130);
            e.Graphics.DrawString("公司:" + label105.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 180);
            e.Graphics.DrawString("地址:" + label103.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 20, 210);
            e.Graphics.DrawString("統一編號:" + label99.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 180);
            e.Graphics.DrawString("聯絡電話:" + label97.Text, new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 610, 210);
            e.Graphics.DrawLine(Pens.Black, 8, 235, 810, 235);
            e.Graphics.DrawString("付款方式", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 100, 250);
            e.Graphics.DrawString("工作日期", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 350, 250);
            e.Graphics.DrawString("付款金額", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, 600, 250);

            int x = 100, y = 280;

            string month25;//25.26
            string fmonth25;
            month25 = DateTime.Now.ToString("yyyyMM") + "26";
            fmonth25 = DateTime.Now.AddMonths(-1).ToString("yyyyMM") + "25";

            String sql;
            if (printclick)
            {
                sql = "select * from work_hour where id = '" + PrintMoneyValue[1] + "' and date >'" + fmonth25 + "' and date< '" + month25 + "' and kind = 2 ";
            }
            else
            {
                sql = "select * from work_hour where id = '" + MoneyValue[1] + "' and date >'" + fmonth25 + "' and date< '" + month25 + "' and kind = 2 ";
            }

            DataTable dt = GetDataTable(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0) continue;
                    if (j == 1) continue;
                    if (j == 3) continue;
                    if (j == 4) continue;
                    if (j == 7) continue;
                    if (j == 2)
                    {
                        if (dt.Rows[i][2].ToString() == "2")
                        {
                            e.Graphics.DrawString("月結", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                        else
                        {
                            e.Graphics.DrawString("現金", new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                        }
                    }
                    else if (j == 5)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j + 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }
                    else if (j == 6)
                    {
                        e.Graphics.DrawString(dt.Rows[i][j - 1].ToString(), new Font(new FontFamily("細明體"), 14), System.Drawing.Brushes.Black, x, y);
                    }

                    x = x + 250;

                }
                x = 100;
                y = y + 30;
                e.Graphics.DrawLine(Pens.Black, 8, 1040, 810, 1040);
                e.Graphics.DrawString("總金額(未稅):" + label133.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1060);
                e.Graphics.DrawString("       稅額 : " + label130.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1090);
                e.Graphics.DrawString("總金額(含稅):" + label128.Text, new Font(new FontFamily("細明體"), 15), System.Drawing.Brushes.Black, 560, 1120);
                e.Graphics.DrawString("以上內容如有疑問，請於收到帳單五日內，與本公司聯絡，逾期恕不更改。", new Font(new FontFamily("細明體"), 8), System.Drawing.Brushes.Black, 225, 1150);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (label105.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                print2(printDocument4_PrintPage,choose25);
            }
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            if (label78.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                data348();
                tabControl1.SelectedTab = tabPage3;
                tabControl4.SelectedTab = tabPage16;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (label78.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                data348();
                tabControl1.SelectedTab = tabPage3;
                tabControl4.SelectedTab = tabPage14;
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (label78.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                data348();
                tabControl1.SelectedTab = tabPage3;
                tabControl4.SelectedTab = tabPage17;
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (label78.Text == "尚未選取公司")
            {
                MessageBox.Show("請先選取公司", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                data348();
                tabControl1.SelectedTab = tabPage3;
                tabControl4.SelectedTab = tabPage18;
            }
        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)//上月收入列表選擇付款方式
        {
            string sql = "";
            DateTime date = DateTime.Now;
            string monthDate = date.ToString("yyyyMM");
            monthDate += "01";

            string exmonthDate = date.AddMonths(-1).ToString("yyyyMM");
            exmonthDate += "01";

            if (radioButton14.Checked)
            {
                sql = "select * from work_hour where date >='" + exmonthDate + " '  and date <='" + monthDate + " '  order by date DESC ";
            }
            else if (radioButton13.Checked)
            {
                sql = "select * from work_hour where date >='" + exmonthDate + " '  and date <='" + monthDate + " '  and kind = 2 order by date DESC ";
            }
            else if (radioButton12.Checked)
            {
                sql = "select * from work_hour where date >='" + exmonthDate + " '  and date <='" + monthDate + " '  and kind = 3 order by date DESC ";
            }

            DataTable dt = GetDataTable(sql);
            dataGridView4.DataSource = dt;
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
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)//本月收入列表選擇付款方式
        {
            string sql = "";
            DateTime date = DateTime.Now;
            string monthDate = date.ToString("yyyyMM");
            monthDate += "01";
            if (radioButton9.Checked)
            {
                sql = "select * from work_hour where date >='" + monthDate + " '  order by date DESC ";
            }
            else if (radioButton10.Checked)
            {
                sql = "select * from work_hour where date >='" + monthDate + " '  and kind = 2 order by date DESC ";
            }
            else if (radioButton11.Checked)
            {
                sql = "select * from work_hour where date >='" + monthDate + " '  and kind = 3 order by date DESC ";
            }
            DataTable dt = GetDataTable(sql);
            dataGridView3.DataSource = dt;

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
        }

        private void tabPage16_Layout(object sender, LayoutEventArgs e)//開啟列印管理時
        {
            //設定帳單日期
            label106.Text = DateTime.Now.AddMonths(-1).ToString("yyyy 年 MM")+" 月份帳單";
            label59.Text = DateTime.Now.AddMonths(-2).ToString("yyyy/MM") + "/26 ~ " + DateTime.Now.AddMonths(-1).ToString("yyyy/MM") + "/25 帳單";
            label124.Text = DateTime.Now.ToString("yyyy 年 MM") + " 月份帳單";
            label142.Text = DateTime.Now.AddMonths(-1).ToString("yyyy/MM") + "/26 ~ " + DateTime.Now.ToString("yyyy/MM") + "/25 帳單";
        }

        private void listBox1_Click(object sender, EventArgs e)
        {            
            choosef30 = listBox1.SelectedItem.ToString();
        }

        private void listBox2_Click(object sender, EventArgs e)
        {
            choosef25 = listBox2.SelectedItem.ToString();
        }

        private void listBox3_Click(object sender, EventArgs e)
        {
            choose30 = listBox3.SelectedItem.ToString();
        }

        private void listBox4_Click(object sender, EventArgs e)
        {
            choose25 = listBox4.SelectedItem.ToString();
        }
    }
}