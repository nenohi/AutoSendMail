using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace AutoSendMail
{
    public partial class AutoSendMail : Form
    {
        [DllImport("USER32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void SetCursorPos(int X, int Y);

        [DllImport("USER32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;



        public string DateFormat = "yyyy年MM月dd日（dddd）";
        public bool SendTest = true;
        public Point MousePoint;
        Keys Keys;
        public AutoSendMail()
        {
            InitializeComponent();
            LoadProperties();
            Application.Idle += new EventHandler(Application_Idle);
        }

        private void SendMailTest_Click(object sender, EventArgs e)
        {
            SendMailTest.Enabled = false;
            if (SmptServer.Text == "" || SmptServerPort.Text == "" || SettingUserName.Text == ""||SettingUserPassword.Text=="")
            {
                MessageBox.Show("設定画面を入力してください。");
                return;
            }
            var host = SmptServer.Text;
            var portstr = SmptServerPort.Text;
            int port = Convert.ToInt32(portstr);
            using (var smtp = new MailKit.Net.Smtp.SmtpClient())
            {
                try
                {
                    //開発用のSMTPサーバが暗号化に対応していないときは、次の行をコメントアウト
                    if (SSLConnection.Checked)
                    {
                        smtp.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    }
                    smtp.Connect(host, port, MailKit.Security.SecureSocketOptions.Auto);
                    //認証設定
                    smtp.Authenticate(SettingUserName.Text, SettingUserPassword.Text);

                    //送信するメールを作成する
                    var mail = new MimeKit.MimeMessage();
                    var builder = new MimeKit.BodyBuilder();
                    mail.From.Add(new MimeKit.MailboxAddress(SendUserName.Text, SettingUserName.Text));
                    mail.To.Add(new MimeKit.MailboxAddress(SendUserName.Text,SettingUserName.Text));
                    mail.Subject = "TestSendMail";
                    MimeKit.TextPart textPart = new MimeKit.TextPart("Plain");
                    textPart.Text = "Send By AutoSendMail Tool";
                    var multipart = new MimeKit.Multipart("mixed");
                    multipart.Add(textPart);
                    mail.Body = multipart;
                    //メールを送信する
                    smtp.Send(mail);
                    MessageBox.Show("送信しました。");
                    SendTest = true;
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                }
                finally
                {
                    //SMTPサーバから切断する
                    smtp.Disconnect(true);
                    SendMailTest.Enabled = true;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                SettingUserPassword.PasswordChar = '*';
            }
            else
            {
                SettingUserPassword.PasswordChar = '\0';
            }
        }

        private void SendDataAdd_Click(object sender, EventArgs e)
        {
            if (DateTime.Now.AddDays(-1) > SendData.Value) return;
            foreach (object item in SendDataListBox.Items)
            {
                if (item.ToString() == SendData.Value.ToString(DateFormat))
                {
                    SendDataListBox.SelectedItem = item;
                    return;
                }
            }
            SendDataListBox.Items.Add(SendData.Value.ToString(DateFormat));
            for(int i = 0; i < SendDataListBox.Items.Count; i++)
            {
                DateTime date1 = DateTime.ParseExact(SendDataListBox.Items[i].ToString(), DateFormat, null);
                for(int n = i+1; n < SendDataListBox.Items.Count; n++)
                {
                    DateTime date2 = DateTime.ParseExact(SendDataListBox.Items[n].ToString(), DateFormat, null);
                    if (date1 > date2)
                    {
                        object tmp = SendDataListBox.Items[i];
                        SendDataListBox.Items[i] = SendDataListBox.Items[n];
                        SendDataListBox.Items[n] = tmp;
                    }
                }
            }
        }

        private void SendDataRemove_Click(object sender, EventArgs e)
        {
            if(SendDataListBox.SelectedIndex >= 0)
            {
                SendDataListBox.Items.RemoveAt(SendDataListBox.SelectedIndex);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void SettingSave_Click(object sender, EventArgs e)
        {
            Save();
        }
        public void Save()
        {
            Properties.Settings.Default.SendMailAddress = SendEmail.Text;
            Properties.Settings.Default.SendUserName = SendUserName.Text;
            Properties.Settings.Default.SMTPServer = SmptServer.Text;
            Properties.Settings.Default.SMTPServerPort = SmptServerPort.Text;
            Properties.Settings.Default.UserName = SettingUserName.Text;
            Properties.Settings.Default.UserPass = System.Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes(SettingUserPassword.Text));
            Properties.Settings.Default.IsSSL = SSLConnection.Checked;
            Properties.Settings.Default.Send1CC = SendMessage1CC.Text;
            Properties.Settings.Default.Send1To = SendMessage1To.Text;
            Properties.Settings.Default.Send1Title = SendMessage1Subject.Text;
            Properties.Settings.Default.Send1Text = SendMessage1Text.Text;
            Properties.Settings.Default.Send2To = SendMessage2To.Text;
            Properties.Settings.Default.Send2CC = SendMessage2CC.Text;
            Properties.Settings.Default.Send2Title = SendMessage2Subject.Text;
            Properties.Settings.Default.Send2Text = SendMessage2Text.Text;
            var list = new ArrayList();
            foreach (var item in SendDataListBox.Items)
            {
                list.Add(item.ToString());
            }
            Properties.Settings.Default.DataList = list;
            Properties.Settings.Default.SendMessageTime1 = Send1Time.Value;
            Properties.Settings.Default.SendMessageTime2 = Send2Time.Value;
            Properties.Settings.Default.Macro = checkBox2.Checked;
            Properties.Settings.Default.Save();
        }
        public void LoadProperties()
        {
            if (Properties.Settings.Default.SendMailAddress != null)
            {
                SendEmail.Text = Properties.Settings.Default.SendMailAddress;
            }
            if (Properties.Settings.Default.SendUserName != null)
            {
                SendUserName.Text = Properties.Settings.Default.SendUserName;
            }
            if (Properties.Settings.Default.SMTPServer != null)
            {
                SmptServer.Text = Properties.Settings.Default.SMTPServer;
            }
            if (Properties.Settings.Default.SMTPServerPort != null)
            {
                SmptServerPort.Text = Properties.Settings.Default.SMTPServerPort;
            }
            if (Properties.Settings.Default.UserName != null)
            {
                SettingUserName.Text = Properties.Settings.Default.UserName;
            }
            if (Properties.Settings.Default.UserPass != null)
            {
                SettingUserPassword.Text = System.Text.Encoding.ASCII.GetString(System.Convert.FromBase64String(Properties.Settings.Default.UserPass));
            }
            SSLConnection.Checked = Properties.Settings.Default.IsSSL;
            if (Properties.Settings.Default.Send1CC != null)
            {
                SendMessage1CC.Text = Properties.Settings.Default.Send1CC;
            }
            if (Properties.Settings.Default.Send1To != null)
            {
                SendMessage1To.Text = Properties.Settings.Default.Send1To;
            }
            if (Properties.Settings.Default.Send1Title != null)
            {
                SendMessage1Subject.Text = Properties.Settings.Default.Send1Title;
            }
            if (Properties.Settings.Default.Send1Text != null)
            {
                SendMessage1Text.Text = Properties.Settings.Default.Send1Text;
            }
            if (Properties.Settings.Default.Send2To != null)
            {
                SendMessage2To.Text = Properties.Settings.Default.Send2To;
            }
            if (Properties.Settings.Default.Send2CC != null)
            {
                SendMessage2CC.Text = Properties.Settings.Default.Send2CC;
            }
            if (Properties.Settings.Default.Send2Title != null)
            {
                SendMessage2Subject.Text = Properties.Settings.Default.Send2Title;
            }
            if (Properties.Settings.Default.Send2Text != null)
            {
                SendMessage2Text.Text = Properties.Settings.Default.Send2Text;
            }
            if (Properties.Settings.Default.DataList != null)
            {
                foreach (var item in Properties.Settings.Default.DataList)
                {
                    SendDataListBox.Items.Add(item);
                }
            }
            if (Properties.Settings.Default.SendMessageTime1 != null)
            {
                Send1Time.Value = Properties.Settings.Default.SendMessageTime1;
            }
            if (Properties.Settings.Default.SendMessageTime2 != null)
            {
                Send2Time.Value = Properties.Settings.Default.SendMessageTime2;
            }
            if(Properties.Settings.Default.Macro == true)
            {
                checkBox2.Checked = true;
            }
        }

        private void StartSendMail_Click(object sender, EventArgs e)
        {
            if (!SendTest)
            {
                MessageBox.Show("テストが実行されていないもしくは成功していません。\nサーバー設定タブよりテストをしてください。");
                return;
            }
            StopSendMail.Enabled = true;
            StartSendMail.Enabled = false;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ja-JP");
            SendMailTimer.Start();
        }

        private void SendMailTimer_Tick(object sender, EventArgs e)
        {
            var dateTimes = new List<DateTime>();
            foreach (var item in SendDataListBox.Items)
            {
                dateTimes.Add(DateTime.ParseExact(item.ToString(), DateFormat, null));
                if (DateTime.ParseExact(item.ToString(), DateFormat, null).Date == DateTime.Now.Date)
                {
                    if (Send1Time.Value.ToShortTimeString() == DateTime.Now.ToShortTimeString())
                    {
                        if(SendMessage1To.Text != "" || SendMessage1CC.Text != "")
                        {
                            SendMail(SendMessage1To.Text, SendMessage1CC.Text, SendMessage1Subject.Text, SendMessage1Text.Text);
                        }
                    }
                    else if (Send2Time.Value.ToShortTimeString() == DateTime.Now.ToShortTimeString())
                    {
                        if(SendMessage2To.Text != "" || SendMessage2CC.Text != "")
                        {
                            SendMail(SendMessage2To.Text, SendMessage2CC.Text, SendMessage2Subject.Text, SendMessage2Text.Text);
                        }
                    }
                }
            }

        }

        private void StopSendMail_Click(object sender, EventArgs e)
        {
            SendMailTimer.Stop();
            StopSendMail.Enabled = false;
            StartSendMail.Enabled = true;
        }

        private void SendMail(string SendMailToNameAdress, string SendMailCCNameAdress, string SendMailSubject, string SendMailText)
        {
            var host = SmptServer.Text;
            var portstr = SmptServerPort.Text;
            int port = Convert.ToInt32(portstr);
            using (var smtp = new MailKit.Net.Smtp.SmtpClient())
            {
                try
                {
                    //開発用のSMTPサーバが暗号化に対応していないときは、次の行をコメントアウト
                    if (SSLConnection.Checked)
                    {
                        smtp.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    }
                    smtp.Connect(host, port, MailKit.Security.SecureSocketOptions.Auto);
                    //認証設定
                    smtp.Authenticate(SettingUserName.Text, SettingUserPassword.Text);

                    //送信するメールを作成する
                    var mail = new MimeKit.MimeMessage();
                    var builder = new MimeKit.BodyBuilder();
                    mail.From.Add(new MimeKit.MailboxAddress(SendUserName.Text, SendEmail.Text));


                    string[] SendToNameAdress = SendMailToNameAdress.Split(';');
                    string[] SendCCNameAdress = SendMailCCNameAdress.Split(';');
                    Regex mailreg = new Regex("<(?<MailAdress>.*)>");
                    Regex namereg = new Regex("(?<Name>.*)<");
                    foreach (var item in SendToNameAdress)
                    {
                        var SendToName = (namereg.Match(item.ToString()).Groups["Name"].Value);
                        var SendToAdress = (mailreg.Match(item.ToString()).Groups["MailAdress"].Value);
                        if(SendToName != "" && SendToAdress != "")
                        {
                            mail.To.Add(new MimeKit.MailboxAddress(SendToName, SendToAdress));
                        }
                    }
                    foreach (var item in SendCCNameAdress)
                    {
                        var SendCCName = (namereg.Match(item.ToString()).Groups["Name"].Value);
                        var SendCCAdress = (mailreg.Match(item.ToString()).Groups["MailAdress"].Value);
                        if (SendCCAdress != "" && SendCCName != "")
                        {
                            mail.Cc.Add(new MimeKit.MailboxAddress(SendCCName, SendCCAdress));
                        }
                    }
                    mail.Bcc.Add(new MimeKit.MailboxAddress(SendUserName.Text, SendEmail.Text));
                    mail.Subject = RegPattern(SendMailSubject);
                    MimeKit.TextPart textPart = new MimeKit.TextPart("Plain");
                    textPart.Text = RegPattern(SendMailText);
                    var multipart = new MimeKit.Multipart("mixed");
                    multipart.Add(textPart);
                    mail.Body = multipart;
                    smtp.Send(mail);
                    SendTest = true;
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                    SendTest = false;
                }
                finally
                {
                    //SMTPサーバから切断する
                    smtp.Disconnect(true);
                    SendMailTest.Enabled = true;
                }
            }
        }
        public string RegPattern(string Str)
        {
            string RegedString = Regex.Replace(Str, "{day}", DateTime.Now.Day.ToString()).Replace("{month}", DateTime.Now.Month.ToString());
            while (Regex.IsMatch(RegedString, "{ran:(.+?)}"))
            {
                Match m = Regex.Match(RegedString, "{ran:(?<num>[0-9]*-[0-9]*)}");
                string numstr = m.Groups["num"].Value;
                string[] numstrs = numstr.Split('-');
                int a =Int32.Parse(numstrs[0]);
                int b =Int32.Parse(numstrs[1]);
                int r = 0;
                Random ran = new Random(); 
                if (a < b)
                {
                    r = ran.Next(a, b);
                }
                else
                {
                    r = ran.Next(b, a);
                }
                RegedString = Regex.Replace(RegedString,m.Value, r.ToString());
            }
            return RegedString;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Idle -= new EventHandler(Application_Idle);
            Save();
        }

        private void SendMessage2Test_Click(object sender, EventArgs e)
        {
            var title = RegPattern(SendMessage2Subject.Text);
            var text = RegPattern(SendMessage2Text.Text);
            MessageBox.Show("件名:" + title + "\n"+"内容:\n"+text);
        }

        private void SendMessage1Test_Click(object sender, EventArgs e)
        {
            var title = RegPattern(SendMessage1Subject.Text);
            var text = RegPattern(SendMessage1Text.Text);
            MessageBox.Show("件名:" + title + "\n" + "内容:\n" + text);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }
        Random Random = new Random();
        private void macro_Tick(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                if(MousePoint == System.Windows.Forms.Cursor.Position)
                {
                    RestoreMinimizedWindow();
                    tabControl1.SelectedTab = TabHelp;
                    MousePoint = new Point(Random.Next(21, 320), Random.Next(72, 420));
                    Point mp = this.PointToScreen(MousePoint);
                    System.Windows.Forms.Cursor.Position = mp;
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                }
                else
                {
                    MousePoint = System.Windows.Forms.Cursor.Position;
                    macro.Stop();
                }
            }
        }
        private void Application_Idle(object sender,EventArgs e)
        {
            DateTime dateTime = DateTime.Now;
            if (checkBox2.Checked && 
                MousePoint == System.Windows.Forms.Cursor.Position && 
                dateTime.TimeOfDay>Send1Time.Value.TimeOfDay &&  
                dateTime.TimeOfDay < Send2Time.Value.TimeOfDay)
            {
                if (!macro.Enabled)
                {
                    macro.Start();
                }
            }
            MousePoint = System.Windows.Forms.Cursor.Position;
        }

        private FormWindowState preWindowState;

        public void RestoreMinimizedWindow()
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.WindowState = this.preWindowState;
            }
        }

        private void AutoSendMail_SizeChanged(object sender, EventArgs e)
        {
            //最小化された以外の時に、状態を覚えておく
            if (this.WindowState != FormWindowState.Minimized)
            {
                this.preWindowState = this.WindowState;
            }
        }
    }
}
