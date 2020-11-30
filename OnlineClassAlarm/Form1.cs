using System;
using System.Drawing;
using System.IO;
using System.Media;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using MaterialSkin.Controls;
using MaterialSkin;
using System.Collections.Generic;
using System.Text;

namespace OnlineClassAlarm
{
    public partial class Form1 : MaterialForm
    {
        [DllImport("user32.dll")]
        public static extern int FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern int FindWindowEx(int hWnd1, int hWnd2, string lpsz1, string lpsz2);

        [DllImport("user32.dll")]
        public static extern int SendMessage(int hwnd, int wMsg, int wParam, string lParam);

        [DllImport("user32.dll")]
        public static extern uint PostMessage(int hwnd, int wMsg, int wParam, int lParam);

        public Form1()
        {
            InitializeComponent();
        }
        public string[] tableNum;
        public string[] setting;
        public string[,] finalTable = new string[7, 10];
        public string[] timeTable = new string[10];
        public string tablePath = System.Windows.Forms.Application.StartupPath + @"\table.csv";
        public string iniPath = System.Windows.Forms.Application.StartupPath + @"\packages\setting.ini";
        public bool[] check = new bool[7];
        public List<string> roomList = new List<string>();
        public void readTable()
        {

            if (File.Exists(tablePath))
            {
                tableNum = System.IO.File.ReadAllLines(tablePath, Encoding.Default);
                
            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("table.csv를 작성해주세요.", "OnlineClassAlarm", MessageBoxButtons.OK);
                Environment.Exit(0);
            }
            string[] temp = tableNum[0].Split(',');
            for (int j = 0; j < temp.Length - 1; j++)
            {
                for (int i = 0; i < tableNum.Length - 1; i++)
                {
                    string[] temp2 = tableNum[i + 1].Split(',');
                    finalTable[j, i] = temp2[j];
                }
            }
            for (int i = 0; i < tableNum.Length - 1; i++)
            {
                string[] temp2 = tableNum[i + 1].Split(',');
                timeTable[i] = temp2[temp.Length - 1];
            }
        }
        public void readSetting()
        {
            if (File.Exists(iniPath))
            {
                setting = System.IO.File.ReadAllLines(iniPath, Encoding.Default);
                for (int i = 0; i < setting.Length; i++)
                {
                    var row = new string[] { setting[i] };
                    var item = new ListViewItem(row);
                    listView1.Items.Add(item);
                    roomList.Add(setting[i]);
                }
            }
            else
            {
                System.IO.File.WriteAllText(iniPath, "", Encoding.Default);
            }
        }
        public void sendMessage(string windowName, string message)
        {
            int hd01 = FindWindow(null, windowName);
            int hd03 = FindWindowEx(hd01, 0, "RichEdit50W", "");
            if (!hd03.Equals(IntPtr.Zero))
            {
                SendMessage(hd03, 0x000c, 0, message);
                PostMessage(hd03, 0x0100, 0xD, 0x1C001);
            }
        }
        public void getTime(int day)
        {
            DateTime time1 = DateTime.Now;
            for (int i = 0; i < tableNum.Length - 1; i++)
            {
                string[] timeTemp = timeTable[i].Split(':');
                TimeSpan gap = time1 - new DateTime(time1.Year, time1.Month, time1.Day, Convert.ToInt32(timeTemp[0]), Convert.ToInt32(timeTemp[1]), 0);
                if ((int)gap.TotalSeconds <= 10 && (int)gap.TotalSeconds >= 0)
                {
                    if (!(string.IsNullOrEmpty(roomList[0])))
                    {
                        int temp = i + 1;
                        if (!(string.IsNullOrEmpty(finalTable[day, i])))
                        {
                            if (check[i] != true)
                            {
                                readTable();
                                this.Invoke(new Action(
                                delegate ()
                                {
                                    label1.Text = temp + "교시 " + finalTable[day, i] + "시간입니다.";
                                }));
                                foreach (string s in roomList)
                                {
                                    sendMessage(s, temp + "교시 " + finalTable[day, i] + "시간입니다.");
                                }
                                SystemSounds.Beep.Play();
                                MessageBox.Show(temp + "교시 " + finalTable[day, i] + "시간입니다.", "OnlineClassAlarm", MessageBoxButtons.OK);
                                check[i] = true;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    break;
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            readTable();
            readSetting();
            label1.Text = "알림이 시작되지 않았습니다.";
        }
        private void callMethod()
        {
            while (true)
            {
                DateTime dt = DateTime.Now;
                var day = dt.DayOfWeek;
                switch (day)
                {
                    case DayOfWeek.Monday:
                        getTime(0);
                        break;
                    case DayOfWeek.Tuesday:
                        getTime(1);
                        break;
                    case DayOfWeek.Wednesday:
                        getTime(2);
                        break;
                    case DayOfWeek.Thursday:
                        getTime(3);
                        break;
                    case DayOfWeek.Friday:
                        getTime(4);
                        break;
                    default:
                        break;
                }
                Thread.Sleep(3000);
            }
        }
        public Thread t1;
        private void button2_Click(object sender, EventArgs e)
        {
            label1.Text = "알림이 시작되었습니다.";
            t1 = new Thread(callMethod);
            t1.Start();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!(string.IsNullOrEmpty(textBox1.Text)))
            {
                var row = new string[] { textBox1.Text };
                var item = new ListViewItem(row);
                listView1.Items.Add(item);
                roomList.Add(textBox1.Text);
                if (File.Exists(iniPath))
                {
                    System.IO.File.Delete(iniPath);
                    setting = roomList.ToArray();
                    System.IO.File.WriteAllLines(iniPath, setting, Encoding.Default);
                }
                else if (!(File.Exists(iniPath)))
                {
                    setting = roomList.ToArray();
                    System.IO.File.WriteAllLines(iniPath, setting, Encoding.Default);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 1)
            {
                int index = listView1.SelectedItems[0].Index;
                listView1.Items.RemoveAt(index);
                roomList.RemoveAt(index);
                if (File.Exists(iniPath))
                {
                    System.IO.File.Delete(iniPath);
                    setting = roomList.ToArray();
                    System.IO.File.WriteAllLines(iniPath, setting, Encoding.Default);
                }
                else if (!(File.Exists(iniPath)))
                {
                    setting = roomList.ToArray();
                    System.IO.File.WriteAllLines(iniPath, setting, Encoding.Default);
                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (t1 != null)
            {
                t1.Abort();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (t1 != null)
            {
                label1.Text = "알림이 시작되지 않았습니다.";
                t1.Abort();
            }
        }
    }
}
