using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace MonitorPorts
{
    public partial class FilterForm : Form
    {

        #region 变量
        public string Protocol;
        public string SourceIP_Port;
        public string SourcePort;
        public string DestIP_Port;
        public string DestPort;
        public string maxIntervalVtoZ;
        public string maxIntervalZtoV; 
        
        public List<string> listSourIP;
        public List<string> listSourPort;
        public List<string> listDestIP;
        public List<string> listDestPort;

        public Dictionary<string, string> dic;

        public Dictionary<string, string> dicIP = new Dictionary<string, string>();

        public int PcapLength;
        public string PcapLengthUnit;
        public static double PcapLengthNum;


        #endregion


        #region 构造函数
        public FilterForm()
        {
            Protocol = "TCP,UDP,SCTP,ICMP,IGMP,";
            SourceIP_Port = "";
            SourcePort = "";
            DestIP_Port = "";
            DestPort = "";
            InitializeComponent();

            listSourIP = new List<string>();
            listSourPort = new List<string>();
            listDestIP = new List<string>();
            listDestPort = new List<string>();
            //SourceListBox1 = new CheckedListBox();
            //DestListBox2 = new CheckedListBox();
            dic = ReadIPlist();
            ShowInForm(dic);
            LoadPcapLength();
        }
        #endregion


        #region 窗体控件

        private void button_OK_Click(object sender, EventArgs e)
        {
            SetPcapLength();
            SaveSettings();
            this.Hide();

            if (listSourIP.Count != 0)
            {
                for (int iS = 0; iS <= listSourIP.Count() - 1; iS++)
                {
                    SourceIP_Port += listSourIP[iS] + "-" + listSourPort[iS] + ";";
                }
            }

            //保存并显示目的IP和端口
            if (listDestIP.Count != 0)
            {
                for (int iD = 0; iD <= listDestIP.Count() - 1; iD++)
                {
                    DestIP_Port += listDestIP[iD] + "-" + listDestPort[iD] + ";";

                }
            }

            maxIntervalVtoZ = textBox1.Text;
            maxIntervalZtoV = textBox2.Text;

            FileStream fs = new FileStream("a.txt", FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入
            sw.WriteLine(Protocol);
            sw.WriteLine(SourceIP_Port);
            sw.WriteLine(DestIP_Port);
            sw.WriteLine(maxIntervalVtoZ);
            sw.WriteLine(maxIntervalZtoV);
            //清空缓冲区
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();

        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }


        private void FilterForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }

        #endregion


        #region 内部方法
        public void SetPcapLength()
        {
            PcapLength = Convert.ToInt16(this.txtPcapLength.Text);
            if (this.Unit1.Checked)
            {
                PcapLengthUnit = "MB";
            }
            else
            {
                PcapLengthUnit = "GB";
            }
            FileStream fs = new FileStream("PcapLength.txt", FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(PcapLength.ToString());
            sw.WriteLine(PcapLengthUnit);
            sw.Flush();
            sw.Close();
            fs.Close();
            GetPcapLengthNum();
        }

        public void GetPcapLengthNum()
        {
            if (PcapLengthUnit == "MB")
            {
                PcapLengthNum = PcapLength * 1024 * 1024;
            }
            if (PcapLengthUnit == "GB")
            {
                PcapLengthNum = PcapLength * 1024 * 1024 * 1024;
            }
        }

        /// <summary>
        /// 读取Pcap文件大小的历史记录
        /// </summary>
        private void LoadPcapLength()
        {
            string PcapLengthPath = "PcapLength.txt";
            StreamReader sr = new StreamReader(PcapLengthPath, Encoding.Default);
            string Line = null;
            if ((Line = sr.ReadLine()) != null)
            {
                this.txtPcapLength.Text = Line;
                if ((Line = sr.ReadLine()) != null)
                {
                    if (Line == "GB")
                    {
                        this.Unit2.Checked = true;
                        this.Unit1.Checked = false;
                    }
                    if (Line == "MB")
                    {
                        this.Unit1.Checked = true;
                        this.Unit2.Checked = false;
                    }
                }
            }
            sr.Close();
        }


        /// <summary>
        /// 保存本页面设置信息，存为txt文件
        /// </summary>
        private void SaveSettings()
        {
            //协议
            string ProtocolString = "";
            Protocol = ProtocolString;

            for (int i = 0; i < SourceListBox1.Items.Count; i++)
            {
                if (SourceListBox1.GetItemChecked(i))
                {
                    string sourIPandPort = dic[SourceListBox1.Items[i].ToString()];
                    listSourIP.Add(sourIPandPort.Split(new char[] { '-' })[0]);
                    listSourPort.Add(sourIPandPort.Split(new char[] { '-' })[1]);
                }
            }
            for (int i = 0; i < DestListBox2.Items.Count; i++)
            {
                if (DestListBox2.GetItemChecked(i))
                {
                    string destIPandPort = dic[SourceListBox1.Items[i].ToString()];
                    listDestIP.Add(destIPandPort.Split(new char[] { '-' })[0]);
                    listDestPort.Add(destIPandPort.Split(new char[] { '-' })[1]);
                }
            }


        }

        

        /// <summary>
        /// 将IPlist信息显示在本界面的两个ListBox中
        /// </summary>
        /// Dictionary<IP对应名称,IP-端口>
        /// <param name="dic"></param>
        public  void ShowInForm(Dictionary<string,string> dic )
        {


            foreach(string str in dic.Keys)
            {
                SourceListBox1.Items.Add(str);
                DestListBox2.Items.Add(str);
            }
        }

        /// <summary>
        /// 加载IPlist等基本配置信息，读取为字典表格式
        /// Dictionary<IP对应名称,IP-端口>
        /// </summary>
        private Dictionary<string, string> ReadIPlist()
        {
            Dictionary<string, string> tmpdic = new Dictionary<string, string>();
            List<string> list = new List<string>();
            string path = "IPlist.csv";
            StreamReader sr = new StreamReader(path, Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                list.Add(line);
            }
            for (int i = 1; i < list.Count; i++)
            {
                string[] strs = list[i].Split(new char[] { ',' });
                string IP_Port = strs[1] + "-" + strs[2];
                tmpdic.Add(strs[0], IP_Port);
            }
            return tmpdic;
        }



        #endregion

        private void Unit2_CheckedChanged(object sender, EventArgs e)
        {
            if (Unit2.Checked)
            {
                Unit1.Checked = false;
                Unit2.Checked = true;
            }
        }

        private void Unit1_CheckedChanged(object sender, EventArgs e)
        {
            if (Unit1.Checked)
            {
                Unit2.Checked = false;
                Unit1.Checked = true;
            }
        }




    }
}
