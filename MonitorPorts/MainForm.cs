using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Management;
using System.Net.Sockets;
using System.Threading;
using System.Xml;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace MonitorPorts
{
    public partial class MainForm : Form
    {
        #region 全局变量
        List<IPList> IPlist = new List<IPList>();
        private int ItemID;
        private string captureTime;
        private readonly string _path = Application.StartupPath;// 程序所在路径。
        private bool m_runningFlg; //程序运行标志位 TRUE：运行 FALSE：停止
        ushort DeleteDataInterval = 7;    //单位:d,设置多长时间删除文件

        private int FilterCount;
        private SnifferService Sniffer;
        private FilterForm FilterOptionForm;
        private SettingSaves SettingSaves;
        Excel.Application myExcel;
        string FileName = null;

        bool bStatus = false;//标记按钮当前状态,false表示按钮目前显示的是扫描，true显示的是暂停

        delegate void listViewDelegate(string ID, DateTime Capturetime, string Protocol, string Interval, string SourceIP, string SourcePort, string DestIP, string DestPort, 
                                        string AllLength, string MessageBodyTxt, string MessageBodyLen, string MessageBodyHex ,byte[] ReceiveBuf,int BufLength);
        delegate void refresh(DateTime Time, string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort, uint IPHeaderLength, Byte[] IPHeaderBuffer, uint MessageLength,
                                        Byte[] MessageBuffer, uint PacketLength, Byte[] PacketBuffer,byte[] ReceiveBuf,int Buflength); 
        const int ICMPDataOffset = 4;//ICMP包头长度
        const int IGMPDataOffset = 4;//IGMP包头长度
        const int TCPDataOffset = 20;//TCP包头长度
        const int UDPDataOffset = 8;//UDP包头长度
        const int SCTPDataOffset = 12;//SCTP包头长度(通用数据头)
        //edit by shuya
        byte[] bytes = new byte[30];
        List<CBTCByte> listNode;
        List<double> list = new List<double>();
        Dictionary<string, DateTime> dicDoubleIP_CaptureTime = new Dictionary<string,DateTime>();
        Dictionary<string, int> listVtoZDoubleIP = new Dictionary<string, int>();
        Dictionary<string, int> listZtoVDoubleIP = new Dictionary<string, int>();
        List<Color> listVtoZColor = new List<Color>();
        List<Color> listZtoVColor = new List<Color>();

        DateTime countInterval;
        List<string> xData = new List<string> { "0-500", "500-1000", "1000-2400", "≥2400" };
        List<int> yData = new List<int> { 0, 0, 0, 0 };


        //实例化生成pcapfile类
        GeneratePcapFile genePcapFile = null;
        //文件路径
        private string _fileSavePath = "";

        public string FileSavePath
        {
            get { return _fileSavePath; }
            set { _fileSavePath = value; }
        }
        #endregion

        /// <summary>
        /// 构造函数
        /// </summary>
        public MainForm()
        {
            ItemID = 0;
            FilterCount = 0;
            FilterOptionForm = new FilterForm();
            SettingSaves = new MonitorPorts.SettingSaves();
            //edit by shuya
            listNode = ReadCBTCXml();
            
            Sniffer = new SnifferService();
            Sniffer.PacketArrival += new SnifferService.PacketArrivedEventHandler(SnifferSocket_PacketArrival);
            myExcel = new Excel.Application();

            InitializeComponent();

            GetIPList();

            genePcapFile = new GeneratePcapFile();
            FileName = System.DateTime.Now.ToString("yyyy") + "." + System.DateTime.Now.ToString("MM")+ "." + System.DateTime.Now.ToString("dd") + "  " + System.DateTime.Now.ToString("HH：mm");
            // 建立pcap文件
            _fileSavePath = Application.StartupPath + "\\pcap";// 程序所在路径。
            genePcapFile.CreatPcap(_fileSavePath, FileName);
        }


        #region 窗体控件
        //开始
        private void Button_Start_Click(object sender, EventArgs e)
        {
            SettingSaves = new MonitorPorts.SettingSaves();
            countInterval = DateTime.Now;

            if (bStatus == false)
            {
                Clear();
                myExcel.Visible = true;
                
                myExcel.Workbooks.Add(true); 
                
                Button_Start.Enabled = false;

                listZtoVColor.Clear();
                listVtoZColor.Clear();
                chartVtoZ1.Series.Clear();
                chartZtoV1.Series.Clear();
                Sniffer.Start();
                Button_Stop.Enabled = true;
                Buton_Find.Enabled = false;
                Button_ClearList.Enabled = false;
                Button_Filter.Enabled = false;

                ServiceStatus.Text = "开始监听";
                UpdateStatus();

            }
            else
            {
                Button_Start.Enabled = false;
                Sniffer.Start();
                Button_Stop.Enabled = true;
                Buton_Find.Enabled = false;
                Button_ClearList.Enabled = false;
                Button_Filter.Enabled = false;
                ServiceStatus.Text = "开始监听";
                UpdateStatus();
            }
        }
        //停止
        private void Button_Stop_Click(object sender, EventArgs e)
        {
            Button_Stop.Enabled = false;
            Sniffer.Stop();
            Button_Start.Enabled = true;
            Buton_Find.Enabled = true;
            Button_ClearList.Enabled = true;
            Button_Filter.Enabled = true;
            bStatus = false;//状态切换
            ServiceStatus.Text = "停止监听";
            UpdateStatus();
            //myExcel.Visible = true ;
        }
        //清除
        private void Button_ClearList_Click(object sender, EventArgs e)
        {
            Clear();
        }
        //暂停
        private void Button_Pause_Click(object sender, EventArgs e)
        {
            Button_Stop.Enabled = false;
            Sniffer.Stop();
            Button_Start.Enabled = true;
            bStatus = true;//状态切换
            ServiceStatus.Text = "暂停监听";
            UpdateStatus();
        }
        //筛选设置
        private void Button_Filter_Click(object sender, EventArgs e)
        {
            if (FilterOptionForm.ShowDialog() == DialogResult.OK)
            {
                FilterOptionForm.Show();
            }
        }
        //离线查找
        private void Button_Find_Click(object sender, EventArgs e)
        {
            QueryForm queryFrm = new QueryForm();
            this.Hide();
            queryFrm.ShowDialog();
        }

        //关闭界面
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("你确定要退出吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
            }
            else
            {
                System.Environment.Exit(0);
            }
        }
        #endregion


        #region 内部方法


        private void BuiltPcapFile(byte[] packet, int packetlenth)
        {
            genePcapFile.WritePacketData(System.DateTime.Now, packet, packetlenth);
        }
        public System.Drawing.Color GetRandomColor()
        {
            Random RandomNum_First = new Random((int)DateTime.Now.Ticks);
            System.Threading.Thread.Sleep(RandomNum_First.Next(50));
            Random RandomNum_Sencond = new Random((int)DateTime.Now.Ticks);

            int int_Red = RandomNum_First.Next(256);
            int int_Green = RandomNum_Sencond.Next(256);
            int int_Blue = (int_Red + int_Green > 400) ? 0 : 400 - int_Red - int_Green;
            int_Blue = (int_Blue > 255) ? 255 : int_Blue;

            return System.Drawing.Color.FromArgb(int_Red, int_Green, int_Blue);
        }
        public void GetIPList()
        {
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
                IPList List = new IPList();
                string[] strs = list[i].Split(new char[] { ',' });
                List.Name = strs[0];
                List.IP = strs[1];
                List.Port = strs[2];
                List.Type = strs[3];
                List.IsChoose = false;
                IPlist.Add(List);
            }
        }
     
        public void Data_Receive(DateTime Time, string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort, uint IPHeaderLength, Byte[] IPHeaderBuffer, uint MessageLength,
                                Byte[] MessageBuffer, uint PacketLength, Byte[] PacketBuffer,byte[] ReceiveBuf,int Buflength)
        {

            //ItemID++;
            captureTime = Time.ToString("HH:mm:ss:fff");

            string doubleIP_Port = SourceIP + "-" + SourcePort + ";" + DestIP + "-" + DestPort;
            // 数据显示
            listViewDelegate listDelegate = new listViewDelegate(AddItem);
            int IPHeaderLen = (int)IPHeaderLength;
            string IPHeaderHex = GetDataHex(IPHeaderBuffer, 0, IPHeaderLen);

            int MessageHeaderLen = GetMessageHeaderLen(Protocol);
            string MessageHeaderHex = GetDataHex(MessageBuffer, 0, MessageHeaderLen);

            int MessageBodyLen = (int)MessageLength - MessageHeaderLen;
            string MessageBodyTxt = GetDataTxt(MessageBuffer, MessageHeaderLen, MessageBodyLen);
            string MessageBodyHex = GetDataHex(MessageBuffer, MessageHeaderLen, MessageBodyLen);
            int MessageALL = IPHeaderLen + MessageHeaderLen + MessageBodyLen;
            string IPHeader = Convert.ToString(IPHeaderLen);
            string MessageHLength = Convert.ToString(MessageHeaderLen);
            string AllLength = Convert.ToString(MessageALL);
            FilterCount++;

            ////////////////////////////////
            /////////林庆庆编写/////////////
            if (DataFilterByIPAndPort(Protocol, SourceIP, SourcePort, DestIP, DestPort))
            {
                ItemID++;
                if (dicDoubleIP_CaptureTime.Keys.Contains(doubleIP_Port))
                {
                    double interval = (Time - dicDoubleIP_CaptureTime[doubleIP_Port]).TotalSeconds;
                    Array.Copy(MessageBuffer, 8, bytes, 0, 30);
                    dicDoubleIP_CaptureTime[doubleIP_Port] = Time;
                    listView_Data.Invoke(listDelegate, ItemID.ToString(), Time, interval.ToString(), Protocol, SourceIP, SourcePort, DestIP, DestPort, AllLength, MessageBodyLen.ToString(), MessageBodyTxt, MessageBodyHex,ReceiveBuf,Buflength);
                }
                else
                {
                    string type = null;
                    foreach (var item in IPlist)
                    {
                        if (SourceIP.Contains(item.IP) && item.Port == SourcePort)
                        {
                            type = item.Type;
                        }
                    }
                    if (type == "车")
                    {
                        listVtoZDoubleIP.Add(SourceIP + "-" + SourcePort + "--" + DestIP + "-" + DestPort, listVtoZDoubleIP.Count);
                        listVtoZColor.Add(GetRandomColor());
                        Series s1 = new Series();
                        s1.Name = SourceIP + "-" + SourcePort + "--" + DestIP + "-" + DestPort;
                        this.chartVtoZ1.Series.Add(s1);
                    }
                    else
                    {
                        listZtoVDoubleIP.Add(SourceIP + "-" + SourcePort + "--" + DestIP + "-" + DestPort, listZtoVDoubleIP.Count);
                        listZtoVColor.Add(GetRandomColor());
                        Series s1 = new Series();
                        s1.Name = SourceIP + "-" + SourcePort + "--" + DestIP + "-" + DestPort;
                        this.chartZtoV1.Series.Add(s1);
                    }
                    dicDoubleIP_CaptureTime.Add(doubleIP_Port,Time);
                    double interval = (Time - dicDoubleIP_CaptureTime[doubleIP_Port]).TotalSeconds;
                    Array.Copy(MessageBuffer, 8, bytes, 0, 30);
                    listView_Data.Invoke(listDelegate, ItemID.ToString(), Time, interval.ToString(), Protocol, SourceIP, SourcePort, DestIP, DestPort, AllLength, MessageBodyLen.ToString(), MessageBodyTxt, MessageBodyHex,ReceiveBuf,Buflength);
                }
            }
            /////////林庆庆编写/////////////
            ////////////////////////////////
        }


        ////////////////////////////////
        /////////林庆庆编写/////////////
        public bool DataFilterByIPAndPort(string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort)
        {
            List<string> SrcIPPart = new List<string>();
            List<string> DstIPPart = new List<string>();
            foreach (var Src in SettingSaves.listSourIPandPort)
            {
                string[] src = Src.Split(new char[] { '-' });
                SrcIPPart.Add(src[0]);
            }
            foreach (var Dst in SettingSaves.listDestIPandPort)
            {
                string[] dst = Dst.Split(new char[] { '-' });
                DstIPPart.Add(dst[0]);
            }
            bool a = false;
            bool b = false;
            foreach (var Src in SrcIPPart)
            {
                if (SourceIP.Contains(Src))
                {
                    a = true;
                }
            }
            foreach (var Dst in DstIPPart)
            {
                if (DestIP.Contains(Dst))
                {
                    b = true;
                }
            }
            string SrcType = null;
            string DstType = null;
            foreach (var item in IPlist)
            {
                if (SourceIP.Contains(item.IP) && SourcePort == item.Port)
                {
                    SrcType = item.Type;
                    break;
                }
            }
            foreach (var item in IPlist)
            {
                if (DestIP.Contains(item.IP) && DestPort == item.Port)
                {
                    DstType = item.Type;
                    break;
                }
            }
            if (a == false || b == false)
            {
                return false;
            }
            if (SrcType == DstType)
            {
                return false;
            }
            return true;
        }
        /////////林庆庆编写/////////////
        ////////////////////////////////


        /// <summary>
        /// 将捕获的数据添加到ListView中
        /// </summary>
        private void AddItem(string ID, DateTime Capturetime, string Interval, string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort, string AllLength, string MessageBodyLen, string MessageBodyTxt, string MessageBodyHex,byte[] ReceiveBuf,int BufLength)
        {

            if (DataFilter(Protocol, SourceIP, SourcePort, DestIP, DestPort))
            {

                System.IO.FileInfo File = new FileInfo(_fileSavePath + "\\" + FileName + ".pcap");
                if (File.Length < FilterForm.PcapLengthNum)
                {
                    BuiltPcapFile(ReceiveBuf, BufLength);
                }
                else
                {
                    genePcapFile = new GeneratePcapFile();
                    FileName = System.DateTime.Now.ToString("yyyy") + "." + System.DateTime.Now.ToString("MM") + "." + System.DateTime.Now.ToString("dd") + "  " + System.DateTime.Now.ToString("HH：mm"); ;
                    // 建立pcap文件
                    _fileSavePath = Application.StartupPath + "\\pcap";// 程序所在路径。
                    genePcapFile.CreatPcap(_fileSavePath, FileName);
                }

                string SrcType = null;
                foreach (var item in IPlist)
	            {
                    if (SourceIP.Contains(item.IP) && SourcePort == item.Port)
	                {
                        SrcType = item.Type;
	                }
	            }
                if (SrcType == "车")
                {
                    //异常报警
                    if (Convert.ToDouble(Interval) > SettingSaves.MaxIntervalVtoZ)
                    {
                        listView_Data.Items.Add(new ListViewItem(new string[] { ID, Capturetime.ToString("yyyy-MM-dd"), Capturetime.ToString("HH:mm:ss:fff"), Interval, SourceIP, DestIP, MessageBodyHex }));
                        SaveToAlarmLog(ID, Capturetime.ToString("yyyy-MM-dd"), Capturetime.ToString("HH:mm:ss:fff"), Interval, SourceIP, DestIP, MessageBodyHex);
                        List<string> alarmFiles = GetFiles(Capturetime.AddSeconds(-5), Capturetime.AddSeconds(+5), SourceIP, DestIP);

                        try
                        {
                           myExcel.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                           for (int i = 0; i <= alarmFiles.Count(); i++)
                            {
                                string[] alarmFileSplited = alarmFiles[i].Split(',');
                                for (int j = 0; j < alarmFileSplited.Count(); j++)
                                {
                                    myExcel.Cells[i + 1, j + 1] = alarmFileSplited[j];
                                }
                            }
                        }
                        catch
                        {

                        }
                    }
                        
                    //画图
                    string doubleIP = SourceIP + "-" + SourcePort + "--" + DestIP + "-" + DestPort;
                    int indexOfSeries = listVtoZDoubleIP[doubleIP];
                    if (this.chartVtoZ1.Series[indexOfSeries].Points.Count() >= 1500)
                    {
                        this.chartVtoZ1.Series[indexOfSeries].Points.Remove(this.chartVtoZ1.Series[indexOfSeries].Points[0]);
                    }
                    this.chartVtoZ1.Series[indexOfSeries].Points.AddY(Convert.ToDouble(Interval));

                    this.chartVtoZ1.Series[indexOfSeries].ChartType = SeriesChartType.Line;
                    chartVtoZ1.Series[indexOfSeries].BorderWidth = 3;
                    chartVtoZ1.Series[indexOfSeries].Color = listVtoZColor[indexOfSeries];

                    double interval =Convert.ToDouble( Interval)*1000;
                    if (interval >= 0 && interval < 500)              yData[0]++;
                    else if (interval >= 500 && interval < 1000)      yData[1]++;
                    else if (interval >= 1000 && interval < 2400)     yData[2]++;
                    else yData[3]++;
                    if((DateTime.Now-countInterval).TotalMilliseconds>=10000)
                    {
                        chartVtoZ2.Series[0]["PieLabelStyle"] = "Outside";//将文字移到外侧
                        chartVtoZ2.Series[0]["PieLineColor"] = "Black";//绘制黑色的连线。
                        chartVtoZ2.Series[0].Points.DataBindXY(xData, yData);
                    }

                }
                else
                {
                    //异常报警
                    if (Convert.ToDouble(Interval) > SettingSaves.MaxIntervalZtoV)
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { ID,Capturetime.ToString("yyyy-MM-dd"), Capturetime.ToString("HH:mm:ss:fff"), Interval, SourceIP, DestIP, MessageBodyHex }));

                        List<string> alarmFiles = GetFiles(Capturetime.AddSeconds(-5), Capturetime.AddSeconds(+5), SourceIP, DestIP);
                        try
                        {
                            myExcel.Sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            for (int i = 0; i <= alarmFiles.Count(); i++)
                            {
                                string[] alarmFileSplited = alarmFiles[i].Split(',');
                                for (int j = 0; j < alarmFileSplited.Count(); j++)
                                {
                                    myExcel.Cells[i + 1, j + 1] = alarmFileSplited[j];
                                }

                            }
                        }
                        catch
                        {

                        }
                    }

                    //画图
                    string doubleIP = SourceIP + "-" + SourcePort + "--" + DestIP + "-" + DestPort;
                    int indexOfSeries = listZtoVDoubleIP[doubleIP];

                    if (this.chartZtoV1.Series[indexOfSeries].Points.Count() >= 1500)
                    {
                        this.chartZtoV1.Series[indexOfSeries].Points.Remove(this.chartZtoV1.Series[indexOfSeries].Points[0]);
                    }
                    this.chartZtoV1.Series[indexOfSeries].Points.AddY(Convert.ToDouble(Interval));

                    this.chartZtoV1.Series[indexOfSeries].ChartType = SeriesChartType.Line;
                    chartZtoV1.Series[indexOfSeries].BorderWidth = 3;
                    chartZtoV1.Series[indexOfSeries].Color = listZtoVColor[indexOfSeries];
 
                    double interval = Convert.ToDouble(Interval) * 1000;
                    if (interval >= 0 && interval < 500) yData[0]++;
                    else if (interval >= 500 && interval < 1000) yData[1]++;
                    else if (interval >= 1000 && interval < 2400) yData[2]++;
                    else yData[3]++;
                    if ((DateTime.Now - countInterval).TotalMilliseconds >= 10000)
                    {
                        chartZtoV2.Series[0]["PieLabelStyle"] = "Outside";//将文字移到外侧
                        chartZtoV2.Series[0]["PieLineColor"] = "Black";//绘制黑色的连线。
                        chartZtoV2.Series[0].Points.DataBindXY(xData, yData);
                    }

                }


                //lastCaptureTime = Capturetime;

            }
            UpdateStatus();
        }

        /// <summary>
        /// 本地保存
        /// </summary>
        private void SaveToTxt(string ID, DateTime Capturetime, string Interval, string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort, string AllLength, string MessageBodyLen, string MessageBodyTxt, string MessageBodyHex)
        {
            string savepath = "";
            string strCatDate = Capturetime.ToString("MM-dd");
            string strCatTime = Capturetime.ToString("HH-mm");
            savepath = _path + @"\BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + Capturetime.Year+"\\"+strCatDate;
            //path = "BytesFilesLogbk.txt";
            //获得字节数组
            string head = ID + ",";
            head += captureTime + ",";
            head += Interval + ",";
            head += Protocol + "," + SourceIP + "," + SourcePort + "," + DestIP + "," + DestPort + "," + AllLength + "," + MessageBodyLen + "," + MessageBodyHex.Replace("\r\n", string.Empty);
            
            if (!Directory.Exists(savepath))
            {
                Directory.CreateDirectory(savepath);
            }

            savepath += "\\" + strCatTime + ".txt";

            try
            {
                using (StreamWriter file = new StreamWriter(savepath, true))
                {
                    file.WriteLine(head);// 直接追加文件末尾，换行   
                }
            }
            catch (Exception e)
            {
                return;
            }           
        }

        #region 每隔7天删除一次数据
        /// <summary>
        /// 每隔一定时间删除一次数据。
        /// </summary>
        /// <param name="delInterval">删除间隔，单位：天。</param>
        private void DeleteOverdueData(ushort delInterval)
        {
            DateTime dtNow = DateTime.Now;
            TimeSpan timeSpan = new TimeSpan(delInterval, 0, 0, 0);


            List<string> fileList1 = new List<string>();
            List<string> fileList = new List<string>();
            DateTime delTime = dtNow - timeSpan;

            try
            {

                string cfgDirPath = _path + @"\BytesFilesLog";

                DirectoryInfo Dir = new DirectoryInfo(cfgDirPath);

                if (Directory.Exists(cfgDirPath))

                {
                    foreach (DirectoryInfo d in Dir.GetDirectories())//查找子目录  
                    {
                        string filePath = _path + @"\BytesFilesLog" + "\\" + d + "\\" + delTime.ToString("yyyy", CultureInfo.CurrentCulture) + @"\";
                        fileList1.AddRange(Directory.GetDirectories(filePath));
                        for (int i = 0; i < fileList1.Count; i++)
                        {
                            DateTime fileCrtTime = File.GetCreationTime(fileList1[i]);
                            if (fileCrtTime < delTime)
                            {
                                fileList.Add(fileList1[i]);
                            }
                        }
                    }
                }
                DeleteFile(fileList, delInterval);
            }
            catch (Exception e)
            {
                String fileLogName = System.IO.Path.Combine(System.IO.Path.Combine(Application.StartupPath, "ErrorLogData"), "Monitor" + DateTime.Now.ToString("yyyy_MM_dd", CultureInfo.CurrentCulture) + ".log");
                if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(fileLogName)))
                {
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(fileLogName));
                }
                System.IO.StreamWriter log = new System.IO.StreamWriter(new System.IO.FileStream(fileLogName, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite));
                System.Text.StringBuilder buff = new System.Text.StringBuilder();

                buff.Append(String.Format(CultureInfo.CurrentCulture, "-----------------------------start({0})------------------------------------\n", DateTime.Now.ToString("o", System.Globalization.CultureInfo.CurrentCulture)));

                buff.Append(e.Message + "\n");

                buff.Append("---------------------------------------End---------------------------------------\n");

                log.WriteLine(buff.ToString());
                log.Flush();
                log.Close();
            }
        }

        /// <summary>
        /// Deletes the file.
        /// </summary>
        /// <param name="rootDir">The dir path.</param>
        private void DeleteFile(List<string> rootDir, ushort delInterval)
        {
            try
            {
                ArrayList _fileList = new ArrayList();

                foreach (string aa in rootDir)
                {
                    if (Directory.Exists(aa))
                    {
                        _fileList.AddRange(Directory.GetFiles(aa));

                        string[] files = Directory.GetFiles(aa);

                        if (files.Length != 0)
                        {
                            Directory.Delete(aa, true);
                        }
                    }
                }

                DateTime crntTime = DateTime.Now;

                // 午夜1-4点的时候执行此操作。
                //if (crntTime.Hour >= 1 && crntTime.Hour < 4)
                //{
                // 删除7天以前的压缩文件
                DateTime lastSevendaysDateTime = crntTime.AddDays(-delInterval);

                if (_fileList != null && _fileList.Count > 0)
                {
                    foreach (string tmpFileName in _fileList)
                    {
                        if (File.Exists(tmpFileName))
                        {
                            DateTime fileCrtTime = File.GetCreationTime(tmpFileName);
                            if (fileCrtTime < lastSevendaysDateTime)  // 7天以前数据
                            {
                                try
                                {
                                    File.Delete(tmpFileName);
                                }
                                catch (System.Exception exp)
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                String fileLogName = System.IO.Path.Combine(System.IO.Path.Combine(Application.StartupPath, "ErrorLogData"), "Monitor" + DateTime.Now.ToString("yyyy_MM_dd", CultureInfo.CurrentCulture) + ".log");
                if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(fileLogName)))
                {
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(fileLogName));
                }
                System.IO.StreamWriter log = new System.IO.StreamWriter(new System.IO.FileStream(fileLogName, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite));
                System.Text.StringBuilder buff = new System.Text.StringBuilder();

                buff.Append(String.Format(CultureInfo.CurrentCulture, "-----------------------------start({0})------------------------------------\n", DateTime.Now.ToString("o", System.Globalization.CultureInfo.CurrentCulture)));

                buff.Append(ex.Message + "\n");

                buff.Append("---------------------------------------End---------------------------------------\n");

                log.WriteLine(buff.ToString());
                log.Flush();
                log.Close();
                //return;
            }

        }
        #endregion

        private void SaveToAlarmLog(string ID, string  CaptureDate,string Capturetime ,string Interval, string SourceIP, string DestIP, string MessageBodyHex)
        {
            string path = string.Empty;
            path = "AlarmLog";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string head = ID + ",";
            head += Capturetime + ",";
            head += Interval + ",";
            head += SourceIP + ","  + DestIP + "," + MessageBodyHex.Replace("\r\n", string.Empty);
            //System.IO.Directory.CreateDirectory(path);
            path +=  "\\" + CaptureDate+".csv";
            //using (System.IO.StreamWriter file = new System.IO.StreamWriter("BytesFilesLogbk.txt", true))
            using (StreamWriter file = new StreamWriter(path, true))
            {
                file.WriteLine(head);// 直接追加文件末尾，换行   
                file.Close();
            }
        }
        //private bool DataFilter_IPandPort(string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort)
        //{
        //    int tmp = SettingSaves.Protocol.IndexOf(Protocol);
        //    if (SettingSaves.Protocol.IndexOf(Protocol) == -1) return false;

        //    string strSour = SourceIP + "-" + SourcePort;
        //    string strDest = DestIP + "-" + DestPort;

        //    if (!(SettingSaves.listSourIPandPort.Contains(strSour) || SettingSaves.listDestIPandPort.Contains(strDest))) return false;
        //    if (SettingSaves.DicIPPortProperty[strSour] == SettingSaves.DicIPPortProperty[strDest]) return false;


        //    return true;
        //}

        private bool DataFilter(string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort)
        {
            if (DataFilterByIPAndPort(Protocol, SourceIP, SourcePort, DestIP, DestPort))
            {
                return true;
            }
            return false;
        }

        private void UpdateStatus()
        {
            PacketStatus.Text = "收到数据包:" + FilterCount.ToString() + ", 筛选:" + ItemID.ToString();
        }

        /// <summary>
        /// 将二进制数据转换成16进制
        /// </summary>
        /// <returns></returns>
        private string GetDataHex(Byte[] Data, int index, int count)
        {
            string DataHex = "";
            for (int i = index; i < index + count; i++)
            {
                if (i > index && (i - index) % 16 == 0)
                {
                    DataHex += "\r\n";
                }
                if (Data[i].ToString("X").Length != 1)
                {
                    DataHex += Data[i].ToString("X") + " ";
                }
                else
                {
                    DataHex += "0" + Data[i].ToString("X") + " ";
                }
            }

            return DataHex;
        }

        private int GetMessageHeaderLen(string Protocol)
        {
            switch (Protocol)
            {
                case "ICMP": return ICMPDataOffset;
                case "IGMP": return IGMPDataOffset;
                case "TCP": return TCPDataOffset;
                case "UDP": return UDPDataOffset;
                case "SCTP": return SCTPDataOffset;
                case "UNKNOW": return 0;
                default: return 0;
            }
        }

        private string GetDataTxt(Byte[] Data, int index, int count)
        {
            Byte[] Temp = new Byte[Data.Length];

            Data.CopyTo(Temp, 0);

            for (int i = index; i < index + count; i++)
            {
                if (Temp[i] == 0)
                {
                    Temp[i] = 46;
                }
            }
            return System.Text.Encoding.Default.GetString(Temp, index, count);
        }

        private void Clear()
        {
            ItemID = 0;
            captureTime = string.Empty;
            FilterCount = 0;
            listView1.Items.Clear();
            listView_Data.Items.Clear();
            for (int i = 0; i < chartVtoZ1.Series.Count();i++ )
            {
                chartVtoZ1.Series[i].Points.Clear();
            }
            for (int i = 0; i < chartZtoV1.Series.Count(); i++)
            {
                chartZtoV1.Series[i].Points.Clear();
            }
            for (int i = 0; i < chartVtoZ2.Series.Count(); i++)
            {
                chartVtoZ2.Series[i].Points.Clear();
            }
            for (int i = 0; i < chartZtoV2.Series.Count(); i++)
            {
                chartZtoV2.Series[i].Points.Clear();
            }
            //chart1 = new Chart();
            //TextBox_Hex.Text = "";
            UpdateStatus();
        }

        //edit by shuya
        private List<CBTCByte> ReadCBTCXml()
        {
            XmlDocument doc = new XmlDocument();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreComments = true;
            XmlReader reader = XmlReader.Create(@"ZC--VOBC.xml", settings);
            doc.Load(reader);
            reader.Close();
            XmlNode root = doc.SelectSingleNode("GeneralPacket");
            XmlNodeList listXn = root.ChildNodes;

            List<CBTCByte> listNode = new List<CBTCByte>();

            foreach (XmlNode xn in listXn)
            {
                CBTCByte node = new CBTCByte();
                XmlElement xe = (XmlElement)xn;
                node.index = Convert.ToInt16(xe.GetAttribute("index"));
                node.name = xe.GetAttribute("name").ToString();
                node.len = Convert.ToInt16(xe.GetAttribute("len"));
                node.defValue = xe.GetAttribute("defValue").ToString();
                listNode.Add(node);
            }
            return listNode;
        }

        static private List<string> GetFiles(DateTime _startTime, DateTime _endTime, string SourceIP, string DestIP)
        {
            string path = string.Empty;
            List<string> filesList = new List<string>();
            #region 跨天查询
            if (_startTime.Year != DateTime.Now.Year || _startTime.Month != DateTime.Now.Month || DateTime.Now.Day - _startTime.Day > 7)
            {
                MessageBox.Show("只能查询7日内的数据，请重新输入！");
                return null;
            }
            if (_endTime.Year != DateTime.Now.Year || _endTime.Month != DateTime.Now.Month || DateTime.Now.Day - _endTime.Day > 7)
            {
                MessageBox.Show("只能查询7日内的数据，请重新输入！");
                return null;
            }

            if (_startTime > _endTime)//开始时间不能大于终止时间。
            {
                MessageBox.Show("开始时间大于终止时间！");
                return null;
            }
            else
            {
                try
                {
                    #region 将符合的txt文件存入数组中

                    #region 起止时间不是同一天

                    if (_startTime.Day != _endTime.Day)
                    {
                        int _day = _startTime.Day;
                        for (; _day <= _endTime.Day; _day++)
                        {
                            #region 起始那天
                            if (_day == _startTime.Day)
                            {
                                for (int h = _startTime.Hour; h < 24; h++)
                                {
                                    if (h == _startTime.Hour)
                                    {

                                        for (int m = _startTime.Minute; m < 60; m++)
                                        {
                                            path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM") + "-" + _day.ToString().PadLeft(2, '0') 
                                                    + "\\" + h.ToString().PadLeft(2, '0') + "-" + m.ToString().PadLeft(2, '0') + ".txt";
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    ReceiveTime recvTime = new ReceiveTime(line, path);
                                                    if (recvTime.Second >= _startTime.Second)
                                                    {
                                                        filesList.Add(line.ToString());
                                                    }
                                                }
                                                sr.Close();
                                            }
                                            catch (Exception e)
                                            {

                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 0; m < 60; m++)
                                        {
                                            path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM") + "-" + _day.ToString().PadLeft(2, '0')
                                                    + "\\" + h.ToString().PadLeft(2, '0') + "-" + m.ToString().PadLeft(2, '0') + ".txt";
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    filesList.Add(line.ToString());
                                                }
                                                sr.Close();
                                            }
                                            catch (Exception e)
                                            {

                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 终止那天
                            else if (_day == _endTime.Day)
                            {
                                for (int h = 0; h < _endTime.Hour; h++)
                                {
                                    if (h == _endTime.Hour)
                                    {
                                        for (int m = 0; m < _endTime.Minute; m++)
                                        {
                                            path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _endTime.ToString("yyyy-MM") + "-"
                                                    +_day.ToString().PadLeft(2, '0') + "\\" + h.ToString().PadLeft(2, '0') + "-" 
                                                    + m.ToString().PadLeft(2, '0') + ".txt";
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    ReceiveTime recvTime = new ReceiveTime(line, path);
                                                    if (recvTime.Second <= _endTime.Second)
                                                    {
                                                        filesList.Add(line.ToString());
                                                    }
                                                }
                                                sr.Close();
                                            }
                                            catch (Exception e)
                                            {

                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 0; m < 60; m++)
                                        {
                                            path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy-MM") + "-"
                                                +_day.ToString().PadLeft(2, '0') + "\\" + h.ToString().PadLeft(2, '0') + "-" 
                                                + m.ToString().PadLeft(2, '0') + ".txt";
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    filesList.Add(line.ToString());
                                                }
                                                sr.Close();
                                            }
                                            catch (Exception e)
                                            {

                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 中间天
                            else //中间天
                            {
                                for (int h = 0; h < 24; h++)
                                {
                                    for (int m = 0; m < 60; m++)
                                    {
                                        path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM") + "-"
                                            +_day.ToString().PadLeft(2, '0') + "\\" + h.ToString().PadLeft(2, '0') + "-" 
                                            + m.ToString().PadLeft(2, '0') + ".txt";

                                        try
                                        {
                                            StreamReader sr = new StreamReader(path, Encoding.Default);
                                            String line;
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                filesList.Add(line.ToString());
                                            }
                                            sr.Close();
                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                }
                            }
                            #endregion
                            //fileArrayList.AddRange(Directory.GetFiles(Application.StartupPath + @"\Log" + @"\" + path + @"\" + _startTime.Year + @"\" + _month));  //将文件名存入数组
                        }
                    }
                    #endregion

                    #region 开始和结束是同一天
                    else //开始和结束是同一天，在一个文件夹里面找
                    {
                        string strCatDate = _startTime.ToString("yyyy-MM-dd");
                        #region 同一个小时
                        if (_startTime.Hour == _endTime.Hour)
                        {
                            if (_startTime.Minute == _endTime.Minute)//同分钟，在一个文件里面查找
                            {
                                string strCatTime = _startTime.ToString("HH-mm");
                                path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM-dd") + "\\" + strCatTime + ".txt";
                                try
                                {
                                    StreamReader sr = new StreamReader(path, Encoding.Default);
                                    String line;
                                    while ((line = sr.ReadLine()) != null)
                                    {
                                        ReceiveTime recvTime = new ReceiveTime(line, path);
                                        if (recvTime.Second >= _startTime.Second && recvTime.Second <= _endTime.Second)
                                        {
                                            filesList.Add(line.ToString());
                                        }
                                    }
                                    sr.Close();
                                }
                                catch (Exception e)
                                { }
                            }
                            else //同小时不同分钟
                            {

                                for (int min = _startTime.Minute; min <= _endTime.Minute; min++)
                                {
                                    path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM-dd") + "\\" +  _startTime.Hour + "-" + min + ".txt";
                                    try
                                    {
                                        StreamReader sr = new StreamReader(path, Encoding.Default);
                                        String line;
                                        while ((line = sr.ReadLine()) != null)
                                        {
                                            ReceiveTime recvTime = new ReceiveTime(line, path);
                                            if (min == _startTime.Minute)
                                            {
                                                if (recvTime.Second >= _startTime.Second)
                                                {
                                                    filesList.Add(line.ToString());
                                                }
                                            }
                                            else if (min == _endTime.Minute)
                                            {
                                                if (recvTime.Second <= _endTime.Second)
                                                {
                                                    filesList.Add(line.ToString());
                                                }
                                            }
                                            else
                                            {
                                                filesList.Add(line.ToString());
                                            }
                                        }
                                        sr.Close();
                                    }
                                    catch (Exception e)
                                    { }
                                }
                            }
                        }
                        #endregion

                        #region 不同小时
                        else
                        {
                            for (int hour = _startTime.Hour; hour <= _endTime.Hour; hour++)
                            {
                                if (hour == _startTime.Hour)
                                {
                                    for (int m = _startTime.Minute; m < 60; m++)
                                    {
                                        path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM-dd") + "\\" + hour.ToString().PadLeft(2, '0') + "-" + m.ToString().PadLeft(2, '0') + ".txt";
                                        if (m == _startTime.Minute)
                                        {
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    ReceiveTime recvTime = new ReceiveTime(line, path);
                                                    if (recvTime.Second >= _startTime.Second)
                                                    {
                                                        filesList.Add(line.ToString());
                                                    }
                                                }
                                                sr.Close();
                                            }
                                            catch (Exception e)
                                            { }
                                        }
                                        else
                                        {
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    ReceiveTime recvTime = new ReceiveTime(line, path);
                                                    filesList.Add(line.ToString());

                                                }
                                                sr.Close();
                                            }
                                            catch (Exception e)
                                            { }
                                        }

                                    }
                                }
                                else if (hour == _endTime.Hour)
                                {
                                    for (int m = 0; m <= _endTime.Minute; m++)
                                    {
                                        path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _endTime.ToString("yyyy") + "\\" + _endTime.ToString("MM-dd") + "\\" 
                                                + hour.ToString().PadLeft(2, '0') + "-" + m.ToString().PadLeft(2, '0') + ".txt";
                                        if (m == _endTime.Minute)
                                        {
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    ReceiveTime recvTime = new ReceiveTime(line, path);
                                                    if (recvTime.Second <= _endTime.Second)
                                                    {
                                                        filesList.Add(line.ToString());
                                                    }
                                                    sr.Close();
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                        }
                                        else
                                        {
                                            try
                                            {
                                                StreamReader sr = new StreamReader(path, Encoding.Default);
                                                String line;
                                                while ((line = sr.ReadLine()) != null)
                                                {
                                                    ReceiveTime recvTime = new ReceiveTime(line, path);
                                                    filesList.Add(line.ToString());
                                                }
                                                sr.Close();

                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    for (int m = 0; m < 60; m++)
                                    {
                                        path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") +"\\" + _startTime.ToString("MM-dd")+ "\\" + hour.ToString().PadLeft(2, '0') + "-" + m.ToString().PadLeft(2, '0') + ".txt";
                                        try
                                        {
                                            StreamReader sr = new StreamReader(path, Encoding.Default);
                                            String line;
                                            while ((line = sr.ReadLine()) != null)
                                            {
                                                filesList.Add(line.ToString());
                                            }
                                            sr.Close();
                                        }
                                        catch (Exception ex)
                                        {

                                        }

                                    }
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion

                    #endregion
                }
                catch (Exception ex)
                {
                    //return null;
                }
            }
            #endregion


            return filesList;
        }


        #endregion

        //基于事件触发的方法，当接收到数据包时执行
        void SnifferSocket_PacketArrival(object Sender, RawSocket.PacketArrivedEventArgs args)
        {
            this.BeginInvoke(new refresh(Data_Receive), args.Date, args.Protocol, args.OriginationAddress, args.OriginationPort, args.DestinationAddress, args.DestinationPort, args.IPHeaderLength, args.IPHeaderBuffer, args.MessageLength, args.MessageBuffer, args.PacketLength, args.PacketBuffer,args.ReceiveBuf,args.BufLength);
        }


        #region 删除文件线程
        /// <summary>
        /// 删除文件线程。
        /// </summary>
        private void DelFileThread()
        {
            string pathConfigurationInfo = Application.StartupPath + @"\DeleteDataInterval.txt";
            StreamReader objReader = new StreamReader(pathConfigurationInfo, Encoding.GetEncoding("gb2312"));
            string interval = objReader.ReadLine();
            DeleteDataInterval = Convert.ToUInt16(interval);
            while (m_runningFlg)
            {
                DeleteOverdueData(DeleteDataInterval);

                // 30秒执行一次删除操作。
                Thread.Sleep(300000);
            }
        }
        #endregion

        #region 程序载入
        private void Form1_Load(object sender, EventArgs e)
        {
            m_runningFlg = true;

            //启动删除文件线程
            Thread innerDelFileThread = new Thread(DelFileThread);
            innerDelFileThread.IsBackground = true;
            innerDelFileThread.Start();
            return;
        }
        #endregion



    }
    public class ReceiveTime
    {
        private int _year;//初始化变量
        private int _month;//初始化变量
        private int _day;//初始化变量
        private int _hour;//初始化变量
        private int _minute;//初始化变量
        private int _second;//初始化变量

        #region 属性
        public int Year
        {
            set { _year = value; }//可写
            get { return _year; }//可读
        }
        public int Month
        {
            set { _month = value; }//可写
            get { return _month; }//可读
        }
        public int Day
        {
            set { _day = value; }//可写
            get { return _day; }//可读
        }
        public int Hour
        {
            set { _hour = value; }//可写
            get { return _hour; }//可读
        }
        public int Minute
        {
            set { _minute = value; }//可写
            get { return _minute; }//可读
        }
        public int Second
        {
            set { _second = value; }//可写
            get { return _second; }//可读
        }
        #endregion

        //构造函数
        public ReceiveTime(string line, string path)
        {
            string[] pathBytes = path.Split(new char[] { '\\' });
            string[] dateBytes = pathBytes[3].Split(new char[] { '-' });
            _year = Convert.ToInt32(pathBytes[2]);
            _month = Convert.ToInt32(dateBytes[0]);
            _day = Convert.ToInt32(dateBytes[1]);

            string[] lineBytes = line.Split(new char[] { ',' });
            string[] timeBytea = lineBytes[1].Split(new char[] { ':' });
            _hour = Convert.ToInt32(timeBytea[0]);
            _minute = Convert.ToInt32(timeBytea[1]);
            _second = Convert.ToInt32(timeBytea[2]);
        }
    }

}
