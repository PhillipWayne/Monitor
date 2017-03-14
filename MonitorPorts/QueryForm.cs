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

using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.IO;
using System.Windows.Forms.DataVisualization.Charting;

namespace MonitorPorts
{
    public partial class QueryForm : Form
    {
        #region 全局变量
        private int ItemID;
        private double captureTime;
        private int FilterCount;
        DateTime startTime;
        //private OpenAnalysis OpenAnalysis;
        private SettingSaves SettingSaves;

        delegate void listViewDelegate(string ID, string Capturetime, string Protocol, string SourceIP, string SourcePort, string DestIP, string DestPort, string AllLength, string MessageBodyTxt, string MessageBodyLen, string MessageBodyHex);
        Dictionary<string, string> dic;
        #endregion

        #region 构造函数
        public QueryForm()
        {
            ItemID = 0;
            FilterCount = 0;
            SettingSaves = new MonitorPorts.SettingSaves();

            InitializeComponent();
            dic = ReadIPlist();
            ShowInForm(dic);
        }

        #endregion

        #region 窗体控件

        //导出EXCEL
        private void Button_OutputExcel_Click(object sender, EventArgs e)
        {
            ExportToExecl();
        }
                        
        //清空
        private void Button_Clear_Click(object sender, EventArgs e)
        {
            Clear();
        }

        //关闭界面
        private void QueryForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            MainForm fm = new MainForm(); //实例一个新窗口；
            this.Dispose();   //释放内存
            fm.ShowDialog();
        }

        //查询
        private void Button_Filter_Click(object sender, EventArgs e)
        {
            Clear();
            DateTime startData = startDate.Value.Date; // dateTimePickerStart.Value.Date;//开始日期

            DateTime endData = endDate.Value.Date;//终止日期

            DateTime _startTime = startData + new TimeSpan(startHour.Value.Hour, startHour.Value.Minute, startHour.Value.Second);//开始时间
            //startData + new TimeSpan(QueryFilterOptionForm._startTime.Hour, QueryFilterOptionForm._startTime.Minute, QueryFilterOptionForm._startTime.Second); ;//开始时间
            DateTime _endTime = endData + new TimeSpan(endHour.Value.Hour, endHour.Value.Minute, endHour.Value.Second); ;//终止时间
            //endData + new TimeSpan(QueryFilterOptionForm._endTime.Hour, QueryFilterOptionForm._endTime.Minute, QueryFilterOptionForm._endTime.Second); ;//终止时间
            string sourIP = string.Empty;
            for (int i = 0; i < SourceListBox1.Items.Count; i++)
            {
                if (SourceListBox1.GetItemChecked(i))
                {
                    sourIP = dic[SourceListBox1.Items[i].ToString()];
                }
            }
            string destIP = string.Empty;
            for (int i = 0; i < DestListBox2.Items.Count; i++)
            {
                if (DestListBox2.GetItemChecked(i))
                {
                    destIP = dic[DestListBox2.Items[i].ToString()];
                }
            }

            List<string> fileList = new List<string>();
            fileList = GetFiles(_startTime, _endTime, sourIP, destIP);

            if (fileList == null || fileList.Count == 0)
            {
                MessageBox.Show("没有找到符合条件的文件！");
                return;
            }

            else
            {
                foreach (string strFile in fileList)
                {
                    //listView_Data.Items.Add(new ListViewItem(new string[] { ID, Capturetime.ToString("HH:mm:ss:fff"), Interval, Protocol, SourceIP, SourcePort, DestIP, DestPort, AllLength, MessageBodyLen, MessageBodyTxt, MessageBodyHex }));
                    string[] tmp = strFile.Split(',');
                    listView_Data.Items.Add(new ListViewItem(strFile.Split(',')));

                    //画图
                    this.Chart_Offline.Series[0].Points.AddY(Convert.ToDouble(tmp[2]));
                    this.Chart_Offline.Series[0].ChartType = SeriesChartType.Line;
                    Chart_Offline.Series[0].BorderWidth = 3;
                    Chart_Offline.Series[0].Color = Color.Blue;
                    Update();

                }
            }
        }

        //选中列表中某一列
        private void listView_Data_ItemSelectionChanged_1(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            TextBox_Hex.Text = e.Item.SubItems[10].Text;

        }

        #endregion

        #region 内部方法
        /// <summary>
        /// 执行导出数据
        /// </summary>
        private void ExportToExecl()
        {
            saveFileDialog1.Filter = "Excel文件|*.xls|所有文件|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (!System.String.IsNullOrEmpty(saveFileDialog1.FileName))
                {
                    string ExcelFileName = saveFileDialog1.FileName;

                    DoExport(ExcelFileName);
                }
            }
        }

        /// <summary>
        /// 具体导出的方法
        /// </summary>
        /// <param name="listView">ListView</param>
        /// <param name="strFileName">导出到的文件名</param>
        private void DoExport(string strFileName)
        {
            int rowNum = listView_Data.Items.Count;
            int columnNum = listView_Data.Columns.Count;
            //int columnNum = listView_Data.Items[0].SubItems.Count;
            int rowIndex = 1;
            int columnIndex = 0;
            if (rowNum == 0 || string.IsNullOrEmpty(strFileName))
            {
                return;
            }
            if (rowNum > 0)
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("无法创建excel对象，可能您的系统没有安装excel");
                    return;
                }
                xlApp.DefaultFilePath = "";
                xlApp.DisplayAlerts = true;
                xlApp.SheetsInNewWorkbook = 1;
                Microsoft.Office.Interop.Excel.Workbook xlBook = xlApp.Workbooks.Add(true);
                //将ListView的列名导入Excel表第一行
                foreach (ColumnHeader dc in listView_Data.Columns)
                {
                    columnIndex++;
                    xlApp.Cells[rowIndex, columnIndex] = dc.Text;
                }
                //将ListView中的数据导入Excel中
                for (int i = 0; i < rowNum; i++)
                {
                    rowIndex++;
                    columnIndex = 0;
                    for (int j = 0; j < columnNum; j++)
                    {
                        columnIndex++;
                        //注意这个在导出的时候加了“\t” 的目的就是避免导出的数据显示为科学计数法。可以放在每行的首尾。
                        xlApp.Cells[rowIndex, columnIndex] = Convert.ToString(listView_Data.Items[i].SubItems[j].Text) + "\t";
                    }
                }
                //例外需要说明的是用strFileName,Excel.XlFileFormat.xlExcel9795保存方式时 当你的Excel版本不是95、97 而是2003、2007 时导出的时候会报一个错误：异常来自 HRESULT:0x800A03EC。 解决办法就是换成strFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal。
                xlBook.SaveAs(strFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlApp = null;
                xlBook = null;
                MessageBox.Show("OK");
            }
        }

        //界面底部数字更新
        private void UpdateStatus()
        {
            PacketStatus.Text = "收到数据包:" + ItemID.ToString() + ", 筛选:" + FilterCount.ToString();
        }        

        /// <summary>
        /// 查找报文
        /// </summary>
        /// <param name="_startTime">起始时间</param>
        /// <param name="_endTime">终止时间</param>
        /// <param name="SourceIP">源IP</param>
        /// <param name="DestIP">目的IP</param>
        /// <returns>报文列表</returns>
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
                                                    + _day.ToString().PadLeft(2, '0') + "\\" + h.ToString().PadLeft(2, '0') + "-"
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
                                                + _day.ToString().PadLeft(2, '0') + "\\" + h.ToString().PadLeft(2, '0') + "-"
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
                                            + _day.ToString().PadLeft(2, '0') + "\\" + h.ToString().PadLeft(2, '0') + "-"
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
                                    path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM-dd") + "\\" + _startTime.Hour + "-" + min + ".txt";
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
                                        path = "BytesFilesLog" + "\\" + SourceIP + "--" + DestIP + "\\" + _startTime.ToString("yyyy") + "\\" + _startTime.ToString("MM-dd") + "\\" + hour.ToString().PadLeft(2, '0') + "-" + m.ToString().PadLeft(2, '0') + ".txt";
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

        /// <summary>
        /// 将IPlist信息显示在本界面的两个ListBox中
        /// </summary>
        /// Dictionary<IP对应名称,IP-端口>
        /// <param name="dic"></param>
        private void ShowInForm(Dictionary<string, string> dic)
        {
            foreach (string str in dic.Keys)
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
                string IP = strs[1];
                tmpdic.Add(strs[0], IP);
                //listIP.Add(strs[1]);
            }
            return tmpdic;
        }        

        //清空
        private void Clear()
        {
            listView_Data.Items.Clear();
            for (int i = 0; i < Chart_Offline.Series.Count; i++)
            {
                Chart_Offline.Series[i].Points.Clear();
            }
            ItemID = 0;
            startTime = DateTime.Now;
            captureTime = 0;
            FilterCount = 0;
            TextBox_Hex.Text = "";
            UpdateStatus();
        }


        #endregion

        private void SourceListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            for(int i=0;i<SourceListBox1.Items.Count;i++)
            {
                SourceListBox1.SetItemChecked(i,false);
            }
            if(SourceListBox1.CheckedItems==null)
            {
                SourceListBox1.SetItemChecked(SourceListBox1.SelectedIndex, false);
            }
            else
            {
                SourceListBox1.SetItemChecked(SourceListBox1.SelectedIndex, true);
            }
        }

        private void DestListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < DestListBox2.Items.Count; i++)
            {
                DestListBox2.SetItemChecked(i, false);
            }
            if (DestListBox2.CheckedItems == null)
            {
                DestListBox2.SetItemChecked(DestListBox2.SelectedIndex, false);
            }
            else
            {
                DestListBox2.SetItemChecked(DestListBox2.SelectedIndex, true);
            }
        }

    }
}
