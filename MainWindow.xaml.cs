using ExportExcel;
using GalaSoft.MvvmLight.CommandWpf;
using LiveCharts;
using LiveCharts.Wpf;
using Microsoft.Win32;
using OfficeOpenXml;
using ReportPro.img;
using ReportPro.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ReportPro
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public DataSet dataSet;
        public DataSet dataSetAts;
        public int CourrentBadNumber = 0;
        public int cdata { get { return CourrentBadNumber; } set { CourrentBadNumber = value; OnPropertyChanged("cdata"); } }
        public SeriesCollection seriesCollection { get; set; }
        string K = "", m = "";
        public int ateMax;
        public int AteMax { get { return ateMax; } set { ateMax = value; OnPropertyChanged("AteMax"); } }
        public int actualQuntity;
        public int ActualQuntity { get { return actualQuntity; } set { actualQuntity = value; OnPropertyChanged("ActualQuntity"); } }
        public int atsMax;
        public int AtsMax { get { return atsMax; } set { atsMax = value; OnPropertyChanged("AtsMax"); } }
        public DataView BadnessDataSet;
        public DataView badnessDataSet { get { return BadnessDataSet; } set { BadnessDataSet = value; OnPropertyChanged("badnessDataSet"); } }
        public string[] BadTime = { DateTime.Now.AddDays(-6).ToString("d"), DateTime.Now.AddDays(-5).ToString("d"), DateTime.Now.AddDays(-4).ToString("d"), DateTime.Now.AddDays(-3).ToString("d"), DateTime.Now.AddDays(-2).ToString("d"), DateTime.Now.AddDays(-1).ToString("d"), DateTime.Now.AddDays(+1).ToString("d") };
        public DataSet testset;
        public List<int> numberAte;
        public List<string> timeAte;
        public string Informcontent;
        public string InformContent { get { return Informcontent; } set { Informcontent = value; OnPropertyChanged("Informcontent"); } }
        public SeriesCollection seriesCollectionAts { get; set; }
        public RelayCommand AteReset { get; set; }

        public RelayCommand AtsReset { get; set; }

        public Thread thread;

        public DataView Orderdata;

        public DataView orderdata { get { return Orderdata; } set { Orderdata = value; OnPropertyChanged("orderdata"); } }

        public RelayCommand ClickSelect { get; set; }

        public RelayCommand SingleTrial { get; set; }

        public List<string> po = new List<string>();

        public DataView podataView1;

        public DataView PodataView1 { get { return podataView1; } set { podataView1 = value; OnPropertyChanged("podataView1"); } }

        public int pocount { get; set; }

        public int Pocount { get { return pocount; } set { pocount = value; OnPropertyChanged("pocount"); } }

        public decimal unitCost { get; set; }

        public decimal UnitCost
        {
            get { return unitCost; }
            set { if (value < 0) throw new ArgumentException("工單不能為空！"); else { unitCost = value; } OnPropertyChanged("unitCost"); }
        }
        public MainWindow()
        {
            InitializeComponent();
            UpdateProcedure();
            DataContext = this;
            MultithreadigStart();
            AteReset = new RelayCommand(AteResetdata);
            AtsReset = new RelayCommand(AtsResetdata);
            ClickSelect = new RelayCommand(Orderselect);
            content.Content = new AbnormalNotice();
            thread = new Thread(selectdata);
            SingleTrial = new RelayCommand(SingleTrialClick);
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public event EventHandler CanExecuteChanged;

        public void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
            }
            else
            {
                WindowState = WindowState.Maximized;
            }
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            this.Close();
            Process.GetCurrentProcess().Kill();
        }
        private void DataGrid_Loaded()
        {
            Dispatcher.Invoke(new Action(delegate { matec.Visibility = Visibility.Visible; }));
            int a = 0, b = 0, c = 0, d = 0, y = 0, f = 0;
            List<string> data = new List<string>();
            Dispatcher.Invoke(new Action(delegate { a = WO.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { b = Part.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { c = EndNumber.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { d = InitialNumber.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { y = StartDate.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { f = DateClosed.Text.Length; }));
            if (a > 0)
            {
                Dispatcher.Invoke(new Action(delegate { data.Add("and a.work_order ='" + WO.Text + "'"); }));
            }
            else
            {
                data.Add("");
            }
            if (b > 0)
            {
                Dispatcher.Invoke(new Action(delegate { data.Add("and a.part_no = '" + Part.Text + "'"); }));
            }
            else
            {
                data.Add("");
            }
            if (c > 0)
            {
                Dispatcher.Invoke(new Action(delegate { data.Add("and a.serial_number >=  '" + InitialNumber.Text + "'"); }));
            }
            else
            {
                data.Add("");
            }
            if (d > 0)
            {
                Dispatcher.Invoke(new Action(delegate { data.Add("and a.serial_number <= '" + EndNumber.Text + "'"); }));
            }
            else
            {
                data.Add("");
            }
            if (y > 0)
            {
                Dispatcher.Invoke(new Action(delegate { data.Add("and a.update_time >= to_date('" + Convert.ToDateTime(StartDate.Text).ToString("yyy-MM-dd") + "','yyyy-mm-dd')"); }));
            }
            else
            {
                data.Add("");
            }
            if (f > 0)
            {
                Dispatcher.Invoke(new Action(delegate { data.Add("and a.update_time <= to_date('" + Convert.ToDateTime(DateClosed.Text).ToString("yyy-MM-dd") + "','yyyy-mm-dd')"); }));
            }
            else
            {
                data.Add("");
            }
            string SQL = String.Format("SELECT A.WORK_ORDER, H.PART_NO,B.CUSTOMER_SN,A.SERIAL_NUMBER,J.TEST_ITEM,J.FREQUENCY,J.VOLTAGE,J.VOUT,J.VOUT_ID,NVL(J.LOAD, J.SPC_DESC) LOAD_2,J.USL,J.LSL,J.SPC_ITEM,A.SPC_VALUE,A.SPC_RESULT,D.PDLINE_NAME,F.PROCESS_NAME,G.TERMINAL_NAME,EMP_NAME,A.UPDATE_TIME,E.STAGE_NAME,J.PC_INDEX FROM sajet.G_SPC_TM A,SAJET.G_SN_STATUS B,SAJET.SYS_TERMINAL C,SAJET.SYS_PDLINE D,SAJET.SYS_STAGE E,SAJET.SYS_PROCESS F,SAJET.SYS_TERMINAL G,SAJET.SYS_PART H,SAJET.SYS_EMP I,SAJET.SYS_SPC J WHERE 1 = 1 AND A.SERIAL_NUMBER = B.SERIAL_NUMBER AND A.TERMINAL_ID = C.TERMINAL_ID AND A.PDLINE_ID = D.PDLINE_ID AND A.STAGE_ID = E.STAGE_ID AND A.PROCESS_ID = F.PROCESS_ID AND A.TERMINAL_ID = G.TERMINAL_ID AND A.PART_ID = H.PART_ID AND A.EMP_ID = I.EMP_ID AND A.SPC_ID = J.SPC_ID AND A.PROCESS_ID = '100024' {0} \t {1} \t {2} \t {3} \t {4} \t {5}   AND (NVL (USL, 0) <> 0 OR NVL (LSL, 0) <> 0) ORDER BY A.SERIAL_NUMBER, J.PC_INDEX, A.UPDATE_TIME", data[0], data[1], data[2], data[3], data[4], data[5]);
            DataSet dataSet = Oraclecnn.OracleCnn(SQL);
            Dispatcher.Invoke(new Action(delegate { Items1.ItemsSource = dataSet.Tables[0].DefaultView; }));
            Dispatcher.Invoke(new Action(delegate { Sum.Text = dataSet.Tables[0].Rows.Count.ToString(); }));
            Dispatcher.Invoke(new Action(delegate { matec.Visibility = Visibility.Collapsed; }));
        }

        private void TextBlock_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Thread thread = new Thread(DataGrid_Loaded);
            thread.Start();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            cpk.Visibility = Visibility.Collapsed;
            record.Visibility = Visibility.Collapsed;
            SingleTrialDrid.Visibility = Visibility.Collapsed;
            YieldRat.Visibility = Visibility.Visible;
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            cpk.Visibility = Visibility.Visible;
            record.Visibility = Visibility.Collapsed;
            SingleTrialDrid.Visibility = Visibility.Collapsed;
            YieldRat.Visibility = Visibility.Collapsed;
        }

        private void Select2_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Thread thread = new Thread(Getdatagrid1);
            thread.Start();
        }
        private void Getdatagrid1()
        {

            List<string> str = fieldscreen();
            string strin = "";
            foreach (var item in str)
            {
                strin += " " + item;
            }
            string sql = @"Select D.Process_Name ""製程名稱"", B.Work_Order ""工單號碼"",B.REMARK as ""工單備註"",B.WO_OPTION2 AS ""工廠備註"", E.PART_NO ""成品料號"", F.MODEL_NAME ""機種名稱"", B.TARGET_QTY ""工單目標數"", B.INPUT_QTY ""工單投入"", B.OUTPUT_QTY ""工單產出數"",   SUM(PASS_QTY) ""良品數"", SUM(FAIL_QTY) ""不良品數"", SUM(REPASS_QTY) ""回流良品"", SUM(REFAIL_QTY) ""回流不良"",  SUM(A.OUTPUT_QTY)  ""製程產出數"", Round(nvl(SUM(PASS_QTY)/decode(SUM(PASS_QTY+FAIL_QTY),'0',null,SUM(PASS_QTY+FAIL_QTY)),0)*100,2)||'%' ""良率"" From SAJET.G_SN_COUNT A,SAJET.SYS_PDLINE C,SAJET.G_WO_BASE B,SAJET.SYS_PROCESS D,SAJET.SYS_PART E,SAJET.SYS_MODEL F,SAJET.SYS_FACTORY L where 1=1  " + strin + "  and  B.DEFAULT_PDLINE_ID = C.PDLINE_ID and A.WORK_ORDER = B.WORK_ORDER and A.PROCESS_ID = D.PROCESS_ID and A.PART_ID = E.PART_ID and E.MODEL_ID = F.MODEL_ID(+) AND B.FACTORY_ID = L.FACTORY_ID(+) AND D.PROCESS_NAME NOT IN ('IPQC','OQC') Group By D.Process_Name,D.Process_Code,B.Work_Order,E.PART_NO,F.MODEL_NAME,B.TARGET_QTY,B.INPUT_QTY,B.OUTPUT_QTY,B.REMARK,B.WO_OPTION2 Order By B.Work_Order desc,D.Process_Code,D.Process_Name";
            DataSet = Oraclecnn.OracleCnn(sql);
            Dispatcher.Invoke(new Action(delegate { YieldRatLordTabel.ItemsSource = DataSet.Tables[0].DefaultView; }));
            Dispatcher.Invoke(new Action(delegate { Sum2.Text = DataSet.Tables[0].Rows.Count.ToString(); }));
            Dispatcher.Invoke(new Action(delegate { ProgressBar2.Visibility = Visibility.Collapsed; }));
        }
        private void YieldRatLordTabel_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            ProgressBar3.Visibility = Visibility.Visible;
            Thread thread = new Thread(Getdatagrid2);
            thread.Start();
        }
        private void ComboWorkorder_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (ComboWorkorder.Items.Count < 1)
            {
                string sql = "select work_order from sajet.g_wo_base";
                DataSet dataSet = Oraclecnn.OracleCnn(sql);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    ComboWorkorder.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
        }
        private void combopart_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (combopart.Items.Count < 1)
            {
                string sql = "select part_no from sajet.sys_part";
                DataSet dataSet = Oraclecnn.OracleCnn(sql);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    combopart.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
        }
        private void combomodel_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (combomodel.Items.Count < 1)
            {
                string sql = "select model_name from sajet.sys_model";
                DataSet dataSet = Oraclecnn.OracleCnn(sql);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    combomodel.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
        }
        private void workstation_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (comboworkstation.Items.Count < 1)
            {
                string sql = "select process_name from sajet.sys_process";
                DataSet dataSet = Oraclecnn.OracleCnn(sql);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    comboworkstation.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
        }
        private void comboprodustionline_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (comboprodustionline.Items.Count < 1)
            {
                string sql = "select pdline_name from SAJET.sys_pdline where enabled = 'Y' ";
                DataSet dataSet = Oraclecnn.OracleCnn(sql);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    comboprodustionline.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
        }
        private void combofactory_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (combofactory.Items.Count < 1)
            {
                string sql = "select factory_code from sajet.sys_factory";
                DataSet dataSet = Oraclecnn.OracleCnn(sql);
                for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
                {
                    combofactory.Items.Add(dataSet.Tables[0].Rows[i][0].ToString());
                }
            }
        }
        private List<string> fieldscreen()
        {
            int a = 0;
            int b = 0;
            int c = 0;
            int d = 0;
            int e = 0;
            int f = 0;
            int g = 0;
            int h = 0;
            Dispatcher.Invoke(new Action(delegate { a = combofactory.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { b = comboprodustionline.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { c = ComboWorkorder.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { d = combopart.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { e = combomodel.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { f = comboworkstation.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { g = starttime.Text.Length; }));
            Dispatcher.Invoke(new Action(delegate { h = endtime.Text.Length; }));
            List<string> list = new List<string>();
            if (a > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and l.FACTORY_CODE = '" + combofactory.Text + "'"); }));
            }
            else
            {
                list.Add("");
            }
            if (b > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and pdline_name = '" + comboprodustionline.Text + "'"); }));
            }
            else
            {
                list.Add("");
            }
            if (c > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and b.work_order = '" + ComboWorkorder.Text + "'"); }));
            }
            else
            {
                list.Add("");
            }
            if (d > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and e.part_no = '" + combopart.Text + "' "); }));
            }
            else
            {
                list.Add("");
            }
            if (e > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and f.model_name = '" + combomodel.Text + "'"); }));
            }
            else
            {
                list.Add("");
            }
            if (f > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and d.process_name in '" + comboworkstation.Text + "' "); }));
            }
            else
            {
                list.Add("");
            }
            if (g > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and TO_DATE(A.WORK_DATE,'yyyy/mm/dd')>= to_date('" + starttime.Text + "','yyyy/mm/dd') "); }));
            }
            else
            {
                list.Add("");
            }
            if (h > 0)
            {
                Dispatcher.Invoke(new Action(delegate { list.Add("and TO_DATE(A.WORK_DATE,'yyyy/mm/dd') <= to_date('" + endtime.Text + "','yyyy/mm/dd')"); }));
            }
            else
            {
                list.Add("");
            }
            return list;
        }
        private void Getdatagrid2()
        {
            int l = 0;
            Dispatcher.Invoke(new Action(delegate { l = YieldRatLordTabel.SelectedItems.Count; }));
            try
            {
                if (l >= 1)
                {
                    string s = "";
                    string d = "";
                    string c = "";
                    Dispatcher.Invoke(new Action(delegate { K = starttime.Text; }));
                    Dispatcher.Invoke(new Action(delegate { m = endtime.Text; }));
                    Dispatcher.Invoke(new Action(delegate { s = (YieldRatLordTabel.Columns[0].GetCellContent(YieldRatLordTabel.SelectedItems[0]) as TextBlock).Text; }));
                    Dispatcher.Invoke(new Action(delegate { d = (YieldRatLordTabel.Columns[1].GetCellContent(YieldRatLordTabel.SelectedItems[0]) as TextBlock).Text; }));
                    Dispatcher.Invoke(new Action(delegate { c = (YieldRatLordTabel.Columns[4].GetCellContent(YieldRatLordTabel.SelectedItems[0]) as TextBlock).Text; }));
                    string sql = @"Select A.WORK_ORDER ""工單號碼"",n.remark""工單備註"",E.Part_No ""成品料號"",A.SERIAL_NUMBER ""生產序號"",A.CUSTOMER_SN ""客戶序號"", SAJET.Sj_Snstatus_Result(A.CURRENT_STATUS)  ""狀態"", O.WEIGHT ""彩盒重量"",E.OPTION11 ""彩盒重量下限"",E.OPTION12 ""彩盒重量上限"", P.WEIGHT ""外箱重量"",E.OPTION9 ""外箱重量下限"",E.OPTION10 ""外箱重量上限"", C.PDLINE_NAME ""線別"",D.PROCESS_NAME ""製程名稱"", TO_CHAR(A.OUT_PROCESS_TIME, 'YYYY/MM/DD HH24:MI:SS') ""過站時間"", F.EMP_NAME ""作業人員"" ,decode(A.CURRENT_STATUS,'0','',decode(I.DEFECT_ID,'10000063',nvl(sajet.culpritdata(A.SERIAL_NUMBER), I.DEFECT_DESC),I.DEFECT_DESC)) ""不良現象描述"" ,decode(A.CURRENT_STATUS,'0','',K.REASON_DESC)  ""不良原因"",decode(A.CURRENT_STATUS,'0','',K.REASON_DESC2)  ""不良原因描述2"",decode(A.CURRENT_STATUS,'0','',L.DUTY_DESC)  ""責任歸屬"",decode(A.CURRENT_STATUS,'0','',M.lOCATION)  ""不良位置"",decode(A.CURRENT_STATUS,'0','',M.ITEM_NO)  ""零件代碼"" From SAJET.G_SN_TRAVEL A,SAJET.SYS_PDLINE C,SAJET.SYS_PROCESS D,SAJET.SYS_EMP F,SAJET.SYS_PART E,SAJET.SYS_MODEL G,SAJET.G_SN_DEFECT H,SAJET.SYS_DEFECT I,SAJET.G_SN_REPAIR J,SAJET.SYS_REASON K,SAJET.SYS_DUTY L,SAJET.G_SN_REPAIR_LOCATION M,SAJET.G_WO_BASE N,SAJET.G_BOX_WEIGHT O,SAJET.G_CARTON_WEIGHT P Where  0=0 and D.PROCESS_NAME = '" + s + "' and A.Work_Order = '" + d + "' and E.PART_NO = '" + c + "'and OUT_PROCESS_TIME >= TO_DATE('" + K + "','yyyy/mm/dd') and out_process_time <= to_date('" + m + "','yyyy/mm/dd')  and A.PDLINE_ID = C.PDLINE_ID(+) and A.PROCESS_ID = D.PROCESS_ID(+) and A.PART_ID = E.PART_ID(+) and E.MODEL_ID = G.MODEL_ID(+) and A.EMP_ID = F.EMP_ID(+) and A.SERIAL_NUMBER=H.SERIAL_NUMBER(+) and A.PROCESS_ID=H.PROCESS_ID(+) and A.TERMINAL_ID=H.TERMINAL_ID(+) AND A.OUT_PROCESS_TIME=H.REC_TIME(+) and H.DEFECT_ID=I.DEFECT_ID(+) AND H.RECID=J.RECID(+) and J.REASON_ID = K.REASON_ID(+) and J.DUTY_ID=L.DUTY_ID(+) and a.work_order=n.work_order(+) AND H.RECID=M.RECID(+) and A.BOX_NO=O.BOX(+) and A.CARTON_NO=P.CARTON(+) ";
                    DataSet dataSet = Oraclecnn.OracleCnn(sql);
                    Dispatcher.Invoke(new Action(delegate { YieldRatAssistantTable.ItemsSource = dataSet.Tables[0].DefaultView; }));
                    Dispatcher.Invoke(new Action(delegate { ProgressBar3.Visibility = Visibility.Collapsed; }));
                }
                else
                {
                    Dispatcher.Invoke(new Action(delegate { ProgressBar3.Visibility = Visibility.Collapsed; }));
                }
            }
            catch
            { }
        }

        private void Derive2_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (YieldRatLordTabel.Items.Count == 0 || thread.IsAlive)
            {
                Dispatcher.Invoke(new Action(delegate { snacktext.Content = "未查询数据/数据导出中"; }));
                Dispatcher.Invoke(new Action(delegate { bulletinboard.IsActive = true; }));
                return;
            }
            thread.Start();
        }
        public DataSet DataSet;
        public DataSet dataSet1;
        private void SnackbarMessage_ActionClick(object sender, RoutedEventArgs e)
        {
            bulletinboard.IsActive = false;
        }
        private void selectdata()
        {
            Dispatcher.Invoke(new Action(delegate { K = starttime.Text; }));
            Dispatcher.Invoke(new Action(delegate { m = endtime.Text; }));
            if (YieldRatLordTabel.Items.Count < 1)
            {
                bulletinboard.IsActive = true;
            }
            else
            {
                Dispatcher.Invoke(new Action(delegate { progressbar1.Visibility = Visibility.Visible; }));
                Dispatcher.Invoke(new Action(delegate { progressbar1.Maximum = DataSet.Tables[0].Rows.Count; }));
                for (int i = 0; i < DataSet.Tables[0].Rows.Count; i++)
                {
                    double s = Math.Round(i / (double)DataSet.Tables[0].Rows.Count, 2) * 100;
                    Dispatcher.Invoke(new Action(delegate { progressbar1.Value = i; }));
                    Dispatcher.Invoke(new Action(delegate { bfb.Text = s + "%"; }));
                    string stringBuilder = "'" + DataSet.Tables[0].Rows[i]["製程名稱"].ToString() + "'";
                    string stringBuilder1 = "'" + DataSet.Tables[0].Rows[i]["工單號碼"].ToString() + "'";
                    string stringBuilder2 = "'" + DataSet.Tables[0].Rows[i]["成品料號"].ToString() + "'";
                    string sql = @"Select A.WORK_ORDER ""工單號碼"",n.remark""工單備註"",E.Part_No ""成品料號"",A.SERIAL_NUMBER ""生產序號"",A.CUSTOMER_SN ""客戶序號"", SAJET.Sj_Snstatus_Result(A.CURRENT_STATUS)  ""狀態"", O.WEIGHT ""彩盒重量"",E.OPTION11 ""彩盒重量下限"",E.OPTION12 ""彩盒重量上限"", P.WEIGHT ""外箱重量"",E.OPTION9 ""外箱重量下限"",E.OPTION10 ""外箱重量上限"", C.PDLINE_NAME ""線別"",D.PROCESS_NAME ""製程名稱"", TO_CHAR(A.OUT_PROCESS_TIME, 'YYYY/MM/DD HH24:MI:SS') ""過站時間"", F.EMP_NAME ""作業人員"" ,decode(A.CURRENT_STATUS,'0','',decode(I.DEFECT_ID,'10000063',nvl(sajet.culpritdata(A.SERIAL_NUMBER), I.DEFECT_DESC),I.DEFECT_DESC)) ""不良現象描述"" ,decode(A.CURRENT_STATUS,'0','',K.REASON_DESC)  ""不良原因"",decode(A.CURRENT_STATUS,'0','',K.REASON_DESC2) ""不良原因描述2"",decode(A.CURRENT_STATUS,'0','',L.DUTY_DESC)  ""責任歸屬"",decode(A.CURRENT_STATUS,'0','',M.lOCATION)  ""不良位置"",decode(A.CURRENT_STATUS,'0','',M.ITEM_NO)  ""零件代碼"" From SAJET.G_SN_TRAVEL A,SAJET.SYS_PDLINE C,SAJET.SYS_PROCESS D,SAJET.SYS_EMP F,SAJET.SYS_PART E,SAJET.SYS_MODEL G,SAJET.G_SN_DEFECT H,SAJET.SYS_DEFECT I,SAJET.G_SN_REPAIR J,SAJET.SYS_REASON K,SAJET.SYS_DUTY L,SAJET.G_SN_REPAIR_LOCATION M,SAJET.G_WO_BASE N,SAJET.G_BOX_WEIGHT O,SAJET.G_CARTON_WEIGHT P Where  0=0 and D.PROCESS_NAME = " + stringBuilder.ToString() + " and A.Work_Order = " + stringBuilder1.ToString() + " and E.PART_NO = " + stringBuilder2.ToString() + " and  OUT_PROCESS_TIME >= TO_DATE('" + K + "','yyyy/mm/dd') and out_process_time <= to_date('" + m + "','yyyy/mm/dd') and A.PDLINE_ID = C.PDLINE_ID(+) and A.PROCESS_ID = D.PROCESS_ID(+) and A.PART_ID = E.PART_ID(+) and E.MODEL_ID = G.MODEL_ID(+) and A.EMP_ID = F.EMP_ID(+) and A.SERIAL_NUMBER=H.SERIAL_NUMBER(+) and A.PROCESS_ID=H.PROCESS_ID(+) and A.TERMINAL_ID=H.TERMINAL_ID(+) AND A.OUT_PROCESS_TIME=H.REC_TIME(+) and H.DEFECT_ID=I.DEFECT_ID(+) AND H.RECID=J.RECID(+) and J.REASON_ID = K.REASON_ID(+) and J.DUTY_ID=L.DUTY_ID(+) and a.work_order=n.work_order(+) AND H.RECID=M.RECID(+) and A.BOX_NO=O.BOX(+) and A.CARTON_NO=P.CARTON(+) ";
                    DataSet dataSet = Oraclecnn.OracleCnn(sql);
                    if (dataSet1 == null)
                    {
                        dataSet1 = dataSet.Copy();
                    }
                    else
                    {
                        dataSet1.Merge(dataSet);
                    }
                }
                Thread thread = new Thread(() => excelsave(DataSet, dataSet1));
                thread.Start();
            }
            Dispatcher.Invoke(new Action(delegate { progressbar1.Visibility = Visibility.Collapsed; }));
        }
        private void excelsave(DataSet dataSet, DataSet dataSet1)
        {
            string FilePath = "";
            int d = 1, p = 0;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel表格（*.xlsx）|*.xlsx";
            saveFileDialog.FilterIndex = 1;
            if ((bool)saveFileDialog.ShowDialog())
            {
                Dispatcher.Invoke(new Action(delegate { matec.Visibility = Visibility.Visible; }));
                FilePath = saveFileDialog.FileName.ToString();
                FileInfo file = new FileInfo(FilePath);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage())
                {
                    Dispatcher.Invoke(new Action(delegate { progressbar1.Maximum = dataSet1.Tables[0].Rows.Count; }));
                    var worksheet = package.Workbook.Worksheets.Add("Date");
                    for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataSet.Tables[0].Columns[i].ColumnName.ToString();
                    }
                    for (int i = 0; i < dataSet.Tables[0].Columns.Count; i++)
                    {
                        for (int j = 0; j < dataSet.Tables[0].Rows.Count; j++)
                        {
                            worksheet.Cells[j + 2, i + 1].Value = dataSet.Tables[0].Rows[j][i].ToString();
                        }
                    }
                    var worksheets = package.Workbook.Worksheets.Add("Dateil-" + d);
                    for (int i = 0; i < dataSet1.Tables[0].Columns.Count; i++)
                    {
                        worksheets.Cells[1, i + 1].Value = dataSet1.Tables[0].Columns[i].ColumnName.ToString();
                    }
                    for (int i = 0; i < dataSet1.Tables[0].Rows.Count; i++)
                    {
                        p++;
                        for (int j = 0; j < dataSet1.Tables[0].Columns.Count; j++)
                        {
                            if (i == 1048575 * d)
                            {
                                d++; p = 0;
                                worksheets = package.Workbook.Worksheets.Add("Dateil-" + d);
                                for (int k = 0; k < dataSet1.Tables[0].Columns.Count; k++)
                                {
                                    worksheets.Cells[1, k + 1].Value = dataSet1.Tables[0].Columns[k].ColumnName.ToString();
                                }
                                p++;
                            }
                            worksheets.Cells[p + 1, j + 1].Value = dataSet1.Tables[0].Rows[i][j].ToString();
                        }
                    }
                    package.SaveAs(file);
                }
                Dispatcher.Invoke(new Action(delegate { matec.Visibility = Visibility.Collapsed; }));
                Dispatcher.Invoke(new Action(delegate { snacktext.Content = "导出成功！"; }));
                Dispatcher.Invoke(new Action(delegate { bulletinboard.IsActive = true; }));
            }
        }

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            ProgressBar2.Visibility = Visibility.Visible;
            Thread thread = new Thread(Getdatagrid1);
            thread.Start();
        }
        private void Select_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            string sql = @"select c.work_order as ""工單"",f.PART_NO as ""料號"",e.model_name as ""機種"", a.serial_number as ""序號"",b.PROCESS_NAME as ""工作站"",retest_time as ""測試時間"",d.EMP_NAME as ""員工"" from G_SN_BADNESs a ,sys_process b,g_sn_status c,sys_emp d,sys_model e ,sys_part f where A.PROCESS_ID	= B.PROCESS_ID	and a.serial_number = c.serial_number and c.PART_ID = f.PART_ID	and f.model_id = e.MODEL_ID(+)	and a.retest_emp_id = d.emp_id(+)";
            DataSet dataSet = Oraclecnn.OracleCnn(sql);
            Recordsheet.ItemsSource = dataSet.Tables[0].DefaultView;
        }

        private void Derive_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Thread thread = new Thread(Excel_exoprt);
            thread.Start();
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            cpk.Visibility = Visibility.Collapsed;
            YieldRat.Visibility = Visibility.Collapsed;
            record.Visibility = Visibility.Visible;
            MultithreadigStart();
        }
        private void Excel_exoprt()
        {
            Dispatcher.Invoke(new Action(delegate { matec.Visibility = Visibility.Visible; }));
            DataView dataView = (DataView)Items1.ItemsSource;
            if (dataView.Table.Rows.Count == 0)
            {
                return;
            }
            DataTable dataTable = new DataTable("Title");
            dataTable.Columns.Add("FirstColum");
            string[] array = new string[]
            {
                "Test Item",
                "Vin(AC/Hz)",
                "Load",
                "Vout",
                "S/N",
                "MAX_SPEC",
                "MIN_SPEC",
                "UNIT",
                "MAX",
                "MIN",
                "AVG",
                "STD",
                "Cpu",
                "Cpl",
                "Cp > 1",
                "Ca < 1",
                "Cpk > 1",
                "Result",
                "PC_IDX"
            };
            foreach (string text in array)
            {
                dataTable.Rows.Add(new string[]
                {
                    text
                });
            }

            foreach (DataRow dataRow in (dataView.Table).Select(" ", "SERIAL_NUMBER, PC_INDEX, UPDATE_TIME"))
            {

                string text2 = "C" + dataRow["PC_INDEX"].ToString();
                if (dataTable.Columns.Contains(text2))
                {
                    continue;
                }
                dataTable.Columns.Add(text2);
                dataTable.Rows[Array.IndexOf<string>(array, "Test Item")][text2] = dataRow["TEST_ITEM"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "Vin(AC/Hz)")][text2] = dataRow["FREQUENCY"].ToString() + "/" + dataRow["VOLTAGE"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "Load")][text2] = dataRow["LOAD_2"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "Vout")][text2] = dataRow["VOUT"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "S/N")][text2] = dataRow["VOUT_ID"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "MAX_SPEC")][text2] = dataRow["USL"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "MIN_SPEC")][text2] = dataRow["LSL"].ToString();
                dataTable.Rows[Array.IndexOf<string>(array, "UNIT")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "MAX")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "MIN")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "AVG")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "STD")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "Cpu")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "Cpl")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "Cp > 1")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "Ca < 1")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "Cpk > 1")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "Result")][text2] = "";
                dataTable.Rows[Array.IndexOf<string>(array, "PC_IDX")][text2] = dataRow["PC_INDEX"].ToString();
            }
            DataTable dataTable2 = dataTable.Clone();
            DataTable dataTable3 = dataTable.Clone();
            string text3 = string.Empty;
            foreach (DataRow dataRow2 in (dataView.Table).Select("", "SERIAL_NUMBER, PC_INDEX, UPDATE_TIME"))
            {
                if (text3 != dataRow2["SERIAL_NUMBER"].ToString())
                {
                    text3 = dataRow2["SERIAL_NUMBER"].ToString();
                    dataTable2.Rows.Add(new string[]
                    {
                        text3
                    });
                }
                string text4 = "C" + dataRow2["PC_INDEX"].ToString();
                if (dataTable2.Columns.Contains(text4))
                {
                    dataTable2.Rows[dataTable2.Rows.Count - 1][text4] = dataRow2["SPC_VALUE"].ToString();
                }
            }
            for (int l = 0; l < 5; l++)
            {
                dataTable3.ImportRow(dataTable.Rows[l]);
            }
            foreach (object obj in dataTable2.Rows)
            {
                DataRow row = (DataRow)obj;
                dataTable3.ImportRow(row);
            }
            for (int m = 5; m < dataTable.Rows.Count; m++)
            {
                dataTable3.ImportRow(dataTable.Rows[m]);
            }
            ExcelUtil excelUtil = new ExcelUtil();
            excelUtil.CPKData(dataTable3, 6, dataTable2.Rows.Count);
            excelUtil.WriteToFile(ExportType.FileDialog);
            Dispatcher.Invoke(new Action(delegate { matec.Visibility = Visibility.Collapsed; }));
            Dispatcher.Invoke(new Action(delegate { snacktext.Content = "導出成功"; }));
            Dispatcher.Invoke(new Action(delegate { bulletinboard.IsActive = true; }));
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {

        }
        private void UpdateProcedure()
        {
            AutoUpdaterDotNET.AutoUpdater.Start("http://192.168.3.124/update.xml");
        }
        private void Updatedate()
        {
            do
            {
                Thread.Sleep(5000);
                dataSet = Oraclecnn.OracleCnn("select sni,snm from g_record_ate where to_char(sysdate,'hh24:mi:ss') between start_time and end_time");
                if (dataSet.Tables[0].Rows.Count != 0)
                {
                    cdata = Convert.ToInt32(dataSet.Tables[0].Rows[0]["SNI"]);
                    AteMax = Convert.ToInt32(dataSet.Tables[0].Rows[0]["SNM"]);
                    dataSetAts = Oraclecnn.OracleCnn("select sni,snm from g_record where to_char(sysdate,'hh24:mi:ss') between start_time and end_time");
                    ActualQuntity = Convert.ToInt32(dataSetAts.Tables[0].Rows[0]["SNI"]);
                    AtsMax = Convert.ToInt32(dataSetAts.Tables[0].Rows[0]["SNM"]);
                    badnessDataSet = Oraclecnn.OracleCnn(@"select c.work_order as ""工單"",f.PART_NO as ""料號"",e.model_name as ""機種"", a.serial_number as ""序號"",b.PROCESS_NAME as ""工作站"",retest_time as ""測試時間"",d.EMP_NAME as ""員工"" from G_SN_BADNESs a ,sys_process b,g_sn_status c,sys_emp d,sys_model e ,sys_part f where A.PROCESS_ID	= B.PROCESS_ID	and a.serial_number = c.serial_number and c.PART_ID = f.PART_ID	and f.model_id = e.MODEL_ID(+)	and a.retest_emp_id = d.emp_id(+) order by retest_time desc").Tables[0].DefaultView;

                    if (cdata >= AteMax || actualQuntity >= AtsMax)
                    {
                        Dispatcher.Invoke(new Action(delegate { content.Content = new AbnormalNotice(); }));
                        Dispatcher.Invoke(new Action(delegate { content.Visibility = Visibility.Visible; }));
                        Dispatcher.Invoke(new Action(delegate { this.WindowState = WindowState.Maximized; }));
                    }
                }


            } while (record.Visibility == Visibility.Visible);
        }
        private void MultithreadigStart()
        {

            testset = Oraclecnn.OracleCnn("select sum(testsum)test_sum, test_date,process_name from (select count(to_char(retest_time,'yyyy-mm-dd'))testsum ,to_char(retest_time,'yyyy-mm-dd') test_date,substr(b.PROCESS_NAME,1,3) process_name From g_sn_badness a,sys_process b where retest_time between to_date('" + BadTime[0] + "','yyyy-mm-dd')and to_date('" + BadTime[6] + "','yyyy-mm-dd') and a.process_id= B.PROCESS_ID group by to_char(retest_time,'yyyy-mm-dd'),b.process_name order by to_char(retest_time,'yyyy-mm-dd'),b.process_name) a group by test_date ,process_name order by test_date");
            var query = from p in testset.Tables[0].AsEnumerable()
                        where p.Field<string>("PROCESS_NAME") == "ATE"
                        select p;
            numberAte = new List<int>();
            timeAte = new List<string>();
            foreach (var item in query)
            {
                numberAte.Add(Convert.ToInt32(item[0]));
                timeAte.Add(item[1].ToString());
            }
            seriesCollection = new SeriesCollection
                 {
                    new LineSeries
                    {
                       Values = new ChartValues<int>(numberAte.ToArray())
                    }
                 };
            myaxisx.Labels = timeAte.ToArray();
            var quertats = from l in testset.Tables[0].AsEnumerable()
                           where l.Field<string>("PROCESS_NAME") == "ATS"
                           select l;
            numberAte.Clear();
            timeAte.Clear();
            foreach (var item in quertats)
            {
                numberAte.Add(Convert.ToInt32(item[0]));
                timeAte.Add(item[1].ToString());
            }
            seriesCollectionAts = new SeriesCollection
            {
                new LineSeries
                {
                    Values = new ChartValues<int>( numberAte.ToArray())
                }
            };
            myaxisxAts.Labels = timeAte.ToArray();
            Thread thread = new Thread(Updatedate);
            thread.Start();
        }
        private void AteResetdata()
        {
            Oraclecnn.OracleCnn("update g_record_ate set sni = '0' where to_char(sysdate,'hh24:mi:ss') between start_time and end_time");
            cdata = 0;
            bulletinboard.IsActive = true;
            InformContent = "重置成功";
        }
        private void AtsResetdata()
        {
            Oraclecnn.OracleCnn("update g_record set sni = '0' where to_char(sysdate,'hh24:mi:ss') between start_time and end_time");
            ActualQuntity = 0;
            bulletinboard.IsActive = true;
            InformContent = "重置成功";
        }

        private void Orderselect()
        {
            List<string> data = new List<string>();
            PodataView1 = T_sql.Sqldata("SELECT 訂單 from (SELECT concat(rtrim(ltrim(TD001)),'-',rtrim(ltrim(TD002)),'-',rtrim(ltrim(TD003))) 訂單 FROM  COPTD WHERE TD016 = 'N')a where 訂單 like '" + potextbox.Text + '-' + oddtextbox.Text + "%'").Tables[0].DefaultView;
            for (int i = 0; i < PodataView1.Table.Rows.Count; i++)
            {
                data.Add(PodataView1.Table.Rows[0][0].ToString());
            }
            orderdata = T_sql.Sqldata("select*from (SELECT CAST(A.库存 AS real)库存,A.列标签,cast(A.未结 as real)未结,cast(A.未进 as real)未进 ,CAST(A.需求 AS real)需求,CAST(A.已請未購 AS real)已請未購,CAST(A.預計出貨 AS real)預計出貨,CAST(A.預計用料 AS real)預計用料,CAST(A.总计 AS real)总计,CAST(A.库存+A.未进+A.已請未購+A.未结-A.預計用料-A.預計出貨-A.需求 AS real) 欠料  FROM (SELECT  B.MD003 列标签, SUM(MD006 * A.TD008)需求,SUM(MD006)总计,MAX(MB064)库存,ISNULL((select sum(TD008-TD015) from PURTD where TD004 = B.MD003 AND TD016 = 'N' AND TD018 = 'Y' group by TD004),0)未进,ISNULL((select SUM(TB009) from PURTB where TB004 = B.MD003 AND TB039 = 'N' GROUP BY TB004),0)已請未購,ISNULL((select SUM(TA015-TA017) froM MOCTA WHERE TA006 =B.MD003),0)未结,ISNULL((SELECT SUM(TB004-TB005)  FROM MOCTB Z,MOCTA C WHERE TB003 = B.MD003 AND Z.TB018 IN ('Y','N')AND Z.TB001 = C.TA001 AND Z.TB002 = C.TA002 AND C.TA011 = '1'),0)預計用料 ,ISNULL((SELECT SUM(TD008-TD009) FROM COPTD WHERE TD004 = B.MD003 AND TD016 = 'N' AND TD001 != '229'),0)預計出貨 FROM COPTD A,BOMMD B,INVMB C  WHERE A.TD001 ='270F'  AND A.TD004 = B.MD001 AND B.MD003 = C.MB001  GROUP BY B.MD003)A)a where a.欠料 < 0").Tables[0].DefaultView;
        }

        private void PackIcon_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            if (Beenstanding.Width == 0)
            {
                ss.Angle = 180;
                Beenstanding.Width = 1025;
            }
            else
            {
                ss.Angle = 1;
                Beenstanding.Width = 0;
            }
        }

        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            if (wo.Text.Length <= 0 && number.Text.Length <= 0)
            {
                Notifications("工單與序列不能同時為空");
            }
            else
            {
                List<string> Condition = new List<string>();
                if (wo.Text.Length <= 0)
                {
                    Condition.Add(" ");
                }
                else
                {
                    Condition.Add("a.work_order ='" + wo.Text.ToString() + "'and" + "");
                }
                if (number.Text.Length <= 0)
                {
                    Condition.Add(" ");
                }
                else
                {
                    string s = string.Empty;
                    var numbers = number.Text.ToString().Split(',', ' ', '，', '\r', '\n');
                    for (int i = 0; i < numbers.Length; i++)
                    {

                        string v = i >= 1 ? ",'" + numbers[i].ToString() + "'" : "'" + numbers[i].ToString() + "'";
                        s += v;
                    }
                    Condition.Add("a.SERIAL_NUMBER in (" + s + " ) and" + "");
                }
                Beenstandingdata.ItemsSource = Oraclecnn.OracleCnn(@"select A.SERIAL_NUMBER AS ""PSU 序號"" ,C.PROCESS_NAME AS ""製程"",B.SERIAL_NUMBER AS ""Sub Part - 1"",case when B.CURRENT_STATUS = 0 then 'Pass' else 'Fail'end AS ""狀態"",(select item_part_sn from G_SN_KEYPARTS  where SERIAL_NUMBER = A.ITEM_PART_SN )AS ""Sub Part - 2"",IN_PROCESS_TIME AS ""過站時間"" from G_SN_KEYPARTS a,G_SN_TRAVEL b,SYS_PROCESS C where " + Condition[0] + " " + Condition[1] + "  A.ITEM_PART_SN = B.SERIAL_NUMBER AND B.PROCESS_ID = C.PROCESS_ID order by A.SERIAL_NUMBER,C.PROCESS_NAME").Tables[0].DefaultView;
            }
        }

        private void SingleTrialClick()
        {
            cpk.Visibility = Visibility.Collapsed;
            record.Visibility = Visibility.Collapsed;
            YieldRat.Visibility = Visibility.Collapsed;
            SingleTrialDrid.Visibility = Visibility.Visible;
        }
        private void Notifications(string message)
        {
            Showadnormal();
            InformContent = message;
            bulletinboard.IsActive = true;
        }
        async Task Showadnormal()
        {
            await Task.Run(() => { Thread.Sleep(3000); });
            bulletinboard.IsActive = false;
        }
    }

}
