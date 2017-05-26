using System;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Xml;
using SWF = System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing.Printing;
using System.Text.RegularExpressions;
using System.Linq;
using System.Globalization;
using Microsoft.Win32;
using System.Xml.Linq;
using System.Data;

namespace Ecoview_V2._0
{
    public partial class EcoviewProfessional1 : Form
    {
        public EcoviewProfessional1()
        {
            InitializeComponent();
            // this.StartPosition = FormStartPosition.WindowsDefaultBounds;
            if (ComPodkl == false)
            {
                this.подключитьToolStripMenuItem.Enabled = true;
                button2.Enabled = true;

                button12.Enabled = false;
                button14.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = false;
                this.информацияToolStripMenuItem.Enabled = false;
                this.калибровкаToolStripMenuItem.Enabled = false;
                this.темновойТокToolStripMenuItem.Enabled = false;
                this.измеритьToolStripMenuItem.Enabled = false;

                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = false;
            }
            //   groupBox3.Enabled = false;
            // groupBox2.Enabled = false;
            tabPage4.Parent = null;
            radioButton1.Enabled = false;
            radioButton2.Enabled = false;
            radioButton3.Enabled = false;
            radioButton4.Enabled = false;
            radioButton5.Enabled = false;
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", 0);
            AgroText2.Text = string.Format("{0:0.0000}", 0);
        }
        public string version = "264";
        public bool nonPort;
        public SerialPort newPort;
        public SWF.TextBox[] textBox = new SWF.TextBox[20];
        public SWF.TextBox[] textBoxCO = new SWF.TextBox[20];
        public string WL_grad1;
        public int NoCaIzm;
        public int NoCaIzm1;
        public int NoCaSer;
        public int NoCaSer1;
        public int countWL = 3;
        public byte[] buffer4;
        public string GW1_2;
        public string GE5_1_0 = "";
        public byte[] GWbuffer;
        public int NoCoIzmer;
        public int NoCoSeria;
        public string edconctr;
        public int StolbecCol = 0;
        public int StolbecCol_1 = 0;
        public string SposobZadan;
        public int Days;
        public bool ComPodkl = false;
        public string DLWL;
        public string versionPribor;
        public bool USE_KO = false;
        public bool IzmerCreate = false;
        public bool IzmerCreate1 = false;
        bool OpenIzmer = false;
        bool OpenIzmer1 = false;
        public bool ComPort = false;
        bool Izmerenie1 = true;
        public string portsName = "";
        public int countSA;
        public string Veshestvo1;
        public string wavelength1 = Convert.ToString(0);
        public string WidthCuvette;
        public string BottomLine;
        public string TopLine;
        public string ND;
        public string Description;
        public string DateTime;
        public string Ispolnitel;
        public string CountSeriya;
        public string CountInSeriya;
        double SUMMX;
        double SUMMY;
        double XY;
        double SUMMY2;
        double SredX, SredY;
        public string aproksim;
        public string Zavisimoct;
        public int countOpenGrad = 0;
        public double k0;
        public double k1;
        public double k2;
        double SUM0, SUM1;
        int cellnull = 0;
        public string Veshestvo2;
        public string wavelength2;
        public string WidthCuvette2 = "";
        public string BottomLine2 = "";
        public string TopLine2 = "";
        public string ND2 = "";
        public string Description2 = "";
        public string DateTime2 = "";
        public string DateTime2_1 = "";
        public string DateTime2_2_1 = "";
        public string Ispolnitel2 = "";
        public string CountSeriya2 = Convert.ToString(3);
        public string CountInSeriya2 = Convert.ToString(3);
        public string[,] Stolbec;
        public string[,] Stolbec_1;
        public string Stolbec1 = "";
        public string Stroka1 = "";
        public string Pogreshnost2 = "";
        public string TypeYravn1 = "";
        public string TimeIzmer1 = "";
        public string USE_CO_XML1 = "";
        public string[] HeaderCells;
        public string[,] Cells1;
        public bool IzmerenieOpen = false;
        public string filepath2;
        public double[] El;
        int count = 0;
        public string filename;
        int cordY = 0;
        public double CellOpt;
        double k0_1 = 0;
        double k1_1 = 0;
        double k2_1 = 0;
        bool USE_KO_1 = false;
        public int selet_rezim = 2;
        bool StopAgro = false;
        public bool StopSpectr = false;
        public string[][,] countScan;
        public void Ecoview_Load(object sender, EventArgs e)
        {

            edconctr = "%";
            SposobZadan = "По СО";

            dateTimePicker1.Text = DateTime;

            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;

            chart1.Series[0].Points.Add();

            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();

            ScanChart.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            ScanChart.Series[0].Points.Add();
            ScanChart.ChartAreas[0].AxisX.Minimum = 400;
            ScanChart.ChartAreas[0].AxisX.Maximum = 600;
            if (ScanTable.Columns[1].HeaderText == "Abs")
            {
                //Array.Sort(massGE);
                ScanChart.ChartAreas[0].AxisY.Minimum = 0;
                ScanChart.ChartAreas[0].AxisY.Maximum = 3;
                ScanChart.ChartAreas[0].AxisX.Minimum = cancel;
                ScanChart.ChartAreas[0].AxisX.Maximum = start;
            }
            else
            {
                //Array.Sort(massGE);
                ScanChart.ChartAreas[0].AxisY.Minimum = 0;
                ScanChart.ChartAreas[0].AxisY.Maximum = 125;
                ScanChart.ChartAreas[0].AxisX.Minimum = cancel;
                ScanChart.ChartAreas[0].AxisX.Maximum = start;
            }
            ScanChart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            ScanChart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            ScanChart.ChartAreas[0].AxisX.Title = ScanTable.Columns["WalveDl"].HeaderText;
            ScanChart.ChartAreas[0].AxisY.Title = ScanTable.Columns["Abs_scan"].HeaderText;



            chart3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart3.Series[0].Points.Add();
            chart3.ChartAreas[0].AxisX.Minimum = 400;
            chart3.ChartAreas[0].AxisX.Maximum = 600;
            if (ScanTable.Columns[1].HeaderText == "Abs")
            {
                //Array.Sort(massGE);
                chart3.ChartAreas[0].AxisY.Minimum = 0;
                chart3.ChartAreas[0].AxisY.Maximum = 3;
                chart3.ChartAreas[0].AxisX.Minimum = cancel;
                chart3.ChartAreas[0].AxisX.Maximum = start;
            }
            else
            {
                //Array.Sort(massGE);
                chart3.ChartAreas[0].AxisY.Minimum = 0;
                chart3.ChartAreas[0].AxisY.Maximum = 125;
                chart3.ChartAreas[0].AxisX.Minimum = cancel;
                chart3.ChartAreas[0].AxisX.Maximum = start;
            }
            chart3.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisX.Title = TableKinetica1.Columns[0].HeaderText;
            chart3.ChartAreas[0].AxisY.Title = TableKinetica1.Columns[1].HeaderText;

            Zavisimoct = "A(C)";
            aproksim = "Линейная";
            Table1.AllowUserToResizeRows = false;
            Podskazka.Text = "Подключите прибор!";
            label27.Visible = false;
            label24.Visible = true;
            label25.Visible = false;
            label26.Visible = false;
            label28.Visible = false;



            ToolTip t1 = new ToolTip();
            t1.SetToolTip(Add_Table2, "Добавить образец");
            ToolTip t = new ToolTip();
            t.SetToolTip(Remove_Table2, "Удалить текущий образец");
            Select_modules();
        }
        public string Ecoview_Header = "";
        public void Select_modules()
        {
            //this.Hide();
            Select _Select = new Select(this);
            // _Select.Owner = this;
            _Select.ShowDialog();
            this.Text = Ecoview_Header;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            podcluchenie();
        }
        public void podcluchenie()
        {
            SettingPort _SettingPort = new SettingPort(this);
            if (nonPort == true)
            {
                _SettingPort.ShowDialog();
            }
            else
            {
                _SettingPort.Dispose();
            }
            //_SettingPort.Close();
            if (nonPort == true && Izmerenie1 != false)
            {
                newPort = new SerialPort();

                try
                {
                    // настройки порта (Communication interface)
                    newPort.PortName = portsName;
                    newPort.BaudRate = 19200;
                    newPort.DataBits = 8;
                    newPort.Parity = System.IO.Ports.Parity.None;
                    newPort.StopBits = System.IO.Ports.StopBits.One;
                    // Установка таймаутов чтения/записи (read/write timeouts)
                    newPort.ReadTimeout = 20000;
                    newPort.WriteTimeout = 20000;
                    //    newPort.DataReceived += new SerialDataReceivedEventHandler(newPort_DataReceived);
                    newPort.RtsEnable = false;
                    newPort.DtrEnable = true;
                    newPort.Open();// MessageBox.Show("ПОРТ ОТКРЫТ " + newPort.PortName);


                    newPort.DiscardInBuffer();
                    newPort.DiscardOutBuffer();
                }
                catch (Exception)
                {
                    MessageBox.Show("Порт не был выбран!");
                    return;

                }

                //char[] OpenPribor = { Convert.ToChar('C'), Convert.ToChar('O'), Convert.ToChar('\r') };
                //newPort.Write(OpenPribor, 0, OpenPribor.Length);

                File.WriteAllText(@"openport.port", string.Empty);
                File.AppendAllText(@"openport.port", portsName, Encoding.UTF8);

                Analis__Activated();
                CO();
                // GW();
                RD();
                //GW();
                //SA();
                Izmerenie1 = false;
                ComPodkl = true;

                SAGE(ref countSA, ref GE5_1_0);
                this.подключитьToolStripMenuItem.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = true;
                this.информацияToolStripMenuItem.Enabled = true;
                this.калибровкаToolStripMenuItem.Enabled = true;
                this.темновойТокToolStripMenuItem.Enabled = true;
                this.измеритьToolStripMenuItem.Enabled = true;
                this.измеритьToolStripMenuItem.Enabled = true;
                this.измеритьToolStripMenuItem.Enabled = true;
                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = false;

                button12.Enabled = true;
                ComPort = true;
                if ((OpenIzmer == true && ComPort == true) || (OpenIzmer1 == true && ComPort == true))
                {
                    button14.Enabled = true;
                }
                else
                {
                    button14.Enabled = false;
                }
                if (ComPort == true)
                {
                    button14.Enabled = true;
                }
                else
                {
                    button14.Enabled = false;
                }
                if (SposobZadan == "Ввод коэффициентов")
                {
                    button14.Enabled = false;
                }
                else
                {
                    button14.Enabled = true;
                }
                if (selet_rezim == 2 || selet_rezim == 6)
                {
                    Podskazka.Text = "Создайте или откройте Градуировку!";
                    label25.Visible = true;
                    label26.Visible = true;
                }
                else
                {
                    if (selet_rezim == 5)
                    {
                        Podskazka.Text = "Создайте Измерение";
                        label25.Visible = true;
                        label26.Visible = false;
                    }
                }
                label27.Visible = false;
                label24.Visible = false;

                label28.Visible = false;
                label33.Visible = false;

            }
        }

        public void ZapicInTable1()
        {
            bool doNotWrite = false;
            double sum = 0.0;
            int startIndexCell = 2;
            int endIndexCell = startIndexCell + NoCaIzm;
            int rowIndex = Table1.CurrentRow.Index;
            //int curentIndex = Table1.CurrentCell.ColumnIndex;

            if (Table1.CurrentCell.ColumnIndex > 2)
            {
                InputBox _InputBox = new InputBox(this);
                _InputBox.ShowDialog();
                if (Table1.CurrentCell.ReadOnly != true)
                {
                    Table1.CurrentCell.Value = string.Format("{0:0.0000}", CellOpt);
                    CellOpt = 0;



                }
                else
                {
                    MessageBox.Show("Запись запрещена!");
                }
            }


            int rownull = 0;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                {
                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;

                            for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                            {
                                if (Table1.Rows[rowIndex].Cells[l].Value == null)
                                {
                                    cellnull++;
                                }
                            }
                        }


                    }
                }
            }



            if (!doNotWrite)
            {
                if (NoCaSer == 1)
                {
                    radioButton1.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                    radioButton2.Enabled = false;
                }
                if (NoCaSer == 2)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                }
                if (NoCaSer >= 3)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                }

                sum = 0.0;
                /*while (true)
                 {
                     int i = Table1.Columns.Count - 1;//С какого столбца начать
                     if (Table1.Columns.Count == 3 + Convert.ToInt32(CountSeriya2))
                         break;
                     //Table1.Columns.RemoveAt(i);
                 }*/

                for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                {
                    if (Table1.Rows[rowIndex].Cells[l].Value == null)
                    {
                        cellnull++;
                    }

                    else
                    {
                        for (int j = 0; j < Table1.Rows.Count - 1; j++)
                        {

                            for (int i1 = startIndexCell + 1; i1 <= endIndexCell; ++i1)
                            {
                                sum += Convert.ToDouble(Table1.Rows[j].Cells[i1].Value);
                                Asred1 = sum / NoCaIzm;
                                // MessageBox.Show(Convert.ToString(Asred1));
                                Table1.Rows[j].Cells["Asred"].Value = string.Format("{0:0.0000}", Asred1);

                            }
                            sum = 0.0;
                        }
                    }
                    Izmerenie1 = true;
                }
                for (int m = 0; m < Table1.Rows.Count - 1; m++)
                {
                    for (int ml = 0; ml < Table1.Rows[m].Cells.Count; ml++)
                    {
                        if (Table1.Rows[m].Cells[ml].Value == null)
                        { doNotWrite = true; }
                    }
                }

                functionAsred();
            }

            int curentIndex = Table1.CurrentCell.ColumnIndex;
            if (curentIndex != Table1.ColumnCount - 1 || rowIndex != Table1.Rows.Count - 2)
            {
                if (rowIndex != Table1.Rows.Count - 2)
                {
                    Table1.CurrentCell = this.Table1[curentIndex, rowIndex + 1];
                }
                else
                {
                    Table1.CurrentCell = this.Table1[curentIndex + 1, 0];
                }
                Table1.EndEdit();
            }

        }

        public void LogoForm()
        {
            Form LogoForm = new Form();
            // LogoForm.BackColor = System.Drawing.Color.White;
            LogoForm.BackgroundImage = System.Drawing.Image.FromFile("Calibrovka.png");
            LogoForm.AutoScaleMode = AutoScaleMode.Font;
            LogoForm.Size = new Size(430, 107);
            LogoForm.Text = "Калибровка...";
            LogoForm.MinimizeBox = false;
            LogoForm.MaximizeBox = false;
            LogoForm.AutoSize = false;
            LogoForm.Name = "LogoFrm";
            LogoForm.ShowInTaskbar = false;
            LogoForm.StartPosition = FormStartPosition.CenterScreen;
            LogoForm.ControlBox = false;
            LogoForm.FormBorderStyle = FormBorderStyle.None;
            /*PictureBox PicBox = new PictureBox();
            PicBox.Size = new Size(307, 179);
            PicBox.Location = new System.Drawing.Point(12, 12);
            PicBox.ImageLocation = "D:\\Analis-samo\\Analis200\\Analis200\bin\x64\\Release\\Calibrovka.png";
            PicBox.SizeMode = PictureBoxSizeMode.Zoom;
            LogoForm.Controls.Add(PicBox);*/
            LogoForm.Show();
        }

        public void SAGE(ref int countSA, ref string GE5_1_0)
        {
            bool message1 = true;
            if (versionPribor.Contains("2"))
            { countSA = 8; }
            else
            {
                countSA = 4;
            }

            LogoForm();

            newPort.Write("SA " + countSA + "\r");

            string indata = newPort.ReadExisting();
            int indata_zero = 0;
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();

                }
            }

            newPort.Write("GE 1\r");
            // Thread.Sleep(500);
            // GEbyteRecieved4_1 = newPort.ReadBufferSize;
            //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
            // Thread.SpinWait(500);
            indata_0 = "";
            for (int i = 0; i <= 5000000; i++)
            {
                indata = newPort.ReadExisting();
                if (indata_0.Contains("\r>"))
                {
                    break;
                }
                indata_0 += indata;
            }
            indata_zero = 0;
            //  indata_0 = "";
            indata_bool = true;
            /* while (indata_bool == true)
             {

                 if (indata.Contains(">"))
                 {
                     indata_0 = indata;
                     indata_bool = false;

                 }
                 else {                   

                         indata = newPort.ReadExisting();
                         indata_0 += indata;

                 }
             }*/
            string GE5_1 = "";
            Regex regex = new Regex(@"\W");
            Regex regex1 = new Regex(@"\D");
            GE5_1 = regex.Replace(indata_0, "");
            GE5_1 = regex1.Replace(GE5_1, "");

            GE5_1_0 = regex.Replace(indata_0, "");
            GE5_1_0 = regex1.Replace(GE5_1, "");
            GEText.Text = GE5_1_0;
            //if(GE5_1 == "")
            {
                double GAText1 = (Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1_0)) * 100;

                GAText.Text = string.Format("{0:0.00}", GAText1);

                double OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1));

                double OptPlot1 = OptPlot - Math.Truncate(OptPlot);
                OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
                while (Convert.ToInt32(GE5_1) > 10000 && countSA > 1)
                {
                    countSA--;
                    newPort.Write("SA " + countSA + "\r");
                    int SAAnalisByteRecieved1_1_1 = newPort.ReadBufferSize;
                    // Thread.Sleep(100);
                    indata = newPort.ReadExisting();
                    indata_zero = 0;
                    indata_0 = "";
                    indata_bool = true;
                    while (indata_bool == true)
                    {

                        if (indata.Contains(">"))
                        {

                            indata_bool = false;

                        }

                        else {
                            indata = newPort.ReadExisting();
                        }
                    }

                    newPort.Write("GE 1\r");
                    // Thread.Sleep(500);
                    // GEbyteRecieved4_1 = newPort.ReadBufferSize;
                    //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                    // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
                    // Thread.SpinWait(500);
                    indata_0 = "";
                    for (int i = 0; i <= 5000000; i++)
                    {
                        indata = newPort.ReadExisting();
                        if (indata_0.Contains("\r>"))
                        {
                            break;
                        }
                        indata_0 += indata;
                    }
                    indata_zero = 0;
                    //  indata_0 = "";
                    indata_bool = true;
                    /* while (indata_bool == true)
                     {

                         if (indata.Contains(">"))
                         {
                             indata_0 = indata;
                             indata_bool = false;

                         }
                         else {                   

                                 indata = newPort.ReadExisting();
                                 indata_0 += indata;

                         }
                     }*/
                    regex = new Regex(@"\W");
                    regex1 = new Regex(@"\D");
                    GE5_1 = regex.Replace(indata_0, "");
                    GE5_1 = regex1.Replace(GE5_1, "");

                    GE5_1_0 = regex.Replace(indata_0, "");
                    GE5_1_0 = regex1.Replace(GE5_1, "");
                    GEText.Text = GE5_1_0;
                    //   if (GE5_1 == "")
                    {
                        GAText1 = (Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1_0)) * 100;


                        GAText.Text = string.Format("{0:0.00}", GAText1);

                        OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1));

                        OptPlot1 = OptPlot - Math.Truncate(OptPlot);
                        OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
                    }
                    /* else
                     {
                         MessageBox.Show("Внимание! При обнулении возникла ошибка! Попробуйте еще раз!");
                      //   SAGE(ref countSA, ref GE5_1_0);
                         return;
                     }*/
                }

            }
            /* else
             {
                 MessageBox.Show("Внимание! При обнулении возникла ошибка! Попробуйте еще раз!");
                // SAGE(ref countSA, ref GE5_1_0);
                 return;

             }*/
            SWF.Application.OpenForms["LogoFrm"].Close();

            if (Izmerenie1 == false)
            {
                // InitializeTimer();
            }
        }
        public double[] scan_mass;
        public void SAGEScan(ref int countSA, ref string GE5_1_0)
        {
            Regex regex1;
            bool message1 = true;
            if (versionPribor.Contains("2"))
            { countSA = 8; }
            else
            {
                countSA = 4;
            }

            LogoForm();
            string GE5_1 = "";
            string indata = newPort.ReadExisting();
            int indata_zero = 0;
            string indata_0 = "";
            bool indata_bool = true;
            int GEbyteRecieved4_1 = newPort.ReadBufferSize;
            byte[] GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            Regex regex = new Regex(@"\W");

            newPort.Write("SA " + countSA + "\r");

            indata = newPort.ReadExisting();
            indata_zero = 0;
            indata_0 = "";
            indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();

                }
            }
            //  Thread.Sleep(500);
            newPort.Write("GE 1\r");
            // Thread.Sleep(500);
            // GEbyteRecieved4_1 = newPort.ReadBufferSize;
            //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
            // Thread.SpinWait(500);
            indata_0 = "";
            for (int i = 0; i <= 5000000; i++)
            {
                indata = newPort.ReadExisting();
                if (indata_0.Contains("\r>"))
                {
                    break;
                }
                indata_0 += indata;
            }
            indata_zero = 0;
            //  indata_0 = "";
            indata_bool = true;
            /* while (indata_bool == true)
             {

                 if (indata.Contains(">"))
                 {
                     indata_0 = indata;
                     indata_bool = false;

                 }
                 else {                   

                         indata = newPort.ReadExisting();
                         indata_0 += indata;

                 }
             }*/
            regex = new Regex(@"\W");
            regex1 = new Regex(@"\D");
            GE5_1 = regex.Replace(indata_0, "");
            GE5_1 = regex1.Replace(GE5_1, "");

            GE5_1_0 = regex.Replace(indata_0, "");
            GE5_1_0 = regex1.Replace(GE5_1, "");

            GEText.Text = GE5_1_0;
            double GAText1 = (Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1_0)) * 100;

            GAText.Text = string.Format("{0:0.00}", GAText1);

            double OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1));

            double OptPlot1 = OptPlot - Math.Truncate(OptPlot);
            OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
            while (Convert.ToInt32(GE5_1) > 10000 && countSA > 1)
            {
                countSA--;
                newPort.Write("SA " + countSA + "\r");
                int SAAnalisByteRecieved1_1_1 = newPort.ReadBufferSize;
                // Thread.Sleep(100);
                indata = newPort.ReadExisting();
                indata_zero = 0;
                indata_0 = "";
                indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }
                //     Thread.Sleep(500);
                newPort.Write("GE 1\r");
                // Thread.Sleep(500);
                /*  GEbyteRecieved4_1 = newPort.ReadBufferSize;
                  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                  newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1);*/
                //    Thread.SpinWait(500);
                indata_0 = "";
                for (int i = 0; i <= 5000000; i++)
                {
                    indata = newPort.ReadExisting();
                    if (indata_0.Contains("\r>"))
                    {
                        break;
                    }
                    indata_0 += indata;
                }
                indata_zero = 0;
                //  indata_0 = "";
                indata_bool = true;
                /* while (indata_bool == true)
                 {

                     if (indata.Contains(">"))
                     {
                         indata_0 = indata;
                         indata_bool = false;

                     }
                     else {

                         indata = newPort.ReadExisting();
                         indata_0 += indata;

                     }
                 }*/

                regex = new Regex(@"\W");
                regex1 = new Regex(@"\D");
                GE5_1 = regex.Replace(indata_0, "");
                GE5_1 = regex1.Replace(GE5_1, "");

                GE5_1_0 = regex.Replace(indata_0, "");
                GE5_1_0 = regex1.Replace(GE5_1, "");

                GEText.Text = GE5_1_0;

                GAText1 = (Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1_0)) * 100;

                GAText.Text = string.Format("{0:0.00}", GAText1);

                OptPlot = Math.Log10(Convert.ToDouble(GE5_1_0) / Convert.ToDouble(GE5_1));

                OptPlot1 = OptPlot - Math.Truncate(OptPlot);
                OptichPlot.Text = string.Format("{0:0.0000}", OptPlot1);
            }
            SWF.Application.OpenForms["LogoFrm"].Close();
            //  listBox1.Items.Add(GE5_1_0);
            scan_mass[countscan] = Convert.ToDouble(GE5_1_0);
            scan_massSA[countscan] = Convert.ToDouble(countSA);
        }
        double[] scan_massSA;
        public void CO()
        {

            string b = "";
            int byteRecieved = newPort.ReadBufferSize;
            Thread.Sleep(500);
            byte[] buffer = new byte[byteRecieved];
            newPort.Read(buffer, 0, byteRecieved);

            string GW1 = "";

            for (int i = 0; i <= 50; i++)
            {
                GW1 = GW1 + Convert.ToChar(buffer[i]);
            }
            var GWarr = GW1.Split("\r".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);



            GW1_2 = GWarr[2];
            GWNew.Text = GW1_2;
            versionPribor = GWarr[1];
            if (wavelength1 == Convert.ToString(0) || wavelength1 == "")
            {
                wavelength1 = GW1_2;
            }
            else
            {
                bool dlinavoln = true;

                if (versionPribor.Contains("V"))
                {
                    if (Convert.ToDouble(wavelength1.Replace(".", ",")) < 315)
                    {
                        MessageBox.Show("Установленая длина волны выходит за пределы диапазона спектрофотометра, измените настройки градуировки!");
                        dlinavoln = false;
                    }
                    if (Convert.ToDouble(wavelength1.Replace(".", ",")) > 1050)
                    {
                        MessageBox.Show("Установленая длина волны выходит за пределы диапазона спектрофотометра, измените настройки градуировки!");
                        dlinavoln = false;
                    }
                }
                else
                {
                    if (versionPribor.Contains("U") && versionPribor.Contains("2"))
                    {
                        if (Convert.ToDouble(wavelength1.Replace(".", ",")) < 190)
                        {
                            MessageBox.Show("Установленая длина волны выходит за пределы диапазона спектрофотометра, измените настройки градуировки!");
                            dlinavoln = false;
                        }
                        if (Convert.ToDouble(wavelength1.Replace(".", ",")) > 1050)
                        {
                            MessageBox.Show("Установленая длина волны выходит за пределы диапазона спектрофотометра, измените настройки градуировки!");
                            dlinavoln = false;
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(wavelength1.Replace(".", ",")) < 200)
                        {
                            MessageBox.Show("Установленая длина волны выходит за пределы диапазона спектрофотометра, измените настройки градуировки!");
                            dlinavoln = false;
                        }
                        if (Convert.ToDouble(wavelength1.Replace(".", ",")) > 1050)
                        {
                            MessageBox.Show("Установленая длина волны выходит за пределы диапазона спектрофотометра, измените настройки градуировки!");
                            dlinavoln = false;
                        }
                    }
                }

                if (dlinavoln == true)
                {
                    SW();
                }
            }
        }
        public void SW()
        {
            double wevelenght1_double = Convert.ToDouble(wavelength1.Replace(".", ","));

            LogoForm2();
            newPort.Write("SW " + wevelenght1_double.ToString(System.Globalization.CultureInfo.GetCultureInfo("en-US")) + "\r");


            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }


            SWF.Application.OpenForms["LogoFrm2"].Close();
            GWNew.Text = string.Format("{0:0.00}", wavelength1);

        }
        public void LogoForm2()
        {
            Form LogoForm2 = new Form();
            // LogoForm.BackColor = System.Drawing.Color.White;
            LogoForm2.BackgroundImage = System.Drawing.Image.FromFile("Yasnovka_DLWALVE.png");
            LogoForm2.AutoScaleMode = AutoScaleMode.Font;
            LogoForm2.Size = new Size(430, 107);
            LogoForm2.Text = "Установка длины волны...";
            LogoForm2.MinimizeBox = false;
            LogoForm2.MaximizeBox = false;
            LogoForm2.AutoSize = false;
            LogoForm2.Name = "LogoFrm2";
            LogoForm2.ShowInTaskbar = false;
            LogoForm2.StartPosition = FormStartPosition.CenterScreen;
            LogoForm2.ControlBox = false;
            LogoForm2.FormBorderStyle = FormBorderStyle.None;

            LogoForm2.Show();
        }
        string[] RDstring;
        public void RD()
        {
            newPort.Write("RD\r");

            Thread.Sleep(500);
            //  byte[] buffer1 = new byte[byteRecieved1];
            string indata = newPort.ReadExisting();
            string indata_0 = "";
            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {

                    indata = newPort.ReadExisting();

                }
            }

            string substring = "\r";
            int count = (indata.Length - indata.Replace(substring, "").Length) / substring.Length;
            RDstring = new string[count];
            // Regex regex = new Regex(@"\W");
            for (int i = 0; i < count; i++)
            {
                RDstring[i] = indata.Split('\r')[i]; ;
            }

        }
        public void Analis__Activated()
        {
            newPort.Write("CO\r");

        }
        private void подключитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            podcluchenie();
        }

        private void Ecoview_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ComPort == true)
            {
                char[] ClosePribor = { Convert.ToChar('Q'), Convert.ToChar('U'), Convert.ToChar('\r') };
                newPort.Write("QU\r");
                Thread.Sleep(500);
                //  byte[] buffer1 = new byte[byteRecieved1];
                string indata = newPort.ReadExisting();

                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                newPort.Close();
                wavelength1 = Convert.ToString(0);
            }
            else
            {
                SWF.Application.Exit();
                ///  this.Hide();

                /* Select _Select = new Select(this);                
                 _Select.Owner = this;
                 _Select.ShowDialog();*/
                //System.Environment.Exit(0);
                //   Dispose();
                //   Close();

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PortClose();
        }
        public void PortClose()
        {
            ComPort = false;
            ComPodkl = false;
            if (ComPort == false)
            {
                char[] ClosePribor = { Convert.ToChar('Q'), Convert.ToChar('U'), Convert.ToChar('\r') };
                newPort.Write("QU\r");
                Thread.Sleep(500);
                //  byte[] buffer1 = new byte[byteRecieved1];
                string indata = newPort.ReadExisting();
                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else {
                        indata = newPort.ReadExisting();
                    }
                }

                GWNew.Text = null;
                GEText.Text = null;
                GAText.Text = null;
                OptichPlot.Text = null;
                Izmerenie1 = true;

                this.подключитьToolStripMenuItem.Enabled = true;
                button2.Enabled = true;

                button12.Enabled = false;
                button14.Enabled = false;
                this.настройкаПортаToolStripMenuItem.Enabled = false;
                this.информацияToolStripMenuItem.Enabled = false;
                this.калибровкаToolStripMenuItem.Enabled = false;
                this.темновойТокToolStripMenuItem.Enabled = false;
                this.измеритьToolStripMenuItem.Enabled = false;

                this.калибровкаДляОдноволновогоАнализаToolStripMenuItem.Enabled = false;
                button1.Enabled = false;

                newPort.Close();
                wavelength1 = Convert.ToString(0);
                // ComPort = false;

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Information();
        }
        public void Information()
        {
            PriborInformation _PriborInformation = new PriborInformation(this);
            _PriborInformation.ShowDialog();
        }

        private void информацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Information();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PortClose();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            TopMost = false;
            if (selet_rezim == 2)
            {
                Izmerenie1 = true;
                if (tabControl2.SelectedIndex == 0)
                {
                    NewGraduirovca(ref CountInSeriya, ref CountSeriya);
                }
                else
                {
                    NewIzmerenie();
                }
            }
            else
            {
                if (selet_rezim == 1)
                {
                    if (ComPodkl == true)
                    {
                        IzmerenieFR_new();
                    }
                    else
                    {
                        MessageBox.Show("Для проведения измерений необхдимо подключиться к прибору!");
                    }
                }
                else
                {
                    if (selet_rezim == 6)
                    {
                        Izmerenie1 = true;
                        if (tabControl2.SelectedIndex == 0)
                        {
                            NewGraduirovca(ref CountInSeriya, ref CountSeriya);
                        }
                        else
                        {
                            NewIzmerenie();
                        }
                    }
                    else
                    {
                        if (selet_rezim == 5)
                        {
                            if (ComPodkl == true)
                            {

                                FotometrScan();
                                Podskazka.Text = "Измерьте 0 Asb/100%T";
                                label25.Visible = false;
                                label59.Visible = true;
                            }
                            else
                            {
                                MessageBox.Show("Для проведения сканирования необхдимо подключиться к прибору!");
                            }
                        }
                        else
                        {
                            if (selet_rezim == 4)
                            {
                                if (ComPodkl == true)
                                {
                                    KineticaScan();
                                }
                                else
                                {
                                    MessageBox.Show("Для проведения сканирования необхдимо подключиться к прибору!");
                                }
                            }
                            else
                            {
                                if (selet_rezim == 3)
                                {
                                    if (ComPodkl == true)
                                    {
                                        MultiWave();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Для проведения сканирования необхдимо подключиться к прибору!");
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
        public void MultiWave()
        {
            MultiWave _MultiWave = new MultiWave(this);
            _MultiWave.ShowDialog();
        }
        public double interval;
        public double delay;
        public void KineticaScan()
        {
            Kinetica _KineticaScan = new Kinetica(this);
            _KineticaScan.ShowDialog();
        }
        public double start = 0.0;
        public double cancel = 0.0;

        public void FotometrScan()
        {
            FotometrScan _FotometrScan = new FotometrScan(this);
            _FotometrScan.ShowDialog();
        }
        public double[] massWL;
        public double[] massGE;
        delegate void GetMessage(); // 1. Объявляем делегат
        private void GoodMorning()
        {
            TableKinetica1.Rows.Add();
        }
        int timeLeft;
        public void TableKinetica(object sender, EventArgs e)
        {
            //  int x = (int)obj;
            // GetMessage del; // 2. Создаем переменную делегата
            //  del = GoodMorning;

            // timeLeft = timeLeft - Convert.ToInt32(interval * 1000);
            // label56.Text = Convert.ToString(timeLeft);
            timeLeft = timeLeft - Convert.ToInt32(interval);
            label56.Text = Convert.ToString(timeLeft);
            Application.DoEvents();
            //  MessageBox.Show("Интервал: " + interval*1000);
            TableKinetica1.Rows.Add();
            Array.Resize<double>(ref massWL, massWL.Length + 1);
            Array.Resize<double>(ref massGE, massGE.Length + 1);

            string GE5Izmer = "";
            string GE5_1_1 = "";
            while (GE5Izmer == "")
            {
                // SW_Scan();
                GE5Izmer = "";
                GE5_1_1 = "";
                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();
                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {
                        indata_bool = false;
                    }

                    else
                    {
                        indata = newPort.ReadExisting();
                    }
                }
                newPort.Write("GE 1\r");
                // Thread.Sleep(500);
                // GEbyteRecieved4_1 = newPort.ReadBufferSize;
                //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
                // Thread.SpinWait(500);
                indata_0 = "";
                for (int i = 0; i <= 5000000; i++)
                {
                    indata = newPort.ReadExisting();
                    if (indata_0.Contains("\r>"))
                    {
                        break;
                    }
                    indata_0 += indata;
                }

                //  indata_0 = "";
                indata_bool = true;
                /* while (indata_bool == true)
                 {

                     if (indata.Contains(">"))
                     {
                         indata_0 = indata;
                         indata_bool = false;

                     }
                     else {                   

                             indata = newPort.ReadExisting();
                             indata_0 += indata;

                     }
                 }*/
                Regex regex = new Regex(@"\W");
                Regex regex1 = new Regex(@"\D");
                GE5Izmer = regex.Replace(indata_0, "");
                GE5Izmer = regex1.Replace(GE5Izmer, "");
            }
            //  MessageBox.Show("Измерение: " + GE5Izmer + "\rКалибровка: " + GE5_1_0 + "\rОтклонение: " + Convert.ToString(Convert.ToDouble(GE5_1_0) / (Convert.ToDouble(GE5Izmer))) +
            //    "\rПоправочный коэффициент: " + RDstring[countSA]);
            GEText.Text = GE5Izmer;
            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;
            //MessageBox.Show(RDstring[countSA]);
            OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
            double OptPlot1_1 = OptPlot1;
            Application.DoEvents();
            massWL[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            TableKinetica1.Rows[countscan].Cells[0].Value = Convert.ToInt32(interval) * countscan;
            if (TableKinetica1.Columns[1].HeaderText == "Abs")
            {
                TableKinetica1.Rows[countscan].Cells[1].Value = string.Format("{0:0.0000}", OptPlot1_1);
                massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
                TableKinetica1.Rows[countscan].Cells[2].Value =
                    string.Format("{0:0.0}",
                    ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) * 100));
            }
            else
            {
                TableKinetica1.Rows[countscan].Cells[2].Value = string.Format("{0:0.0000}", OptPlot1_1);
                TableKinetica1.Rows[countscan].Cells[1].Value =
                    string.Format("{0:0.0}",
                    ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) * 100));
                /** listBox1.Items.Add(GE5Izmer);*/
                massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);

            }
            massWL[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
            Array.Sort(massWL);
            Array.Sort(massGE);
            double x1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            double y1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
            /*ScanChart.Series[0].Points.AddXY(x1, y1);
            ScanChart.Series[0].ChartType = SeriesChartType.Point;
            ScanChart.ChartAreas[0].AxisY.Crossing = 0;
            ScanChart.ChartAreas[0].AxisX.Crossing = 0;*/

            chart3.Series[countButtonClick].Points.AddXY(x1, y1);
            chart3.Series[countButtonClick].ChartType = SeriesChartType.Line;
            if (TableKinetica1.Rows[countscan].Cells[1].Value != null && TableKinetica1.Rows[countscan].Cells[2].Value != null)
            {
                chart3.ChartAreas[0].AxisX.Minimum = cancel;
                chart3.ChartAreas[0].AxisX.Maximum = start;

            }
            chart3.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisX.Title = TableKinetica1.Columns[0].HeaderText;
            chart3.ChartAreas[0].AxisY.Title = TableKinetica1.Columns[1].HeaderText;
            countscan++;
            //del = GoodMorning;


            if (timeLeft < 0)
            {
                timer2.Enabled = false;
                timer2.Stop();
                MinMax();
                button14.Enabled = true;
                button11.Enabled = false;
            }

        }
        public void MinMax()
        {
            double max = 0.0;
            double min = 0.0;
            countscan = 0;
            for (int i = 0; i < TableKinetica1.Rows.Count; i++)
            {
                double x1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
                double y1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
                if (i == 0)
                {
                    if (Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) > Convert.ToDouble(TableKinetica1.Rows[i + 1].Cells[1].Value))
                    {
                        max = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        dataGridView3.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                        min = max;
                        x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                        y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        chart3.Series[0].Points.AddXY(x1, y1);
                        chart3.Series[0].ChartType = SeriesChartType.Point;
                    }
                    else
                    {
                        min = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        dataGridView4.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                        max = min;
                        x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                        y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        chart3.Series[0].Points.AddXY(x1, y1);
                        chart3.Series[0].ChartType = SeriesChartType.Point;
                    }

                }
                else {
                    if (i + 1 != TableKinetica1.Rows.Count)
                    {
                        if (Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) > Convert.ToDouble(TableKinetica1.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) >= Convert.ToDouble(TableKinetica1.Rows[i + 1].Cells[1].Value)

                            )
                        {
                            max = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            min = max;
                            dataGridView3.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                            x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                            y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            chart3.Series[0].Points.AddXY(x1, y1);
                            chart3.Series[0].ChartType = SeriesChartType.Point;
                        }
                        if ((Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) < Convert.ToDouble(TableKinetica1.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) <= Convert.ToDouble(TableKinetica1.Rows[i + 1].Cells[1].Value))

                            )
                        {
                            min = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            dataGridView4.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                            max = min;
                            x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                            y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            chart3.Series[0].Points.AddXY(x1, y1);
                            chart3.Series[0].ChartType = SeriesChartType.Point;
                        }
                    }
                }
            }

            // countscan++;
        }
        public void IzmerenieFR_new()
        {
            IzmerenieFR _IzmerenieFR = new IzmerenieFR(this);
            _IzmerenieFR.ShowDialog();
        }
        public string CountSeriya1 = "";
        public string CountInSeriya1 = "";
        public void NewGraduirovca(ref string CountInSeriya, ref string CountSeriya)
        {

            CountSeriya1 = CountSeriya;
            CountInSeriya1 = CountInSeriya;
            ParametrsGrad _ParametrsGrad = new ParametrsGrad(this);
            _ParametrsGrad.ShowDialog();
            _ParametrsGrad.button1.Click += (ParametrsGrad, eSlave) =>
            {
                Veshestvo1 = _ParametrsGrad.Veshestvo.Text;
                wavelength1 = string.Format("{0:0.00}", _ParametrsGrad.WL_grad.Text);
                WidthCuvette = _ParametrsGrad.Opt_dlin_cuvet.Text;
                BottomLine = _ParametrsGrad.Down.Text;
                TopLine = _ParametrsGrad.Up.Text;
                ND = _ParametrsGrad.ND.Text;
                Description = _ParametrsGrad.Description.Text;
                DateTime = _ParametrsGrad.dateTimePicker1.Text;
                Ispolnitel = _ParametrsGrad.Ispolnitel.Text;
                CountSeriya1 = _ParametrsGrad.numericUpDown3.Text;
                CountInSeriya1 = _ParametrsGrad.numericUpDown4.Text;
                textBox3.Text = _ParametrsGrad.textBox4.Text;
                Days = Convert.ToInt32(_ParametrsGrad.numericUpDown1.Value);
                label6.Text = dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy");
                edconctr = _ParametrsGrad.Ed.Text;

                if (_ParametrsGrad.radioButton4.Checked == true)
                {
                    Zavisimoct = "A(C)";
                    radioButton4.Checked = true;
                    label14.Text = "A(C)";
                    /*                    textBox4.Text = "";
                                        textBox5.Text = "";
                                        textBox6.Text = "";*/
                }
                else
                {
                    Zavisimoct = "C(A)";
                    radioButton5.Checked = true;
                    label14.Text = "C(A)";
                    /*                   textBox4.Text = "";
                                       textBox5.Text = "";
                                       textBox6.Text = "";*/

                }
                if (_ParametrsGrad.radioButton1.Checked == true)
                {
                    aproksim = "Линейная через 0";
                }
                else
                {
                    if (_ParametrsGrad.radioButton2.Checked == true)
                    {
                        aproksim = "Линейная";
                    }
                    else
                    {
                        aproksim = "Квадратичная";
                    }
                }
                if (_ParametrsGrad.radioButton6.Checked == true)
                {
                    SposobZadan = "По СО";
                }
                else
                {
                    SposobZadan = "Ввод коэффициентов";

                    AgroText0.Text = _ParametrsGrad.k0Text.Text;
                    AgroText1.Text = _ParametrsGrad.k1Text.Text;
                    AgroText2.Text = _ParametrsGrad.k2Text.Text;
                    k0 = Convert.ToDouble(AgroText0.Text);
                    k1 = Convert.ToDouble(AgroText1.Text);
                    k2 = Convert.ToDouble(AgroText2.Text);


                }
                if (_ParametrsGrad.USE_KO.Checked == true)
                {
                    USE_KO = true;


                }
                else
                {
                    USE_KO = false;
                }
                Podskazka.Text = "Измерьте СО!";
                label27.Visible = false;
                label24.Visible = false;
                label25.Visible = false;
                label26.Visible = false;
                label28.Visible = true;
                label33.Visible = true;
            };

            CountSeriya = CountSeriya1;
            CountInSeriya = CountInSeriya1;
            dateTimePicker1.Text = DateTime;
            tabControl2.SelectTab(tabPage3);
            /*chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();*/
            while (true)
            {
                int i = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns.Count == 3 + NoCaIzm)
                    break;
                Table1.Columns.RemoveAt(i);
            }
            Table1.Rows[Table1.Rows.Count - 1].Cells[0].Value = "";
            if (Convert.ToInt32(CountInSeriya) < 3)
            {
                radioButton3.Enabled = false;
            }
            else
            {
                if (Convert.ToInt32(CountInSeriya) < 2)
                {
                    radioButton2.Enabled = false;
                    radioButton3.Enabled = false;
                }
                else
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                }
            }
            /*           RR.Text = "";
                       SKO.Text = "";
                       label21.Text = "";
                       label22.Text = "";*/
            if (SposobZadan == "По СО")
            {
                /*               textBox4.Text = "";
                               textBox5.Text = "";
                               textBox6.Text = "";*/
            }
            if (SposobZadan == "Ввод коэффициентов")
            {
                functionA();
            }

        }
        public void NewIzmerenie()
        {

            New _New = new New(this);
            _New.ShowDialog();
            if (ComPort == true)
            {
                button14.Enabled = true;
            }
            else
            {
                button14.Enabled = false;
            }

        }
        private void новыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Izmerenie1 = true;
            if (tabControl2.SelectedIndex == 0)
            {
                NewGraduirovca(ref CountInSeriya, ref CountSeriya);
            }
            else
            {
                NewIzmerenie();
            }
        }
        public void WLREMOVE1()
        {
            while (true)
            {
                int i = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns[i].Name == "Asred")
                    break;
                Table1.Columns.RemoveAt(i);
            }

        }
        public void WLREMOVEAgro1()
        {
            while (true)
            {
                int i = AgroTable1.Columns.Count - 1;//С какого столбца начать
                if (AgroTable1.Columns[i].Name == "AgroASred")
                    break;
                AgroTable1.Columns.RemoveAt(i);
            }

        }
        public void WLADD1()
        {

            for (int i = 1; i <= NoCaIzm; i++)
            {

                DataGridViewTextBoxColumn firstColumn1 =
                new DataGridViewTextBoxColumn();
                firstColumn1.HeaderText = "A; Сер" + i;
                firstColumn1.Name = "A;Ser (" + i;
                firstColumn1.ValueType = Type.GetType("System.Double");

                Table1.Columns.Add(firstColumn1);
                //firstColumn1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
                //   firstColumn1.EditingControlShowing

            }

            for (int i = 1; i <= NoCaIzm; i++)
            {
                Table1.Columns["A;Ser (" + i].Width = 50;
            }
            Concetr.HeaderText = "Конц " + edconctr;
        }
        public void WLADDAgro1()
        {

            for (int i = 1; i <= NoCaIzm; i++)
            {

                DataGridViewTextBoxColumn firstColumn1 =
                new DataGridViewTextBoxColumn();
                firstColumn1.HeaderText = "A; Сер" + i;
                firstColumn1.Name = "A;Ser (" + i;
                firstColumn1.ValueType = Type.GetType("System.Double");

                AgroTable1.Columns.Add(firstColumn1);
                //firstColumn1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
                //   firstColumn1.EditingControlShowing

            }

            for (int i = 1; i <= NoCaIzm; i++)
            {
                AgroTable1.Columns["A;Ser (" + i].Width = 50;
            }
            Concetr.HeaderText = "Конц " + edconctr;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            switch (selet_rezim)
            {
                case 2:
                    Izmerenie1 = true;
                    if (tabControl2.SelectedIndex == 0)
                    {
                        Open();
                    }
                    else
                    {
                        Open1();
                    }
                    break;
                case 1:
                    IzmerenieFR_Open();
                    break;
                case 6:
                    Izmerenie1 = true;
                    if (tabControl2.SelectedIndex == 0)
                    {
                        Open();
                    }
                    else
                    {
                        Open1();
                    }
                    break;
                case 5:
                    TableScan_Open();
                    break;
                case 4:
                    TableKin_Open();
                    break;
            }
            
        }

        public string F1;
        public string F2;
        bool USE_KO_Izmer = false;
        string filepath1;
        string filepath;
        public void TableKin_Open()
        {
            openFileDialog1.InitialDirectory = "C";
            openFileDialog1.Title = "Open File";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "KIN файл|*.KIN2";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                chart3.Series[0].Points.Clear();
                chart3.Series[1].Points.Clear();
                chart3.Series.Clear();
                chart3.Series.Add("Series1");
                chart3.Series.Add("Series2");
                //listBox1.Items.Clear();
                dataGridView4.Rows.Clear();
                dataGridView3.Rows.Clear();
                try
                {
                    // получаем выбранный файл
                    openFileKin(ref filepath);
                    button3.Enabled = true;
                }
                catch (Exception t) { MessageBox.Show("exeption" + t.Message); }
            }
        }
        public void TableScan_Open()
        {
            openFileDialog1.InitialDirectory = "C";
            openFileDialog1.Title = "Open File";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "SCAN файл|*.SCAN2";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ScanChart.Series[0].Points.Clear();
                ScanChart.Series[1].Points.Clear();
                ScanChart.Series.Clear();
                ScanChart.Series.Add("Series1");
                ScanChart.Series.Add("Series2");
                listBox1.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
                try
                {
                    // получаем выбранный файл
                    openFileScan(ref filepath);
                    button3.Enabled = true;
                }
                catch (Exception t) { MessageBox.Show("exeption" + t.Message); }
            }
        }

        public void Open()
        {
            if (selet_rezim == 2)
            {
                openFileDialog1.InitialDirectory = "C";
                openFileDialog1.Title = "Open File";
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "QS2 файл|*.qs2";
            }
            else
            {
                openFileDialog1.InitialDirectory = "C";
                openFileDialog1.Title = "Open File";
                openFileDialog1.FileName = "";
                openFileDialog1.Filter = "Agro QS2 файл|*.aq2";
            }
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                Table1.Visible = false;
                try
                {
                    // получаем выбранный файл
                    openFile(ref filepath);
                }
                catch (Exception t) { MessageBox.Show("exeption" + t.Message); }

                if (SposobZadan != "Ввод коэффициентов")
                {
                    Table1.Visible = true;
                    groupBox2.Enabled = true;
                    groupBox5.Enabled = true;
                    groupBox3.Enabled = true;
                    TableWrite();
                    AgroText0.Enabled = true;
                    AgroText1.Enabled = true;
                    AgroText2.Enabled = true;
                    RR.Enabled = true;
                    SKO.Enabled = true;
                    label21.Enabled = true;
                    label22.Enabled = true;
                    label14.Enabled = true;
                    button11.Enabled = true;
                }
                else
                {
                    groupBox2.Enabled = false;
                    groupBox5.Enabled = false;
                    groupBox3.Enabled = false;
                    RR.Text = "";
                    SKO.Text = "";
                    label21.Text = "";
                    label22.Text = "";
                    button11.Enabled = false;
                }
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
                button3.Enabled = true;
                button9.Enabled = true;


                Podskazka.Text = "Перейдите в Измерения!";
                label27.Visible = false;
                label24.Visible = false;
                label25.Visible = false;
                label26.Visible = false;
                label28.Visible = false;
                if (Convert.ToInt32(CountInSeriya) < 3)
                {
                    radioButton3.Enabled = false;
                }
                else
                {
                    if (Convert.ToInt32(CountInSeriya) < 2)
                    {
                        radioButton2.Enabled = false;
                        radioButton3.Enabled = false;
                    }
                    else
                    {
                        radioButton1.Enabled = true;
                        radioButton2.Enabled = true;
                        radioButton3.Enabled = true;
                    }
                }

                tabPage4.Parent = tabControl2;
                if (selet_rezim == 6)
                {
                    tabControl2.TabPages[1].Text = "Измерение Агро";
                }
            }
        }
        public void TableWrite()
        {

            int StolbecCol = 3 + Convert.ToInt32(CountSeriya2);

            if (USE_KO == false)
            {
                for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
                {
                    for (int j = 0; j < StolbecCol; j++)
                    {

                        Table1.Rows[i].Cells[j].Value = Stolbec[i, j];
                    }

                }
            }
            else
            {
                for (int i = 0; i < (Convert.ToInt32(CountInSeriya2) + 1); i++)
                {
                    for (int j = 0; j < StolbecCol; j++)
                    {

                        Table1.Rows[i].Cells[j].Value = Stolbec[i, j];
                    }

                }
            }
            NoCaIzm = Convert.ToInt32(CountSeriya2);
            if (TimeIzmer1 == "A (C) - градуировочное уравнение (стандарт)")
            {
                radioButton4.Checked = true;
            }
            else
            {
                radioButton5.Checked = true;
            }
            if (TypeYravn1 == "Линейное через 0")
            {
                radioButton1.Checked = true;
                lineinaya0();
            }
            else
            {
                if (TypeYravn1 == "Линейное")
                {
                    radioButton2.Checked = true;
                    lineinaya();
                }
                else
                {
                    radioButton3.Checked = true;
                    kvadratichnaya();
                }
            }
            OpenIzmer = true;
            if (button12.Enabled == true && OpenIzmer == true)
            {
                IzmerCreate = true;
            }
            else
            {
                IzmerCreate = false;
            }
            if (IzmerCreate == true)
            {
                button14.Enabled = true;
            }
            else
            {
                button14.Enabled = false;
            }

        }
        public void openFileKin(ref string filepath)
        {

        }
        public void openFileScan(ref string filepath)
        {
            filepath = openFileDialog1.FileName;
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(@filepath);
            XmlNodeList nodes = xDoc.ChildNodes;

            foreach (XmlNode n in nodes)
            {
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        if ("Izmerenie".Equals(d.Name))
                        {
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {
                                if ("CountIzmer".Equals(k.Name) && k.FirstChild != null)
                                {
                                    for (int i = 0; i < Convert.ToInt32(k.FirstChild.Value); i++)
                                    {
                                        listBox1.Items.Add("Измерение" + i);
                                    }
                                }
                                if ("TypeIzmer".Equals(k.Name) && k.FirstChild != null)
                                {
                                    ScanTable.Columns[1].HeaderText = k.FirstChild.Value;
                                    if (ScanTable.Columns[1].HeaderText == "Abs")
                                    {
                                        ScanTable.Columns[2].HeaderText = "%T";
                                    }
                                    else
                                    {
                                        ScanTable.Columns[2].HeaderText = "Abs";
                                    }
                                    ScanChart.ChartAreas[0].AxisX.Title = ScanTable.Columns["WalveDl"].HeaderText;
                                    ScanChart.ChartAreas[0].AxisY.Title = ScanTable.Columns["Abs_scan"].HeaderText;
                                }
                            }
                        }
                        /*for(int i = 0; i < listBox1.Items.Count; i++)
                        {
                            
                        }*/
                    }
                }
            }
            Array.Resize<string[,]>(ref countScan, listBox1.Items.Count);
            XDocument xdoc = XDocument.Load(@filepath);
            double[] RowsMax;
            foreach (XElement IzmerScan in xdoc.Element("Data_Izmerenie").Elements("NumberIzmer"))
            {
                XElement CountStrElement = IzmerScan.Element("CountStr");
                XAttribute NumberIzmer = IzmerScan.Attribute("Nomer");
                //   MessageBox.Show(NumberIzmer.Value);
                int StrCount = 0;
                if (CountStrElement != null)
                {
                    StrCount = Convert.ToInt32(CountStrElement.Value);

                }
                countScan[Convert.ToInt32(NumberIzmer.Value)] = new string[StrCount, 3];

                RowsMax = new double[StrCount];
                foreach (XElement IzmerScan1 in xdoc.Element("Data_Izmerenie").Element("NumberIzmer").Elements("Str"))
                {
                    XAttribute nameAttribute = IzmerScan1.Attribute("Nomer");
                    XElement RowsElement = IzmerScan1.Element("Cells0");
                    RowsMax[Convert.ToInt32(nameAttribute.Value)] = Convert.ToDouble(RowsElement.Value);
                }

                ScanChart.ChartAreas[0].AxisX.Minimum = RowsMax[StrCount - 1];
                ScanChart.ChartAreas[0].AxisX.Maximum = RowsMax[0];
            }


            //MessageBox.Show(RowsMax);
            foreach (XElement IzmerScan in xdoc.Element("Data_Izmerenie").Elements("NumberIzmer"))
            {
                XElement CountStrElement = IzmerScan.Element("CountStr");
                XAttribute NumberIzmer = IzmerScan.Attribute("Nomer");
                //   MessageBox.Show(NumberIzmer.Value);
                int StrCount = 0;
                if (CountStrElement != null)
                {
                    StrCount = Convert.ToInt32(CountStrElement.Value);

                }

                Application.DoEvents();
                //  if (chart1.Series.Count > 1) { chart1.Series.RemoveAt(Convert.ToInt32(NumberIzmer.Value)+1); }
                ScanChart.Series.Add("area" + Convert.ToInt32(NumberIzmer.Value));
                Random r = new Random();
                int x = r.Next(100, 200), y = r.Next(100, 200);
                ScanChart.Series["area" + Convert.ToInt32(NumberIzmer.Value)].Color =
                    System.Drawing.Color.FromArgb(
                        (byte)(x - 2 * y),
                        (byte)(y + x),
                        (byte)(y - 2 * x));
                ScanChart.Series["area" + Convert.ToInt32(NumberIzmer.Value)].ChartType = SeriesChartType.Line;
                //   TableScan();
                // Application.DoEvents();
                ///    RowsMax = new String[StrCount];
                double[] massWL = new double[StrCount];
                double[] massGE = new double[StrCount];
                foreach (XElement IzmerScan1 in IzmerScan.Elements("Str"))
                {

                    XAttribute nameAttribute = IzmerScan1.Attribute("Nomer");
                    XElement CellsElement0 = IzmerScan1.Element("Cells0");
                    XElement CellsElement1 = IzmerScan1.Element("Cells1");
                    XElement CellsElement2 = IzmerScan1.Element("Cells2");
                    if (nameAttribute != null && CellsElement0 != null && CellsElement1 != null && CellsElement2 != null)
                    {
                        countScan[Convert.ToInt32(NumberIzmer.Value)][Convert.ToInt32(nameAttribute.Value), 0] = CellsElement0.Value;
                        countScan[Convert.ToInt32(NumberIzmer.Value)][Convert.ToInt32(nameAttribute.Value), 1] = CellsElement1.Value;
                        countScan[Convert.ToInt32(NumberIzmer.Value)][Convert.ToInt32(nameAttribute.Value), 2] = CellsElement2.Value;
                        double x1 = Convert.ToDouble(CellsElement0.Value);
                        double y1 = Convert.ToDouble(CellsElement1.Value);
                        ScanChart.Series["area" + Convert.ToInt32(NumberIzmer.Value)].Points.AddXY(x1, y1);
                        massWL[Convert.ToInt32(nameAttribute.Value)] = Convert.ToDouble(CellsElement0.Value);
                        massGE[Convert.ToInt32(nameAttribute.Value)] = Convert.ToDouble(CellsElement1.Value);

                    }


                }
                double max = 0.0;
                double min = 0.0;

                for (int i = 0; i < StrCount; i++)
                {
                    if (i == 0)
                    {
                        if (massGE[i] > massGE[i + 1])
                        {
                            max = massGE[i];
                            //dataGridView1.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                            min = max;
                            double x1 = massWL[i];
                            double y1 = massGE[i];
                            ScanChart.Series[0].Points.AddXY(x1, y1);
                            ScanChart.Series[0].ChartType = SeriesChartType.Point;
                        }
                        else
                        {
                            min = massGE[i];
                            //  dataGridView2.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                            max = min;
                            double x1 = massWL[i];
                            double y1 = massGE[i];
                            ScanChart.Series[0].Points.AddXY(x1, y1);
                            ScanChart.Series[0].ChartType = SeriesChartType.Point;
                        }

                    }
                    else {
                        if (i + 1 != StrCount)
                        {
                            if (massGE[i] > massGE[i - 1] && massGE[i] >= massGE[i + 1])
                            {
                                max = massGE[i];
                                min = max;
                                // dataGridView1.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                                double x1 = massWL[i];
                                double y1 = massGE[i];
                                ScanChart.Series[0].Points.AddXY(x1, y1);
                                ScanChart.Series[0].ChartType = SeriesChartType.Point;
                            }
                            if (massGE[i] < massGE[i - 1] && massGE[i] <= massGE[i + 1])
                            {
                                min = massGE[i];
                                //  dataGridView2.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                                max = min;
                                double x1 = massWL[i];
                                double y1 = massGE[i];
                                ScanChart.Series[0].Points.AddXY(x1, y1);
                                ScanChart.Series[0].ChartType = SeriesChartType.Point;
                            }
                        }
                    }
                }

                listBox1.SetSelected(0, true);

            }



        }
        public void openFile(ref string filepath)
        {
            WLREMOVE1();
            WLREMOVESTR1();
            filepath = openFileDialog1.FileName;
            параметрыToolStripMenuItem.Enabled = true;
            button10.Enabled = true;

            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(@filepath);

            XmlNodeList nodes = xDoc.ChildNodes;

            foreach (XmlNode n in nodes)
            { // Обрабатываем в цикле только Data_Izmerenie
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    // Читаем в цикле вложенные значения Izmerenie
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        // Обрабатываем в цикле только Izmerenie
                        if ("Izmerenie".Equals(d.Name))
                        {
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {
                                if ("Version".Equals(k.Name) && k.FirstChild != null)
                                {
                                    if (version != k.FirstChild.Value)
                                    {
                                        MessageBox.Show("Внимание, версия программы отличается от версии открываемого документа!\n" +
                                    "Рекомендуем создать новую градуировку!");
                                        break;
                                    }
                                    else
                                    {
                                        // MessageBox.Show("Версия совпадает!");
                                        break;
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Внимание, версия программы отличается от версии открываемого документа!\n" +
"Рекомендуем создать новую градуировку!");
                                    break;
                                }

                            }
                        }
                    }
                }
            }


            foreach (XmlNode n in nodes)
            { // Обрабатываем в цикле только Data_Izmerenie
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    // Читаем в цикле вложенные значения Izmerenie
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        // Обрабатываем в цикле только Izmerenie
                        if ("Izmerenie".Equals(d.Name))
                        {
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {
                                if ("USE_CO_XML".Equals(k.Name) && k.FirstChild != null)
                                {
                                    USE_CO_XML1 = k.FirstChild.Value;
                                    if (USE_CO_XML1 == "true")
                                    {
                                        USE_KO = true;
                                    }
                                    else
                                    {
                                        USE_KO = false;
                                    }

                                }
                                if ("SposobZadan".Equals(k.Name) && k.FirstChild != null)
                                {
                                    SposobZadan = k.FirstChild.Value;

                                }
                            }
                        }
                    }
                }
            }
            // Обходим значения
            foreach (XmlNode n in nodes)
            { // Обрабатываем в цикле только Data_Izmerenie
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    // Читаем в цикле вложенные значения Izmerenie
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        // Обрабатываем в цикле только Izmerenie
                        if ("Izmerenie".Equals(d.Name))
                        {
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {

                                if ("Veshestvo".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Veshestvo1 = k.FirstChild.Value; //Вещество
                                    textBox11.Text = Veshestvo1;
                                    textBox12.Text = Veshestvo1;

                                }
                                if ("wavelength".Equals(k.Name) && k.FirstChild != null)
                                {
                                    wavelength1 = k.FirstChild.Value; //Длина волны
                                    textBox9.Text = wavelength1;
                                    textBox10.Text = wavelength1;
                                }
                                if ("WidthCuvet".Equals(k.Name) && k.FirstChild != null)
                                {
                                    WidthCuvette = k.FirstChild.Value; //Ширина кюветы
                                    textBox2.Text = WidthCuvette;

                                }
                                if ("BottomLine".Equals(k.Name) && k.FirstChild != null)
                                {
                                    BottomLine = k.FirstChild.Value; //Нижняя граница
                                }
                                if ("TopLine".Equals(k.Name) && k.FirstChild != null)
                                {
                                    TopLine = k.FirstChild.Value; //Верхняя граница
                                }

                                if ("ND".Equals(k.Name) && k.FirstChild != null)
                                {
                                    ND = k.FirstChild.Value; //НД
                                }
                                if ("Description".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Description = k.FirstChild.Value; //Примечание
                                    textBox1.Text = Description;

                                }
                                if ("DateTime".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime = k.FirstChild.Value; //Дата
                                    dateTimePicker1.Text = DateTime;
                                }
                                if ("DateTime1_1".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime2_1 = k.FirstChild.Value; //Дата
                                    label6.Text = DateTime2_1;
                                }

                                if ("DateTime1_1_1".Equals(k.Name) && k.FirstChild != null)
                                {
                                    DateTime2_2_1 = k.FirstChild.Value; //Дата
                                    numericUpDown1.Value = Convert.ToInt32(DateTime2_2_1);
                                }
                                if ("Pogreshnost".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Pogreshnost2 = k.FirstChild.Value; //Дата
                                    textBox3.Text = Pogreshnost2;
                                }
                                if ("Ispolnitel".Equals(k.Name) && k.FirstChild != null)
                                {
                                    Ispolnitel = k.FirstChild.Value; //Исполнитель
                                }
                                /*  Stolbec = new string[this.Table1.Rows.Count - 1, this.Table1.Columns.Count];
                                      for (int i = 0; i < this.Table1.Rows.Count - 1; i++)
                                  {
                                      for (int j = 0; j < this.Table1.Columns.Count; j++)
                                      {
                                          if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                          {
                                              Stolbec[i, j] = k.FirstChild.Value;
                                              Table1.Rows[i].Cells[j].Value = Stolbec[i, j];
                                          }
                                      }
                                  }*/
                                if ("TypeIzmer".Equals(k.Name) && k.FirstChild != null)
                                {
                                    TimeIzmer1 = k.FirstChild.Value;

                                }

                                if ("TypeYravn".Equals(k.Name) && k.FirstChild != null)
                                {
                                    TypeYravn1 = k.FirstChild.Value; //Исполнитель
                                    if (SposobZadan != "Ввод коэффициентов")
                                    {
                                        if (TypeYravn1 == "Линейное через 0" || TypeYravn1 == "Линейное")
                                        {
                                            //  MessageBox.Show("Линейное");
                                            Table1.Columns.Add("X*X", "Конц*Конц");
                                            Table1.Columns.Add("X*Y", "Асред*Конц");
                                            /*  Table1.Columns["X*X"].Width = 50;
                                              Table1.Columns["X*Y"].Width = 50;
                                              Table1.Columns["X*X*X"].Width = 50;
                                              Table1.Columns["X*X*X*X"].Width = 50;
                                              Table1.Columns["X*X*Y"].Width = 50;*/
                                        }
                                        else
                                        {
                                            // MessageBox.Show("Квадратичное");
                                            Table1.Columns.Add("X*X", "Конц* Конц");
                                            Table1.Columns.Add("X*Y", "Асред* Конц");
                                            Table1.Columns.Add("X*X*X", "Асред ^3");
                                            Table1.Columns.Add("X*X*X*X", "Асред ^4");
                                            Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                                            /*    Table1.Columns["X*X"].Width = 50;
                                                Table1.Columns["X*Y"].Width = 50;
                                                Table1.Columns["X*X*X"].Width = 50;
                                                Table1.Columns["X*X*X*X"].Width = 50;
                                                Table1.Columns["X*X*Y"].Width = 50;*/
                                        }
                                    }
                                }
                                if (SposobZadan != "Ввод коэффициентов")
                                {
                                    if ("CountSeriyal".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        CountSeriya2 = k.FirstChild.Value; //Количество столбцов
                                        CountSeriya = CountSeriya2;
                                        while (true)
                                        {
                                            int i = Table1.Columns.Count - 1;//С какого столбца начать
                                            if (Table1.Columns[i].Name == "Asred")
                                                break;
                                            Table1.Columns.RemoveAt(i);
                                        }
                                        for (int i = 1; i <= Convert.ToInt32(CountSeriya2); i++)
                                        {

                                            DataGridViewTextBoxColumn firstColumn2 =
                                            new DataGridViewTextBoxColumn();
                                            firstColumn2.HeaderText = "A; Сер" + i;
                                            firstColumn2.Name = "A;Ser (" + i;
                                            Table1.Columns.Add(firstColumn2);
                                        }
                                        for (int i = 1; i <= Convert.ToInt32(CountSeriya2); i++)
                                        {
                                            Table1.Columns["A;Ser (" + i].Width = 50;
                                        }
                                        Concetr.HeaderText = "Конц " + edconctr;
                                    }

                                    if ("edconctr".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        edconctr = k.FirstChild.Value;
                                        Concetr.HeaderText = "Конц " + edconctr;
                                    }
                                    if ("CountInSeriyal".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        CountInSeriya2 = k.FirstChild.Value; //Количество строк
                                        CountInSeriya = CountInSeriya2;
                                        NoCaSer = Convert.ToInt32(CountInSeriya);
                                        if (USE_KO == false)
                                        {
                                            for (int i = 0; i < NoCaSer; i++)
                                            {
                                                Table1.Rows.Add();
                                            }
                                        }
                                        else
                                        {
                                            for (int i = 0; i < (NoCaSer + 1); i++)
                                            {
                                                Table1.Rows.Add();
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if ("k0".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        AgroText0.Text = k.FirstChild.Value;
                                        k0 = Convert.ToDouble(AgroText0.Text);

                                    }
                                    if ("k1".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        AgroText1.Text = k.FirstChild.Value;
                                        k1 = Convert.ToDouble(AgroText1.Text);

                                    }
                                    if ("k2".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        AgroText2.Text = k.FirstChild.Value;

                                        k2 = Convert.ToDouble(AgroText2.Text);
                                    }
                                }

                            }

                        }
                    }
                    if (TypeYravn1 == "Линейное через 0")
                    {
                        aproksim = "Линейная через 0";
                    }
                    else
                    {
                        if (TypeYravn1 == "Линейное")
                        {
                            aproksim = "Линейная";
                        }
                        else
                        {
                            aproksim = "Квадратичная";
                        }
                    }
                    if (SposobZadan != "Ввод коэффициентов")
                    {
                        if (USE_KO == false)
                        {
                            if (TypeYravn1 == "Линейное через 0" || TypeYravn1 == "Линейное")
                            {
                                StolbecCol = 5 + Convert.ToInt32(CountSeriya2);
                            }
                            else
                            {
                                StolbecCol = 8 + Convert.ToInt32(CountSeriya2);
                            }
                            Stolbec = new string[Convert.ToInt32(CountInSeriya2), StolbecCol];
                        }
                        else
                        {
                            if (TypeYravn1 == "Линейное через 0" || TypeYravn1 == "Линейное")
                            {
                                StolbecCol = 5 + Convert.ToInt32(CountSeriya2);
                            }
                            else
                            {
                                StolbecCol = 8 + Convert.ToInt32(CountSeriya2);
                            }
                            Stolbec = new string[(Convert.ToInt32(CountInSeriya2) + 1), StolbecCol];
                        }
                        int stroka = 0;

                        // Читаем в цикле вложенные значения Stroka

                        for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                        {

                            // Обрабатываем в цикле только Stroka
                            if ("Stroka".Equals(d.Name))
                            {
                                int stolbec = 0;
                                //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                                for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                                {


                                    if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                    {

                                        Stolbec[stroka, stolbec] = k.FirstChild.Value;


                                        stolbec++;
                                    }

                                }
                                stroka++;
                            }

                        }
                    }
                    else
                    {
                        functionA();
                    }
                }
            }
            if (ComPort == true)
            {
                SW2();
            }
            else
            {
                //   GWNew.Text = wavelength1;
            }


        }
        string filepathIzmarFR;
        public void IzmerenieFR_Open()
        {
            openFileDialog1.InitialDirectory = "C";
            openFileDialog1.Title = "Open File";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "ISFR файл|*.isfr2";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    filepathIzmarFR = openFileDialog1.FileName;
                    // получаем выбранный файл
                    IzmerenieFR_openFile(ref filepathIzmarFR);
                    button11.Enabled = true;
                    //  button10.Enabled = true;
                    button3.Enabled = true;
                    button9.Enabled = true;
                }
                catch (Exception t) { MessageBox.Show("exeption" + t.Message); }
            }
        }

        public void Open1()
        {
            openFileDialog1.InitialDirectory = "C";
            openFileDialog1.Title = "Open File";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "QA2 файл|*.qa2";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // получаем выбранный файл
                    openFile2(ref filepath2, ref filepath);
                }
                catch (Exception t) { MessageBox.Show("exeption" + t.Message); }



            }
        }
        public void IzmerenieFR_openFile(ref string filepathIzmarFR)
        {
            IzmerenieFR_RowsRemove2();
            string RowsCount = "";
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(@filepathIzmarFR);

            XmlNodeList nodes = xDoc.ChildNodes;
            XDocument xdoc = XDocument.Load(filepathIzmarFR);
            foreach (XElement IzmerenieElement in xdoc.Element("Data_Izmerenie").Elements("Izmerenie"))
            {
                XElement countIzmer1Element = IzmerenieElement.Element("countIzmer1");
                XElement DescriptionElement = IzmerenieElement.Element("Description");
                XElement DateTimeElement = IzmerenieElement.Element("DateTime");
                XElement IspolnitelElement = IzmerenieElement.Element("Ispolnitel");
                if (countIzmer1Element != null && DescriptionElement != null && DateTimeElement != null && IspolnitelElement != null)
                {
                    DateTime = DateTimeElement.Value;
                    Description = DescriptionElement.Value;
                    Ispolnitel = IspolnitelElement.Value;
                    RowsCount = countIzmer1Element.Value; //Количество строк

                    for (int i = 0; i < Convert.ToInt32(RowsCount); i++)
                    {
                        IzmerenieFR_Table.Rows.Add();
                    }
                    StolbecCol_1 = 7;

                    Stolbec_1 = new string[Convert.ToInt32(RowsCount), StolbecCol_1];
                }
            }

            int stroka = 0;

            // Читаем в цикле вложенные значения Stroka
            foreach (XmlNode n in nodes)
            {
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {

                        // Обрабатываем в цикле только Stroka
                        if ("Stroka".Equals(d.Name))
                        {
                            int stolbec = 0;
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {


                                if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                {

                                    Stolbec_1[stroka, stolbec] = k.FirstChild.Value;


                                    stolbec++;
                                }

                            }
                            stroka++;
                        }

                    }
                }
            }
            IzmerenieFR_Table_Write();


        }
        public void openFile2(ref string filepath2, ref string filepath)
        {
            WLREMOVE2();
            WLREMOVESTR2();
            filepath2 = openFileDialog1.FileName;

            параметрыToolStripMenuItem.Enabled = true;
            button10.Enabled = true;
            bool NotFile = false;
            XmlDocument xDoc = new XmlDocument();

            xDoc.Load(@filepath2);

            XmlNodeList nodes = xDoc.ChildNodes;
            foreach (XmlNode n in nodes)
            { // Обрабатываем в цикле только Data_Izmerenie
                if ("Data_Izmerenie".Equals(n.Name))
                {
                    // Читаем в цикле вложенные значения Izmerenie


                    for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                    {
                        // Обрабатываем в цикле только Izmerenie
                        if ("Izmerenie".Equals(d.Name))
                        {
                            //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                            for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                            {
                                if ("filepath".Equals(k.Name) && k.FirstChild != null)
                                {
                                    filepath1 = k.FirstChild.Value;
                                    if (filepath1 == filepath)
                                    {
                                        NotFile = true;



                                    }
                                    else
                                    {
                                        NotFile = false;
                                    }

                                }
                            }
                        }
                    }
                }
            }
            if (NotFile == false)
            {
                MessageBox.Show("Внимание!! Открытый файл измерения не соответсвует файлу градуировки!");
            }
            if (NotFile == true)
            {
                foreach (XmlNode n in nodes)
                { // Обрабатываем в цикле только Data_Izmerenie
                    if ("Data_Izmerenie".Equals(n.Name))
                    {
                        // Читаем в цикле вложенные значения Izmerenie


                        for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                        {
                            // Обрабатываем в цикле только Izmerenie
                            if ("Izmerenie".Equals(d.Name))
                            {
                                //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                                for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                                {
                                    if ("USE_CO_XML".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        USE_CO_XML1 = k.FirstChild.Value;
                                        if (USE_CO_XML1 == "true")
                                        {
                                            USE_KO_Izmer = true;
                                        }
                                        else
                                        {
                                            USE_KO_Izmer = false;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Обходим значения
                foreach (XmlNode n in nodes)
                { // Обрабатываем в цикле только Data_Izmerenie
                    if ("Data_Izmerenie".Equals(n.Name))
                    {
                        // Читаем в цикле вложенные значения Izmerenie
                        for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                        {
                            // Обрабатываем в цикле только Izmerenie
                            if ("Izmerenie".Equals(d.Name))
                            {
                                //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                                for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                                {

                                    if ("WidthCuvet".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        WidthCuvette = k.FirstChild.Value; //Ширина кюветы
                                        int index = Opt_dlin_cuvet.FindString(WidthCuvette);
                                        //  MessageBox.Show(index.ToString());
                                        Opt_dlin_cuvet.SelectedIndex = index;
                                    }

                                    if ("Description".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        Description = k.FirstChild.Value; //Примечание
                                        textBox8.Text = Description;
                                    }
                                    if ("DateTime".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        DateTime = k.FirstChild.Value; //Дата
                                        dateTimePicker2.Text = DateTime;
                                    }

                                    if ("Pogreshnost".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        Pogreshnost2 = k.FirstChild.Value; //Дата
                                        textBox7.Text = Pogreshnost2;
                                    }
                                    if ("F1".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        F1 = k.FirstChild.Value; //F1
                                        F1Text.Text = F1;
                                    }
                                    if ("F2".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        F2 = k.FirstChild.Value; //F1
                                        F2Text.Text = F2;
                                    }

                                    if ("CountSeriyal".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        CountSeriya2 = k.FirstChild.Value; //Количество столбцов
                                        while (true)
                                        {
                                            int i = Table2.Columns.Count - 1;//С какого столбца начать
                                            if (Table2.Columns[i].Name == "Obrazec")
                                                break;
                                            Table2.Columns.RemoveAt(i);
                                        }

                                        for (int i = 1; i <= Convert.ToInt32(CountSeriya2); i++)
                                        {

                                            DataGridViewTextBoxColumn firstColumn2 =
                                            new DataGridViewTextBoxColumn();
                                            firstColumn2.HeaderText = "A; Сер." + i;
                                            firstColumn2.Name = "A;Ser" + i;
                                            firstColumn2.ValueType = Type.GetType("System.Double");
                                            Table2.Columns.Add(firstColumn2);
                                            DataGridViewTextBoxColumn firstColumn3 =
                                            new DataGridViewTextBoxColumn();
                                            firstColumn3.HeaderText = "C, " + edconctr + "; Сер." + i;
                                            firstColumn3.Name = "C,edconctr;Ser." + i;
                                            firstColumn3.ValueType = Type.GetType("System.Double");
                                            Table2.Columns.Add(firstColumn3);
                                            firstColumn3.Width = 50;
                                            firstColumn2.Width = 50;

                                        }
                                        DataGridViewTextBoxColumn firstColumn4 =
                                        new DataGridViewTextBoxColumn();
                                        firstColumn4.HeaderText = "Cср, " + edconctr;
                                        firstColumn4.Name = "Ccr";
                                        firstColumn4.ReadOnly = true;
                                        firstColumn4.ValueType = Type.GetType("System.Double");
                                        Table2.Columns.Add(firstColumn4);

                                        DataGridViewTextBoxColumn firstColumn5 =
                                        new DataGridViewTextBoxColumn();
                                        firstColumn5.HeaderText = "d, %";
                                        firstColumn5.Name = "d%";
                                        firstColumn5.ReadOnly = true;
                                        firstColumn5.ValueType = Type.GetType("System.Double");
                                        Table2.Columns.Add(firstColumn5);
                                        firstColumn4.Width = 100;
                                        firstColumn5.Width = 50;
                                    }
                                    if ("CountInSeriyal".Equals(k.Name) && k.FirstChild != null)
                                    {
                                        CountInSeriya2 = k.FirstChild.Value; //Количество строк
                                        NoCaSer1 = Convert.ToInt32(CountInSeriya2);
                                        if (USE_KO == false)
                                        {
                                            for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
                                            {
                                                Table2.Rows.Add();
                                            }
                                            StolbecCol_1 = 2 + Convert.ToInt32(CountSeriya2) + Convert.ToInt32(CountSeriya2) + 2;

                                            Stolbec_1 = new string[Convert.ToInt32(CountInSeriya2), StolbecCol_1];
                                            //Table2.Rows.Add();
                                        }
                                        else
                                        {
                                            if (USE_KO_Izmer == false && USE_KO == true)
                                            {
                                                Table2.Rows.Add(0, "Контрольный");
                                                for (int i = 0; i < Convert.ToInt32(CountInSeriya2) + 1; i++)
                                                {
                                                    Table2.Rows.Add();
                                                }
                                                StolbecCol_1 = 2 + Convert.ToInt32(CountSeriya2) + Convert.ToInt32(CountSeriya2) + 2;

                                                Stolbec_1 = new string[Convert.ToInt32(CountInSeriya2) + 1, StolbecCol_1];
                                                //Table2.Rows.Add();
                                            }
                                            else
                                            {
                                                for (int i = 0; i < (Convert.ToInt32(CountInSeriya2) + 1); i++)
                                                {
                                                    Table2.Rows.Add();
                                                }
                                                // Table2.Rows.Add();
                                                StolbecCol_1 = 2 + Convert.ToInt32(CountSeriya2) + Convert.ToInt32(CountSeriya2) + 2;

                                                Stolbec_1 = new string[(Convert.ToInt32(CountInSeriya2) + 1), StolbecCol_1];
                                            }
                                        }
                                    }

                                }

                            }
                        }

                        int stroka = 0;

                        // Читаем в цикле вложенные значения Stroka

                        for (XmlNode d = n.FirstChild; d != null; d = d.NextSibling)
                        {

                            // Обрабатываем в цикле только Stroka
                            if ("Stroka".Equals(d.Name))
                            {
                                int stolbec = 0;
                                //Можно, например, в этом цикле, да и не только..., взять какие-то данные
                                for (XmlNode k = d.FirstChild; k != null; k = k.NextSibling)
                                {


                                    if ("Stolbec".Equals(k.Name) && k.FirstChild != null)
                                    {

                                        Stolbec_1[stroka, stolbec] = k.FirstChild.Value;


                                        stolbec++;
                                    }

                                }
                                stroka++;
                            }

                        }
                    }
                }
                TableWrite2();
                button3.Enabled = true;
                button9.Enabled = true;
                печатьToolStripMenuItem1.Enabled = true;
            }



        }
        public void IzmerenieFR_Table_Write()
        {
            int StolbecCol_1 = 7;
            for (int i = 0; i < (Stolbec_1.Length / StolbecCol_1); i++)
            {
                for (int j = 0; j < StolbecCol_1; j++)
                {

                    IzmerenieFR_Table.Rows[i].Cells[j].Value = Stolbec_1[i, j];
                }

            }
        }
        public void TableWrite2()
        {

            int StolbecCol_1 = 2 + Convert.ToInt32(CountSeriya2) + Convert.ToInt32(CountSeriya2) + 2;

            if (USE_KO == false)
            {
                for (int i = 0; i < Convert.ToInt32(CountInSeriya2); i++)
                {
                    for (int j = 0; j < StolbecCol_1; j++)
                    {

                        Table2.Rows[i].Cells[j].Value = Stolbec_1[i, j];
                    }

                }
            }
            else
            {
                if (USE_KO_Izmer == false && USE_KO == true)
                {
                    for (int i = 1; i < (Convert.ToInt32(CountInSeriya2) + 1); i++)
                    {
                        for (int j = 0; j < StolbecCol_1; j++)
                        {

                            Table2.Rows[i].Cells[j].Value = Stolbec_1[i - 1, j];
                        }

                    }
                }
                else
                {
                    for (int i = 0; i < (Convert.ToInt32(CountInSeriya2) + 1); i++)
                    {
                        for (int j = 0; j < StolbecCol_1; j++)
                        {

                            Table2.Rows[i].Cells[j].Value = Stolbec_1[i, j];
                        }

                    }
                }
            }
            NoCaIzm1 = Convert.ToInt32(CountSeriya2);
            OpenIzmer1 = true;
            if (button12.Enabled == true && OpenIzmer1 == true)
            {
                IzmerCreate = true;
            }
            else
            {
                IzmerCreate = false;
            }
            if (IzmerCreate == true)
            {
                button14.Enabled = true;
            }
            else
            {
                button14.Enabled = false;
            }

        }
        public void SW2()
        {
            LogoForm2();
            string SWText1 = wavelength1;
            newPort.Write("SW " + wavelength1 + "\r");
            //  Thread.Sleep(20000);
            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else {
                    indata = newPort.ReadExisting();
                }
            }
            GWNew.Text = wavelength1;
            SWF.Application.OpenForms["LogoFrm2"].Close();
            // GW();
        }
        public void functionA()
        {
            groupBox2.Enabled = false;
            groupBox5.Enabled = false;
            groupBox3.Enabled = false;
            RR.Text = "";
            SKO.Text = "";
            label21.Text = "";
            label22.Text = "";
            // chart1.Series[0].Points.Clear();
            //   chart1.Series[1].Points.Clear();
            if (Zavisimoct == "A(C)")
            {
                if (aproksim == "Линейная через 0")
                {

                    label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
                    double x2 = 0;
                    for (double i = 0; i < 5; i = i + 0.5)
                    {
                        double y2 = i * k1;
                        chart1.Series[1].Points.AddXY(i, y2);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisX.Minimum = 0;

                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        x2 = i;
                    }
                    double xfin = x2 * 1.1;
                    double yfin = xfin * k1;
                    chart1.Series[1].Points.AddXY(xfin, yfin);
                }
                else
                {
                    if (aproksim == "Линейная")
                    {
                        label14.Text = "A(C) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000 ;- 0.0000 ") + "*C ";
                        double x2 = 0;
                        for (double i = 0; i < 5; i = i + 0.5)
                        {
                            double y2 = i * k1 + k0;
                            chart1.Series[1].Points.AddXY(i, y2);
                            chart1.Series[1].ChartType = SeriesChartType.Line;
                            chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                            chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                            chart1.ChartAreas[0].AxisX.Minimum = 0;

                            chart1.ChartAreas[0].AxisY.Minimum = 0;
                            x2 = i;
                        }
                        double xfin = x2 * 1.1;
                        double yfin = xfin * k1 + k0;
                        chart1.Series[1].Points.AddXY(xfin, yfin);
                    }
                    else
                    {
                        label14.Text = "A(C) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000 ;- 0.0000 ") + "*C " + k2.ToString("+ 0.0000 ;- 0.0000 ") + "*C^2";
                        double x2 = 0;
                        for (double i = 0; i < 5; i = i + 0.5)
                        {
                            double y2 = i * k1 + i * i * k2 + k0;
                            chart1.Series[1].Points.AddXY(i, y2);
                            chart1.Series[1].ChartType = SeriesChartType.Line;
                            chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                            chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                            chart1.ChartAreas[0].AxisX.Minimum = 0;

                            chart1.ChartAreas[0].AxisY.Minimum = 0;
                            x2 = i;
                        }
                        double xfin = x2 * 1.1;
                        double yfin = xfin * k1;
                        chart1.Series[1].Points.AddXY(xfin, yfin);
                    }


                }
            }
            else
            {
                if (aproksim == "Линейная через 0")
                {
                    label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A";
                    double x2 = 0;
                    for (double i = 0; i < 5; i = i + 0.5)
                    {
                        double y2 = i * k1;
                        chart1.Series[1].Points.AddXY(i, y2);
                        chart1.Series[1].ChartType = SeriesChartType.Line;
                        chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                        chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                        chart1.ChartAreas[0].AxisX.Minimum = 0;

                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        x2 = i;
                    }
                    double xfin = x2 * 1.1;
                    double yfin = xfin * k1;
                    chart1.Series[1].Points.AddXY(xfin, yfin);
                }
                else
                {
                    if (aproksim == "Линейная")
                    {
                        label14.Text = "C(A) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000 ;- 0.0000 ") + "*A ";
                        double x2 = 0;
                        for (double i = 0; i < 5; i = i + 0.5)
                        {
                            double y2 = i * k1 + k0;
                            chart1.Series[1].Points.AddXY(i, y2);
                            chart1.Series[1].ChartType = SeriesChartType.Line;
                            chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                            chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                            chart1.ChartAreas[0].AxisX.Minimum = 0;

                            chart1.ChartAreas[0].AxisY.Minimum = 0;
                            x2 = i;
                        }
                        double xfin = x2 * 1.1;
                        double yfin = xfin * k1;
                        chart1.Series[1].Points.AddXY(xfin, yfin);
                    }
                    else
                    {
                        label14.Text = "C(A) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000 ;- 0.0000 ") + "*A " + k2.ToString("+ 0.0000 ;- 0.0000 ") + "*A^2";
                        double x2 = 0;
                        for (double i = 0; i < 5; i = i + 0.5)
                        {
                            double y2 = i * k1 + i * k2 + k0;
                            chart1.Series[1].Points.AddXY(i, y2);
                            chart1.Series[1].ChartType = SeriesChartType.Line;
                            chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                            chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                            chart1.ChartAreas[0].AxisX.Minimum = 0;

                            chart1.ChartAreas[0].AxisY.Minimum = 0;
                            x2 = i;
                        }
                        double xfin = x2 * 1.1;
                        double yfin = xfin * k1;
                        chart1.Series[1].Points.AddXY(xfin, yfin);
                    }

                }
            }
        }
        int rowIndex;
        public double Asred1;
        int currentCelledit = 0;
        int rowcelledit = 0;

        public void functionAsred()
        {
            //Table1.Rows.Add();
            while (true)
            {
                int ml = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns.Count == 3 + NoCaIzm)
                    break;
                Table1.Columns.RemoveAt(ml);
            }
            groupBox3.Enabled = true;
            groupBox2.Enabled = true;

            if (radioButton1.Checked == true)
            {
                lineinaya0();
            }
            else
            {
                if (radioButton2.Checked == true)
                {

                    lineinaya();
                }
                else
                {
                    kvadratichnaya();
                }
            }
            Podskazka.Text = "Сохраните градуировку!";
            label27.Visible = true;
            label24.Visible = false;
            label25.Visible = false;
            label26.Visible = false;
            label28.Visible = false;
            label33.Visible = false;
            label14.Enabled = true;
            RR.Enabled = true;
            SKO.Enabled = true;
            label21.Enabled = true;
            label22.Enabled = true;
            AgroText0.Enabled = true;
            AgroText1.Enabled = true;
            AgroText2.Enabled = true;
        }
        int circle;
        public void lineinaya0()
        {

            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            circle = 0;
            XY = 0;
            SUMMY2 = 0;
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", 0);
            AgroText2.Text = string.Format("{0:0.0000}", 0);
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double SUMMSer = 0;
            double SREDSUMMX = 0;

            if (radioButton4.Checked == true)

            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");

                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");

                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;

                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;

                }
                catch
                {
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");

                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;

                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;


                }
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                k0 = 0; k1 = 0; k2 = 0;
                circle = 0;
                XY = 0;
                SUMMY2 = 0;
                double SUMMX = 0;

                if (USE_KO == false)
                {
                    USE_KO_lineinaya0_not();

                }
                else
                {
                    USE_KO_lineinaya0();
                }

            }
            else
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;

                }
                if (USE_KO == false)
                {
                    USE_KO_lineinaya0_not1();
                }
                else
                {
                    USE_KO_lineinaya01();
                }
            }

        }
        public void USE_KO_lineinaya0_not()
        {
            double max = -1;
            int index = -1;
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double[] Table1masStr_1 = new double[NoCaIzm];
            double[] Table1masStr1_1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr_1[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr_1);
                double maxEl = Table1masStr_1[Table1masStr_1.Length - 1];
                double minEl = Table1masStr_1[0];
                double p1 = 2 * ((maxEl - minEl) / (maxEl + minEl)) * 100;
                //  Table1.Rows[i].Cells["P"].Value = string.Format("{0:0.0000}", p1);
                Table1masStr1_1[i] = p1;
            }
            max = -1;
            for (int i = 1; i <= Table1masStr1_1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1_1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1_1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label21.Text = "P,% = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                XY += x * y;
                SUMMY2 += y * y;
                SUMMX += x;

                Table1.Rows[i].Cells["X*X"].Value = y * y;
                Table1.Rows[i].Cells["X*Y"].Value = x * y;
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);

            }

            SREDSUMMX = SUMMX / (Table1.Rows.Count - 1);
            k1 = XY / SUMMY2;

            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index+1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(A) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - xrasch) * 100) / xrasch);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double x2 = 0;
            int Table1_Asred = 0;
            label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (y1 - k1 * x1) * (y1 - k1 * x1);
                yx1 += (y1 - SREDSUMM) * (y1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY(x1, y1);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;


                //     chart1.Series[0].Points.Item.Label = Convert.ToString(circle);
                // chart1.Series[0].IsValueShownAsLabel = true;
                //chart1.Series[0].IsXValueIndexed = true;
                // circle++;
                // double x2 = 0.1 * i;
                // double y2 = x2 / k1;
                x2 = 0;
                if (Table1_Asred == 0)
                {
                    x2 = 0;
                }
                else
                {
                    x2 = x1;
                }
                Table1_Asred++;
                double y2 = x2 * k1;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2;
                //    chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);
                //   chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        public void USE_KO_lineinaya0()
        {
            double max = -1;
            int index = -1;
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            max = -1;
            double[] Table1masStr_1 = new double[NoCaIzm];
            double[] Table1masStr1_1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr_1[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr_1);
                double maxEl = Table1masStr_1[Table1masStr_1.Length - 1];
                double minEl = Table1masStr_1[0];
                double p1 = 2 * ((maxEl - minEl) / (maxEl + minEl)) * 100;
                //  Table1.Rows[i].Cells["P"].Value = string.Format("{0:0.0000}", p1);
                Table1masStr1_1[i] = p1;
            }
            max = -1;
            for (int i = 1; i <= Table1masStr1_1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1_1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1_1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label21.Text = "P,% = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double x1 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double y1 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
            ///Table1.Rows[0].Cells["X*X"].Value = y1 * y1;
            /// Table1.Rows[0].Cells["X*Y"].Value = (x1) * y1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                XY += (x - x1) * y;
                SUMMY2 += y * y;
                SUMMX += (x - x1);

                Table1.Rows[i].Cells["X*X"].Value = y * y;
                Table1.Rows[i].Cells["X*Y"].Value = (x - x1) * y;
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);

            }
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value);

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(A) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }


            SREDSUMMX = SUMMX / (Table1.Rows.Count - 1);
            k1 = XY / SUMMY2;

            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText0.Text = string.Format("{0:0.0000}", 0);

            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - xrasch) * 100) / xrasch);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            int Table1_Asred = 0;
            label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";
            double x2 = 0;
            double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double x1_1 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
            double y1_1 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (y1 - x0 - k1 * x1) * (y1 - x0 - k1 * x1);
                yx1 += (y1 - x0 - SREDSUMM) * (y1 - x0 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            for (int i = 1; i < Table1.Rows.Count - 1; i++)
            {
                y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY(x1, (y1 - y1_1));
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;


                //     chart1.Series[0].Points.Item.Label = Convert.ToString(circle);
                // chart1.Series[0].IsValueShownAsLabel = true;
                //chart1.Series[0].IsXValueIndexed = true;
                // circle++;
                // double x2 = 0.1 * i;
                // double y2 = x2 / k1;
                x2 = 0;
                if (Table1_Asred == 0)
                {
                    x2 = 0;
                }
                else
                {
                    x2 = x1 - x1_1;
                }
                Table1_Asred++;
                double y2 = x2 * k1;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2;
                //    chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);
                //   chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        public void USE_KO_lineinaya0_not1()
        {
            double max = -1;
            int index = -1;
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            max = -1;
            double[] Table1masStr_1 = new double[NoCaIzm];
            double[] Table1masStr1_1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr_1[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr_1);
                double maxEl = Table1masStr_1[Table1masStr_1.Length - 1];
                double minEl = Table1masStr_1[0];
                double p1 = 2 * ((maxEl - minEl) / (maxEl + minEl)) * 100;
                //  Table1.Rows[i].Cells["P"].Value = string.Format("{0:0.0000}", p1);
                Table1masStr1_1[i] = p1;
            }
            max = -1;
            for (int i = 1; i <= Table1masStr1_1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1_1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1_1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label21.Text = "P,% = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                XY += x * y;
                SUMMY2 += y * y;
                Table1.Rows[i].Cells["X*X"].Value = y * y;
                Table1.Rows[i].Cells["X*Y"].Value = x * y;
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
            }
            k1 = XY / SUMMY2;

            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A";
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) * k1;
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k1;

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index+1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(A) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //  double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                //double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k1;
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - x) * 100) / x);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label22.Text = "Макс.Ошибка C(A) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";


            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (x1 - k1 * y1) * (x1 - k1 * y1);
                yx1 += (x1 - SREDSUMM) * (x1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2 = 0;
            int Table1_Asred = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY(x1, y1);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;


                //     chart1.Series[0].Points.Item.Label = Convert.ToString(circle);
                // chart1.Series[0].IsValueShownAsLabel = true;
                //chart1.Series[0].IsXValueIndexed = true;
                // circle++;
                //double y2 = 0.5 * i;
                //double x2 = y2 / k1;

                x2 = 0;
                if (Table1_Asred == 0)
                {
                    x2 = 0;
                }
                else
                {
                    x2 = x1;
                }
                Table1_Asred++;
                double y2 = x2 * k1;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //   chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2), 2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                // chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                // chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)), 2);
                //    chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        public void USE_KO_lineinaya01()
        {
            double max = -1;
            int index = -1;
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            max = -1;
            double[] Table1masStr_1 = new double[NoCaIzm];
            double[] Table1masStr1_1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr_1[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr_1);
                double maxEl = Table1masStr_1[Table1masStr_1.Length - 1];
                double minEl = Table1masStr_1[0];
                double p1 = 2 * ((maxEl - minEl) / (maxEl + minEl)) * 100;
                //  Table1.Rows[i].Cells["P"].Value = string.Format("{0:0.0000}", p1);
                Table1masStr1_1[i] = p1;
            }
            max = -1;
            for (int i = 1; i <= Table1masStr1_1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1_1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1_1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label21.Text = "P,% = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                XY += x * (y - y0);
                SUMMY2 += (y - y0) * (y - y0);
                Table1.Rows[i].Cells["X*X"].Value = (y - y0) * (y - y0);
                Table1.Rows[i].Cells["X*Y"].Value = x * (y - y0);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "";
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
            }
            k1 = XY / SUMMY2;

            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A";
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = (Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value)) * k1;
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k1;

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(C) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(C) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //  double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                //double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k1;
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - x) * 100) / x);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            int Table1_Asred = 0;
            Table1_Asred = 0;
            double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (x1 - k1 * (y1 - x0)) * (x1 - k1 * (y1 - x0));
                yx1 += (x1 - SREDSUMM) * (x1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2 = 0;
            for (int i = 1; i < Table1.Rows.Count - 1; i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY((x1 - x0), y1);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;


                //     chart1.Series[0].Points.Item.Label = Convert.ToString(circle);
                // chart1.Series[0].IsValueShownAsLabel = true;
                //chart1.Series[0].IsXValueIndexed = true;
                // circle++;
                //double y2 = 0.5 * i;
                //double x2 = y2 / k1;
                x2 = 0;
                if (Table1_Asred == 0)
                {
                    x2 = 0;
                }
                else
                {
                    x2 = x1 - x0;
                }
                Table1_Asred++;

                double y2 = x2 * k1;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //   chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2), 2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                // chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                // chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)), 2);
                //    chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1;
            chart1.Series[1].Points.AddXY(xfin, yfin);

        }

        public void lineinaya()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double SUMMX1 = 0;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            SUM0 = 0; SUM1 = 0;
            SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", 0);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            max = -1;
            double[] Table1masStr_1 = new double[NoCaIzm];
            double[] Table1masStr1_1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr_1[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr_1);
                double maxEl = Table1masStr_1[Table1masStr_1.Length - 1];
                double minEl = Table1masStr_1[0];
                double p1 = 2 * ((maxEl - minEl) / (maxEl + minEl)) * 100;
                //  Table1.Rows[i].Cells["P"].Value = string.Format("{0:0.0000}", p1);
                Table1masStr1_1[i] = p1;
            }
            for (int i = 1; i <= Table1masStr1_1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1_1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1_1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //  index = index + 1;
            label21.Text = "P,% = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            if (radioButton4.Checked == true)
            {

                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                }



                catch
                {
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                }
                if (USE_KO == false)
                {
                    USE_KO_lineinaya_not();
                }
                else
                {
                    USE_KO_lineinaya();
                }
            }
            else
            {

                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Асред* Асред");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                }


                if (USE_KO == false)
                {
                    USE_KO_lineinaya1_not();
                }
                else
                {
                    USE_KO_lineinaya1();
                }
            }




        }
        public void USE_KO_lineinaya_not()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double SUMMX1 = 0;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();

            k0 = 0; k1 = 0; k2 = 0;
            SUM0 = 0; SUM1 = 0;
            SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                // double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value);
                // double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                double x1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value);
                SUMMX += x; SUMMY += y;
                XY += x * y;
                SUMMY2 += y * y;
                Table1.Rows[i].Cells["X*X"].Value = y * y;
                Table1.Rows[i].Cells["X*Y"].Value = x * y;
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
            }

            SREDSUMMX = SUMMX / (Table1.Rows.Count - 1);
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index + 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(А) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }
            k0 = (SUMMY2 * SUMMX - SUMMY * XY) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            k1 = ((NoCaSer) * XY - SUMMY * SUMMX) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", 0);
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double xrasch = k1 * x + k0;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - xrasch) * 100) / xrasch);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index + 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double y2 = 0;
            label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C " + k0.ToString("+ 0.0000 ;- 0.0000 ");

            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (y1 - (k1 * x1 + k0)) * (y1 - (k1 * x1 + k0));
                yx1 += (y1 - SREDSUMM) * (y1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));

            double x0 = 0;
            double y0 = x0 * k1 + k0;
            chart1.Series[1].Points.AddXY(x0, y0);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                circle = 1;
                double x1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);



                // chart1.ChartAreas[0].AxisY.Crossing = k0;
                chart1.Series[0].Points.AddXY(y1_1, x1_1);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;
                //  double x2 = 0.1 * i;
                //double y2 = (x2 - k0) / k1;
                y2 = y1_1;
                double x2 = y1_1 * k1 + k0;
                chart1.Series[1].Points.AddXY(y2, x2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(string.Format("{0:0.0000}", Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(string.Format("{0:0.0000}", Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2);
                //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);

                //       chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = y2 * 1.1;
            double yfin = xfin * k1 + k0;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        public void USE_KO_lineinaya()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double SUMMX1 = 0;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            SUM0 = 0; SUM1 = 0;
            SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", 0);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            max = -1;
            double x1_1 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double y1_1 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
            SUMMX += (x1_1 - x1_1); SUMMY += y1_1;
            SUMMX1 += x1_1;
            XY += (x1_1 - x1_1) * y1_1;
            SUMMY2 += y1_1 * y1_1;
            Table1.Rows[0].Cells["X*X"].Value = y1_1 * y1_1;
            Table1.Rows[0].Cells["X*Y"].Value = (x1_1 - x1_1) * y1_1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                SUMMX1 += x;
                SUMMX += (x - x1_1); SUMMY += y;
                XY += (x - x1_1) * y;
                SUMMY2 += y * y;
                Table1.Rows[i].Cells["X*X"].Value = y * y;
                Table1.Rows[i].Cells["X*Y"].Value = (x - x1_1) * y;
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
            }
            SREDSUMMX = SUMMX1 / (Table1.Rows.Count - 1);
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value);

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(А) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }
            k0 = (SUMMY2 * SUMMX - SUMMY * XY) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            k1 = ((NoCaSer) * XY - SUMMY * SUMMX) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", 0);
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double xrasch = k1 * x + k0;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - xrasch) * 100) / xrasch);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C " + k0.ToString("+ 0.0000 ;- 0.0000 ");
            double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (y1 - x0 - (k1 * x1 + k2)) * (y1 - x0 - (k1 * x1 + k2));
                yx1 += (y1 - x0 - SREDSUMM) * (y1 - x0 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2 = x0 - x0;
            double y2 = x2 * k1 + k0;
            chart1.Series[1].Points.AddXY(x2, y2);
            for (int i = 1; i < Table1.Rows.Count - 1; i++)
            {
                circle = 1;
                x1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                y1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                // chart1.ChartAreas[0].AxisY.Crossing = k0;
                chart1.Series[0].Points.AddXY(y1_1, (x1_1 - x0));
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;
                //  double x2 = 0.1 * i;
                //double y2 = (x2 - k0) / k1;
                x2 = y1_1;
                y2 = x2 * k1 + k0;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(string.Format("{0:0.0000}", Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(string.Format("{0:0.0000}", Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2);
                //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);

                //       chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1 + k0;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        public void USE_KO_lineinaya1()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double SUMMX1 = 0;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            SUM0 = 0; SUM1 = 0;
            SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", 0);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            max = -1;
            double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
            double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            SUMMX += x0; SUMMY += y0 - y0;
            XY += x0 * (y0 - y0);
            SUMMY2 += (y0 - y0) * (y0 - y0);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);


                SUMMX += x; SUMMY += (y - y0);
                XY += x * (y - y0);
                SUMMY2 += (y - y0) * (y - y0);
                Table1.Rows[i].Cells["X*X"].Value = (y - y0) * (y - y0);
                Table1.Rows[i].Cells["X*Y"].Value = x * (y - y0);
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
            }
            k0 = (SUMMY2 * SUMMX - SUMMY * XY) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            k1 = ((NoCaSer) * XY - SUMMY * SUMMX) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", 0);
            label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A " + k0.ToString("+ 0.0000 ;- 0.0000 ");
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = (Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value)) * k1 + k0;
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k1 + k0;

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(C) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //  double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                //double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k1 + k0;
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - x) * 100) / x);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";

            x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (x1 - (k1 * (y1 - x0) + k0)) * (x1 - (k1 * (y1 - x0) + k0));
                yx1 += (x1 - SREDSUMM) * (x1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2 = x0 - x0;
            double y2 = x2 * k1 + k0;
            chart1.Series[1].Points.AddXY(x2, y2);
            for (int i = 1; i < Table1.Rows.Count - 1; i++)
            {
                circle = 1;
                double x1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                // chart1.ChartAreas[0].AxisY.Crossing = k0;
                chart1.Series[0].Points.AddXY((x1_1 - x0), y1_1);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;
                // double y2 = 0.5 * i;
                //     double x2 = (y2 - k0) / k1;
                //  double y2 = k1 * x1_1 + k0;
                x2 = x1_1 - x0;
                y2 = x2 * k1 + k0;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2), 2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //      chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)), 2);
                //     chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1 + k0;
            chart1.Series[1].Points.AddXY(xfin, yfin);

        }
        public void USE_KO_lineinaya1_not()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            double SUMMSer = 0;
            double SREDSUMMX = 0;
            double SUMMX1 = 0;

            AgroText0.Text = string.Format("{0:0.0000}", 0);
            AgroText1.Text = string.Format("{0:0.0000}", 0);
            AgroText0.Text = string.Format("{0:0.0000}", 0);
            max = -1;
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            k0 = 0; k1 = 0; k2 = 0;
            SUM0 = 0; SUM1 = 0;
            SUMMX = 0; SUMMY = 0; XY = 0; SUMMY2 = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
                double x1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value);
                // double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value);
                SUMMX += x; SUMMY += y;
                XY += x * y;
                SUMMY2 += y * y;
                Table1.Rows[i].Cells["X*X"].Value = y * y;
                Table1.Rows[i].Cells["X*Y"].Value = x * y;
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(SUMMY2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(XY);
            }
            k0 = (SUMMY2 * SUMMX - SUMMY * XY) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            k1 = ((NoCaSer) * XY - SUMMY * SUMMX) / ((NoCaSer) * SUMMY2 - SUMMY * SUMMY);
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", 0);
            label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A " + k0.ToString("+ 0.0000 ;- 0.0000 ");
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) * k1 + k0;
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k1 + k0;

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(C) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //  double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                //double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k1 + k0;
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - x) * 100) / x);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index + 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (x1 - (k1 * y1 + k0)) * (x1 - (k1 * y1 + k0));
                yx1 += (x1 - SREDSUMM) * (x1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x0 = 0;
            double y0 = x0 * k1 + k0;
            double x2 = 0;
            chart1.Series[1].Points.AddXY(x0, y0);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                circle = 1;
                double x1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y1_1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                // chart1.ChartAreas[0].AxisY.Crossing = k0;
                chart1.Series[0].Points.AddXY(x1_1, y1_1);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;
                // double y2 = 0.5 * i;
                //     double x2 = (y2 - k0) / k1;
                //  double y2 = k1 * x1_1 + k0;
                x2 = x1_1;
                double y2 = x1_1 * k1 + k0;
                chart1.Series[1].Points.AddXY(x2, y2);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2), 2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //      chart1.ChartAreas[0].AxisY.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)), 2);
                //     chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2 * 1.1;
            double yfin = xfin * k1 + k0;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            Zavisimoct = "A(C)";
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();


            if (radioButton1.Checked == true)
            {

                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                lineinaya0();
            }
            else
            {
                if (radioButton2.Checked == true)
                {
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    lineinaya();
                }
                else
                {
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    kvadratichnaya();
                }
            }
        }

        public void kvadratichnaya()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            double x2 = 0; double x3 = 0; double x4 = 0; double xy = 0; double SUMX = 0;
            double SUMY = 0; double x2y = 0;
            double Opred; double OpredA; double OpredB; double OpredC;
            k0 = 0; k1 = 0; k2 = 0;
            max = -1;
            double[] Table1masStr_1 = new double[NoCaIzm];
            double[] Table1masStr1_1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr_1[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr_1);
                double maxEl = Table1masStr_1[Table1masStr_1.Length - 1];
                double minEl = Table1masStr_1[0];
                double p1 = 2 * ((maxEl - minEl) / (maxEl + minEl)) * 100;
                //  Table1.Rows[i].Cells["P"].Value = string.Format("{0:0.0000}", p1);
                Table1masStr1_1[i] = p1;
            }
            for (int i = 1; i <= Table1masStr1_1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1_1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1_1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index + 1;
            label21.Text = "P,% = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            if (radioButton4.Checked == true)
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Асред* Конц");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                    Table1.Columns["X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*Y"].ReadOnly = true;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Конц* Конц");
                    Table1.Columns.Add("X*Y", "Асред* Конц");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                    Table1.Columns["X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*Y"].ReadOnly = true;
                }
                if (USE_KO == false)
                {
                    USE_KO_kvadratichnaya_not();
                }
                else
                {
                    USE_KO_kvadratichnaya();
                }
            }
            else
            {
                try
                {
                    Table1.Columns.Remove("X*X");
                    Table1.Columns.Remove("X*Y");
                    Table1.Columns.Remove("X*X*X");
                    Table1.Columns.Remove("X*X*X*X");
                    Table1.Columns.Remove("X*X*Y");
                    Table1.Columns.Add("X*X", "Асред ^2");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");
                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                    Table1.Columns["X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*Y"].ReadOnly = true;
                }
                catch
                {
                    Table1.Columns.Add("X*X", "Асред ^2");
                    Table1.Columns.Add("X*Y", "Конц* Асред");
                    Table1.Columns.Add("X*X*X", "Асред ^3");
                    Table1.Columns.Add("X*X*X*X", "Асред ^4");
                    Table1.Columns.Add("X*X*Y", "Асред ^2*Конц");

                    Table1.Columns["X*X"].Width = 50;
                    Table1.Columns["X*Y"].Width = 50;
                    Table1.Columns["X*X*X"].Width = 50;
                    Table1.Columns["X*X*X*X"].Width = 50;
                    Table1.Columns["X*X*Y"].Width = 50;
                    Table1.Columns["X*X"].ReadOnly = true;
                    Table1.Columns["X*Y"].ReadOnly = true;
                    Table1.Columns["X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*X*X"].ReadOnly = true;
                    Table1.Columns["X*X*Y"].ReadOnly = true;
                }
                if (USE_KO == false)
                {
                    USE_KO_kvadratichnaya1_not();
                }
                else
                {
                    USE_KO_kvadratichnaya1();
                }

            }
        }
        public void USE_KO_kvadratichnaya_not()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            double x2 = 0; double x3 = 0; double x4 = 0; double xy = 0; double SUMX = 0;
            double SUMY = 0; double x2y = 0;
            double Opred; double OpredA; double OpredB; double OpredC;
            k0 = 0; k1 = 0; k2 = 0;
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {

                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                x2 += x * x;
                x3 += x * x * x;
                x4 += x * x * x * x;
                xy += x * y;
                SUMX += x;
                SUMY += y;
                x2y += x * x * y;
                Table1.Rows[i].Cells["X*X"].Value = string.Format("{0:0.0000}", x * x);
                Table1.Rows[i].Cells["X*Y"].Value = string.Format("{0:0.0000}", x * y);
                Table1.Rows[i].Cells["X*X*X"].Value = string.Format("{0:0.0000}", x * x * x);
                Table1.Rows[i].Cells["X*X*X*X"].Value = string.Format("{0:0.0000}", x * x * x * x);
                Table1.Rows[i].Cells["X*X*Y"].Value = string.Format("{0:0.0000}", x * x * y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(x2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X"].Value = "СУММА = " + Convert.ToString(x3);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X*X"].Value = "СУММА = " + Convert.ToString(x4);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*Y"].Value = "СУММА = " + Convert.ToString(x2y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(xy);

            }
            double SUMMSer = 0;
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index + 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(А) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }
            Opred = x2 * x2 * x2 + SUMX * SUMX * x4 + (NoCaSer) * x3 * x3 - (NoCaSer) * x2 * x4 - x2 * SUMX * x3 - SUMX * x3 * x2;
            OpredA = SUMY * x2 * x2 + SUMX * SUMX * x2y + (NoCaSer) * xy * x3 - (NoCaSer) * x2 * x2y - SUMY * SUMX * x3 - SUMX * xy * x2;
            OpredB = x2 * xy * x2 + SUMY * SUMX * x4 + (NoCaSer) * x3 * x2y - (NoCaSer) * xy * x4 - x2 * SUMX * x2y - SUMY * x3 * x2;
            OpredC = x2 * x2 * x2y + SUMX * xy * x4 + SUMY * x3 * x3 - SUMY * x2 * x4 - x2 * xy * x3 - SUMX * x3 * x2y;

            k2 = OpredA / Opred;
            k1 = OpredB / Opred;
            k0 = OpredC / Opred;
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", k2);
            label14.Text = "A(C) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000 ;- 0.0000 ") + "*C " + k2.ToString("+ 0.0000 ;- 0.0000 ") + "*C^2";
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double xrasch = k1 * x + k2 * x * x + k0;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - xrasch) * 100) / xrasch);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (y1 - (k1 * x1 + k2 * x1 * x1 + k0)) * (y1 - (k1 * x1 + k2 * x1 * x1 + k0));
                yx1 += (y1 - SREDSUMM) * (y1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));

            double x2_1 = 0;
            double y0 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;
            chart1.Series[1].Points.AddXY(x2_1, y0);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY(x, y);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;

                // double x2_1 = 0.3 * i;
                x2_1 = x;
                double y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

                chart1.Series[1].Points.AddXY(x2_1, y2_1);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + x2_1;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + y2_1), 2);
                //chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);
            }
            double xfin = x2_1 * 1.1;
            double yfin = k0 + k1 * xfin + k2 * xfin * xfin;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }

        public void USE_KO_kvadratichnaya()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            double x2 = 0; double x3 = 0; double x4 = 0; double xy = 0; double SUMX = 0;
            double SUMY = 0; double x2y = 0;
            double Opred; double OpredA; double OpredB; double OpredC;
            k0 = 0; k1 = 0; k2 = 0;
            double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {

                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                x2 += x * x;
                x3 += x * x * x;
                x4 += x * x * x * x;
                xy += x * (y - y0);
                SUMX += x;
                SUMY += (y - y0);
                x2y += x * x * (y - y0);
                Table1.Rows[i].Cells["X*X"].Value = string.Format("{0:0.0000}", x * x);
                Table1.Rows[i].Cells["X*Y"].Value = string.Format("{0:0.0000}", x * (y - y0));
                Table1.Rows[i].Cells["X*X*X"].Value = string.Format("{0:0.0000}", x * x * x);
                Table1.Rows[i].Cells["X*X*X*X"].Value = string.Format("{0:0.0000}", x * x * x * x);
                Table1.Rows[i].Cells["X*X*Y"].Value = string.Format("{0:0.0000}", x * x * (y - y0));
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(x2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X"].Value = "СУММА = " + Convert.ToString(x3);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X*X"].Value = "СУММА = " + Convert.ToString(x4);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*Y"].Value = "СУММА = " + Convert.ToString(x2y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(xy);

            }
            double SUMMSer = 0;
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
                SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value);

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(А) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(A) - Не применимо для Nсер. < 3";
            }
            Opred = x2 * x2 * x2 + SUMX * SUMX * x4 + (NoCaSer) * x3 * x3 - (NoCaSer) * x2 * x4 - x2 * SUMX * x3 - SUMX * x3 * x2;
            OpredA = SUMY * x2 * x2 + SUMX * SUMX * x2y + (NoCaSer) * xy * x3 - (NoCaSer) * x2 * x2y - SUMY * SUMX * x3 - SUMX * xy * x2;
            OpredB = x2 * xy * x2 + SUMY * SUMX * x4 + (NoCaSer) * x3 * x2y - (NoCaSer) * xy * x4 - x2 * SUMX * x2y - SUMY * x3 * x2;
            OpredC = x2 * x2 * x2y + SUMX * xy * x4 + SUMY * x3 * x3 - SUMY * x2 * x4 - x2 * xy * x3 - SUMX * x3 * x2y;

            k2 = OpredA / Opred;
            k1 = OpredB / Opred;
            k0 = OpredC / Opred;
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", k2);
            label14.Text = "A(C) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000;- 0.0000") + "*C " + k2.ToString("+ 0.0000;- 0.0000") + "*C^2";
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double xrasch = k1 * x + k2 * x * x + k0;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value);
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - xrasch) * 100) / xrasch);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            y0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (y1 - y0 - (k1 * x1 + k2 * x1 * x1 + k0)) * (y1 - y0 - (k1 * x1 + k2 * x1 * x1 + k0));
                yx1 += (y1 - y0 - SREDSUMM) * (y1 - y0 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2_1 = x0;
            double y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

            chart1.Series[1].Points.AddXY(x2_1, y2_1);
            for (int i = 1; i < Table1.Rows.Count - 1; i++)
            {
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY(x, (y - y0));
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;

                // double x2_1 = 0.3 * i;
                x2_1 = x;
                y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

                chart1.Series[1].Points.AddXY(x2_1, y2_1);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + x2_1;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + y2_1), 2);
                //chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Concetr"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Concetr"].Value)), 2);
            }
            double xfin = x2_1 * 1.1;
            double yfin = k0 + k1 * xfin + k2 * xfin * xfin;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        public void USE_KO_kvadratichnaya1_not()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            double x2 = 0; double x3 = 0; double x4 = 0; double xy = 0; double SUMX = 0;
            double SUMY = 0; double x2y = 0;
            double Opred; double OpredA; double OpredB; double OpredC;
            k0 = 0; k1 = 0; k2 = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                x2 += x * x;
                x3 += x * x * x;
                x4 += x * x * x * x;
                xy += x * y;
                SUMX += x;
                SUMY += y;
                x2y += x * x * y;
                Table1.Rows[i].Cells["X*X"].Value = string.Format("{0:0.0000}", x * x);
                Table1.Rows[i].Cells["X*X*X"].Value = string.Format("{0:0.0000}", x * x * x);
                Table1.Rows[i].Cells["X*X*X*X"].Value = string.Format("{0:0.0000}", x * x * x * x);
                Table1.Rows[i].Cells["X*X*Y"].Value = string.Format("{0:0.0000}", x * x * y);
                Table1.Rows[i].Cells["X*Y"].Value = string.Format("{0:0.0000}", x * y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(x2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X"].Value = "СУММА = " + Convert.ToString(x3);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X*X"].Value = "СУММА = " + Convert.ToString(x4);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*Y"].Value = "СУММА = " + Convert.ToString(x2y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(xy);
            }
            Opred = x2 * x2 * x2 + SUMX * SUMX * x4 + (NoCaSer) * x3 * x3 - (NoCaSer) * x2 * x4 - x2 * SUMX * x3 - SUMX * x3 * x2;
            OpredA = SUMY * x2 * x2 + SUMX * SUMX * x2y + (NoCaSer) * xy * x3 - (NoCaSer) * x2 * x2y - SUMY * SUMX * x3 - SUMX * xy * x2;
            OpredB = x2 * xy * x2 + SUMY * SUMX * x4 + (NoCaSer) * x3 * x2y - (NoCaSer) * xy * x4 - x2 * SUMX * x2y - SUMY * x3 * x2;
            OpredC = x2 * x2 * x2y + SUMX * xy * x4 + SUMY * x3 * x3 - SUMY * x2 * x4 - x2 * xy * x3 - SUMX * x3 * x2y;

            k2 = OpredA / Opred;
            k1 = OpredB / Opred;
            k0 = OpredC / Opred;
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", k2);
            label14.Text = "C(A) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000;- 0.0000") + "*A " + k2.ToString("+ 0.0000;- 0.0000") + "*A^2";
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) * k1 + Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) * Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) * k2 + k0;
                double SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k1 + Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k2 + k0;

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            // index = index + 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(C) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(C) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //  double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                //double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k1 + Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) * k2 + k0;
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - x) * 100) / x);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            //index = index + 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (x1 - (k1 * y1 + k2 * y1 * y1 + k0)) * (x1 - (k1 * y1 + k2 * y1 * y1 + k0));
                yx1 += (x1 - SREDSUMM) * (x1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2_1 = 0;
            double y0 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;
            chart1.Series[1].Points.AddXY(x2_1, y0);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY(x, y);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;
                x2_1 = x;
                double y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

                chart1.Series[1].Points.AddXY(x2_1, y2_1);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2_1), 2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2_1), 2);
                //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)+ (Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value))), 2);
                //  chart1.ChartAreas[0].AxisX.Interval = 5;
            }
        }
        public void USE_KO_kvadratichnaya1()
        {
            double max = -1;
            int index = -1;
            double[] SredOtklMatr = new double[Table1.Rows.Count - 1];
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            double x2 = 0; double x3 = 0; double x4 = 0; double xy = 0; double SUMX = 0;
            double SUMY = 0; double x2y = 0;
            double Opred; double OpredA; double OpredB; double OpredC;
            k0 = 0; k1 = 0; k2 = 0;
            double x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                x2 += (x - x0) * (x - x0);
                x3 += (x - x0) * (x - x0) * (x - x0);
                x4 += (x - x0) * (x - x0) * (x - x0) * (x - x0);
                xy += (x - x0) * y;
                SUMX += (x - x0);
                SUMY += y;
                x2y += (x - x0) * (x - x0) * y;
                Table1.Rows[i].Cells["X*X"].Value = string.Format("{0:0.0000}", (x - x0) * (x - x0));
                Table1.Rows[i].Cells["X*X*X"].Value = string.Format("{0:0.0000}", (x - x0) * (x - x0) * (x - x0));
                Table1.Rows[i].Cells["X*X*X*X"].Value = string.Format("{0:0.0000}", (x - x0) * (x - x0) * (x - x0) * (x - x0));
                Table1.Rows[i].Cells["X*X*Y"].Value = string.Format("{0:0.0000}", (x - x0) * (x - x0) * y);
                Table1.Rows[i].Cells["X*Y"].Value = string.Format("{0:0.0000}", (x - x0) * y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["NoCo"].Value = "n = " + Convert.ToString(Table1.Rows.Count - 1);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Asred"].Value = "СУММА = " + Convert.ToString(SUMMX);
                Table1.Rows[Table1.Rows.Count - 1].Cells["Concetr"].Value = "СУММА = " + Convert.ToString(SUMMY);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X"].Value = "СУММА = " + Convert.ToString(x2);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X"].Value = "СУММА = " + Convert.ToString(x3);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*X*X"].Value = "СУММА = " + Convert.ToString(x4);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*X*Y"].Value = "СУММА = " + Convert.ToString(x2y);
                Table1.Rows[Table1.Rows.Count - 1].Cells["X*Y"].Value = "СУММА = " + Convert.ToString(xy);
            }
            Opred = x2 * x2 * x2 + SUMX * SUMX * x4 + (NoCaSer) * x3 * x3 - (NoCaSer) * x2 * x4 - x2 * SUMX * x3 - SUMX * x3 * x2;
            OpredA = SUMY * x2 * x2 + SUMX * SUMX * x2y + (NoCaSer) * xy * x3 - (NoCaSer) * x2 * x2y - SUMY * SUMX * x3 - SUMX * xy * x2;
            OpredB = x2 * xy * x2 + SUMY * SUMX * x4 + (NoCaSer) * x3 * x2y - (NoCaSer) * xy * x4 - x2 * SUMX * x2y - SUMY * x3 * x2;
            OpredC = x2 * x2 * x2y + SUMX * xy * x4 + SUMY * x3 * x3 - SUMY * x2 * x4 - x2 * xy * x3 - SUMX * x3 * x2y;

            k2 = OpredA / Opred;
            k1 = OpredB / Opred;
            k0 = OpredC / Opred;
            AgroText0.Text = string.Format("{0:0.0000}", k0);
            AgroText1.Text = string.Format("{0:0.0000}", k1);
            AgroText2.Text = string.Format("{0:0.0000}", k2);
            label14.Text = "C(A) = " + k0.ToString("0.0000 ;- 0.0000 ") + k1.ToString("+ 0.0000;- 0.0000") + "*A " + k2.ToString("+ 0.0000;- 0.0000") + "*A^2";
            max = -1;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double Ser1 = (Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value)) * k1 + (Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value)) * (Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value)) * k2 + k0;
                double SUMMSer = 0;
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    double Ser = (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k1 + (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k2 + k0;

                    SUMMSer += (Ser - Ser1) * (Ser - Ser1);
                }
                double SredOtkl = Math.Sqrt(SUMMSer / (NoCaIzm - 1));
                double SredOtklProc = (SredOtkl / Ser1) * 100;
                SredOtklMatr[i] = SredOtklProc;
            }

            // Цикл по всем элементам массива
            // От 0 до размера массива
            for (int i = 1; i <= SredOtklMatr.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= SredOtklMatr[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = SredOtklMatr[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            if (NoCaIzm >= 3)
            {
                SKO.Text = "СКО(C) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            }
            else
            {
                SKO.Text = "СКО(C) - Не применимо для Nсер. < 3";
            }
            max = -1;
            double[] Table1masStr1 = new double[Table1.Rows.Count - 1];
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                //  double y = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                //double xrasch = k1 * x;
                double[] Table1masStr = new double[NoCaIzm];
                for (int j = 1; j <= NoCaIzm; j++)
                {
                    Table1masStr[j - 1] = (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k1 + (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * (Convert.ToDouble(Table1.Rows[i].Cells["A;Ser (" + j].Value) - Convert.ToDouble(Table1.Rows[0].Cells["A;Ser (" + j].Value)) * k2 + k0;
                }
                Array.Sort(Table1masStr);
                double maxEl = Table1masStr[Table1masStr.Length - 1];
                Table1masStr1[i] = Math.Abs(((maxEl - x) * 100) / x);
                //label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.0000}", (((maxEl - xrasch) * 100) / xrasch));
            }
            for (int i = 1; i <= Table1masStr1.Length; i++)
            {
                // Если максимальная стоимость меньше, либо равно текущей проверяемой
                if (max <= Table1masStr1[i - 1])
                {
                    // Запоминаем новое максимальное значение
                    max = Table1masStr1[i - 1];
                    // Запоминаем порядковый номер
                    index = i;
                }
            }
            // max = max / 100;
            index = index - 1;
            label22.Text = "Макс.Ошибка А(С) = " + string.Format("{0:0.00}", max) + "% (CO №" + index + ")";
            double y0 = Convert.ToDouble(Table1.Rows[0].Cells["Concetr"].Value);
            x0 = Convert.ToDouble(Table1.Rows[0].Cells["Asred"].Value);
            double yx = 0;
            double yx1 = 0;
            double SREDSUMM = 0;
            SUMMX = 0;
            for (int i = 0; i < Table1.Rows.Count - 1; i++)
            {
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                SUMMX += y1;
            }
            SREDSUMM = SUMMX / (Table1.Rows.Count - 1);
            for (int i = 0; i < (Table1.Rows.Count - 1); i++)
            {
                double x1 = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);
                double y1 = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);

                yx += (x1 - (k1 * (y1 - x0) + k2 * (y1 - x0) * (y1 - x0) + k0)) * (x1 - (k1 * (y1 - x0) + k2 * (y1 - x0) * (y1 - x0) + k0));
                yx1 += (x1 - SREDSUMM) * (x1 - SREDSUMM);
            }
            RR.Text = "R^2 = " + string.Format("{0:0.0000}", (1 - (yx / yx1)));
            double x2_1 = x0 - x0;
            double y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

            chart1.Series[1].Points.AddXY(x2_1, y2_1);
            for (int i = 1; i < Table1.Rows.Count - 1; i++)
            {
                double x = Convert.ToDouble(Table1.Rows[i].Cells["Asred"].Value);
                double y = Convert.ToDouble(Table1.Rows[i].Cells["Concetr"].Value);

                chart1.Series[0].Points.AddXY((x - x0), y);
                chart1.Series[0].ChartType = SeriesChartType.Point;
                chart1.ChartAreas[0].AxisY.Crossing = 0;
                chart1.ChartAreas[0].AxisX.Crossing = 0;
                x2_1 = x - x0;
                y2_1 = k0 + k1 * x2_1 + k2 * x2_1 * x2_1;

                chart1.Series[1].Points.AddXY(x2_1, y2_1);
                chart1.Series[1].ChartType = SeriesChartType.Line;
                chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + edconctr;
                chart1.ChartAreas[0].AxisX.Minimum = 0;
                //  chart1.ChartAreas[0].AxisX.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Asred"].Value) + x2_1), 2);
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                //   chart1.ChartAreas[0].AxisY.Maximum = Math.Round((Convert.ToDouble(Table1.Rows[Table1.Rows.Count - 2].Cells["Concetr"].Value) + y2_1), 2);
                //   chart1.ChartAreas[0].AxisX.Interval = Math.Round((Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value)+ (Convert.ToDouble(Table1.Rows[3].Cells["Asred"].Value) - Convert.ToDouble(Table1.Rows[2].Cells["Asred"].Value))), 2);
                //  chart1.ChartAreas[0].AxisX.Interval = 5;
            }
            double xfin = x2_1 * 1.1;
            double yfin = k0 + k1 * xfin + k2 * xfin * xfin;
            chart1.Series[1].Points.AddXY(xfin, yfin);
        }
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            Zavisimoct = "C(A)";
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();


            if (radioButton1.Checked == true)
            {

                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                lineinaya0();
            }
            else
            {
                if (radioButton2.Checked == true)
                {
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    lineinaya();
                }
                else
                {
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    kvadratichnaya();
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            aproksim = "Линейная через 0";
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            lineinaya0();
            AgroText0.Enabled = true;
            AgroText1.Enabled = true;
            AgroText2.Enabled = true;
            RR.Enabled = true;
            SKO.Enabled = true;
            label21.Enabled = true;
            label22.Enabled = true;
            label14.Enabled = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            aproksim = "Линейная";
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            lineinaya();
            AgroText0.Enabled = true;
            AgroText1.Enabled = true;
            AgroText2.Enabled = true;
            RR.Enabled = true;
            SKO.Enabled = true;
            label21.Enabled = true;
            label22.Enabled = true;
            label14.Enabled = true;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            aproksim = "Квадратичная";
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            kvadratichnaya();
            AgroText0.Enabled = true;
            AgroText1.Enabled = true;
            AgroText2.Enabled = true;
            RR.Enabled = true;
            SKO.Enabled = true;
            label21.Enabled = true;
            label22.Enabled = true;
            label14.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            switch (selet_rezim)
            {
                case 2:
                    if (tabControl2.SelectedIndex == 0)
                    {
                        if ((Table1.RowCount < 1) && SposobZadan == "По СО")
                        {
                            MessageBox.Show("Создайте Градуировку");

                        }
                        else
                        {
                            Save();
                        }
                    }
                    else
                    {
                        if (Table2.RowCount > 0)
                        {
                            Save1();
                        }
                        else
                        {
                            MessageBox.Show("Создайте Измерение");
                        }
                    }
                    break;
                case 1:
                    SaveFR();
                    break;
                case 6:
                    if (tabControl2.SelectedIndex == 0)
                    {
                        if ((Table1.RowCount < 1) && SposobZadan == "По СО")
                        {
                            MessageBox.Show("Создайте Градуировку");

                        }
                        else
                        {
                            Save();
                        }
                    }
                    else
                    {
                        if (Table2.RowCount > 0)
                        {
                            Save1();
                        }
                        else
                        {
                            MessageBox.Show("Создайте Измерение");
                        }
                    }
                    break;
                case 5:
                    if (ScanTable.RowCount < 2)
                    {
                        MessageBox.Show("Создайте измерение");
                    }
                    else
                    {
                        SaveScan();
                    }
                    break;
                case 4:
                    if(TableKinetica1.RowCount < 2)
                    {
                        MessageBox.Show("Создайте измерение");
                    }
                    else
                    {
                        SaveKin();
                    }
                    break;
            }
        }
        public void SaveKin()
        {
            bool doNotWrite = false;
            for (int j = 0; j < TableKinetica1.RowCount - 1; j++)
            {
                for (int i = 0; i < TableKinetica1.Rows[j].Cells.Count; i++)
                {
                    if (TableKinetica1.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;
                    }
                }
            }
            if (doNotWrite)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                SaveAsKinTable();
            }
        }
        public void SaveScan()
        {
            bool doNotWrite = false;
            for (int j = 0; j < ScanTable.RowCount - 1; j++)
            {
                for (int i = 0; i < ScanTable.Rows[j].Cells.Count; i++)
                {
                    if (ScanTable.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;
                    }
                }
            }
            if (doNotWrite)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                SaveAsScanTable();
            }
        }
        public void SaveFR()
        {
            bool doNotWrite = false;
            for (int j = 0; j < IzmerenieFR_Table.Rows.Count - 1; j++)
            {

                for (int i = 3; i < IzmerenieFR_Table.Rows[j].Cells.Count; i++)
                {
                    if (IzmerenieFR_Table.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                SaveAsIzmerenieFR();
            }
        }

        public void Save()
        {
            if (SposobZadan != "Ввод коэффициентов")
            {
                bool doNotWrite = false;
                for (int j = 0; j < Table1.Rows.Count - 1; j++)
                {

                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;
                            break;

                        }
                    }
                }
                if (doNotWrite == true)
                {
                    MessageBox.Show("Не вся поля таблицы были заполнены!");
                }
                else
                {
                    SaveAs1();


                }
            }
            else
            {
                SaveAs1();

            }

        }
        public void Save1()
        {
            bool doNotWrite = false;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {

                for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                {
                    if (Table2.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                SaveAs2();
            }
        }
        public void SaveAs1()
        {
            if (selet_rezim == 2)
            {
                saveFileDialog1.InitialDirectory = "C";
                saveFileDialog1.Title = "Save as XML File";
                saveFileDialog1.FileName = "";
                saveFileDialog1.Filter = "QS2 файл|*.qs2";
            }
            else
            {
                saveFileDialog1.InitialDirectory = "C";
                saveFileDialog1.Title = "Save as XML File";
                saveFileDialog1.FileName = "";
                saveFileDialog1.Filter = "Agro QS2 файл|*.aq2";
            }
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocument(ref filepath);
                WriteXml(ref filepath);
                button3.Enabled = true;
                button9.Enabled = true;
                печатьToolStripMenuItem1.Enabled = true;
                tabPage4.Parent = tabControl2;
                if (selet_rezim == 6)
                {
                    tabControl2.TabPages[1].Text = "Измерение Агро";
                }
                Podskazka.Text = "Перейдите в Измерения!";
                label27.Visible = false;
                label24.Visible = false;
                label25.Visible = false;
                label26.Visible = false;
                label28.Visible = false;
                label33.Visible = false;
            }
        }
        public void SaveAsKinTable()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as XML File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "KIN файл|*.KIN2";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocumentIzmerenieKin(ref filepath);
                WriteXmlKin(ref filepath);
                button3.Enabled = true;
                button9.Enabled = true;
                печатьToolStripMenuItem1.Enabled = true;

            }
        }
        public void SaveAsScanTable()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as XML File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "SCAN файл|*.SCAN2";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocumentIzmerenieScan(ref filepath);
                WriteXmlIzmerenieScan(ref filepath);
                button3.Enabled = true;
                button9.Enabled = true;
                печатьToolStripMenuItem1.Enabled = true;

            }
        }

        public void SaveAsIzmerenieFR()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as XML File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "ISFR файл|*.isfr2";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocumentIzmerenieFR(ref filepath);
                WriteXmlIzmerenieFR(ref filepath);
                button3.Enabled = true;
                button9.Enabled = true;
                печатьToolStripMenuItem1.Enabled = true;

            }
        }

        public void CreateXMLDocumentIzmerenieScan(ref string filepath)
        {
            filepath = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath, Encoding.UTF8);

            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        public void CreateXMLDocumentIzmerenieKin(ref string filepath)
        {
            filepath = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath, Encoding.UTF8);

            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        private void CreateXMLDocument(ref string filepath)
        {

            filepath = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath, Encoding.UTF8);

            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        public void WriteXmlKin(ref string filepath)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath, FileMode.Open);
            xd.Load(fs);
            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

           //string HeaderCells1 =;
            XmlNode TypeIzmer = xd.CreateElement("TypeIzmer");
            TypeIzmer.InnerText = TableKinetica1.Columns[1].HeaderText;
            Izmerenie.AppendChild(TypeIzmer);

            xd.DocumentElement.AppendChild(Izmerenie);

            XmlNode NumberIzmer = xd.CreateElement("NumberIzmer");

          
            for (int i = 0; i < TableKinetica1.RowCount-1; i++)
            {
                XmlNode Str = xd.CreateElement("Str");
                XmlAttribute attribute2 = xd.CreateAttribute("Nomer");
                attribute2.Value = Convert.ToString(i); // устанавливаем значение атрибута
                Str.Attributes.Append(attribute2); // добавляем атрибут
                NumberIzmer.AppendChild(Str);

                for (int j = 0; j < TableKinetica1.ColumnCount; j++)
                {
               //     HeaderCells1 = this.TableKinetica1.Columns[j].HeaderText;
                    XmlNode Cells1 = xd.CreateElement("Cells" + j);
                    XmlAttribute attribute3 = xd.CreateAttribute("TypeCell");
                    attribute3.Value = TableKinetica1.Columns[j].HeaderText; // устанавливаем значение атрибута
                    Cells1.Attributes.Append(attribute3); // добавляем атрибут
                    Cells1.InnerText = TableKinetica1.Rows[i].Cells[j].Value.ToString();
                    Str.AppendChild(Cells1);
                    //xd.DocumentElement.AppendChild(Cells1);
                }
                //   xd.DocumentElement.AppendChild(Str);
            }
            xd.DocumentElement.AppendChild(NumberIzmer);

            fs.Close();         // Закрываем поток  
            xd.Save(filepath); // Сохраняем файл  

        }
        public void WriteXmlIzmerenieScan(ref string filepath)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath, FileMode.Open);
            xd.Load(fs);
            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

            XmlNode CountIzmer = xd.CreateElement("CountIzmer");
            CountIzmer.InnerText = listBox1.Items.Count.ToString();
            Izmerenie.AppendChild(CountIzmer);

            HeaderCells = new string[1];
            HeaderCells[0] = this.ScanTable.Columns[1].HeaderText;
            XmlNode TypeIzmer = xd.CreateElement("TypeIzmer");
            TypeIzmer.InnerText = HeaderCells[0];
            Izmerenie.AppendChild(TypeIzmer);

            // countScan[countButtonClick - 1][i, k]
            xd.DocumentElement.AppendChild(Izmerenie);
            int countIzmer = 0;

            HeaderCells = new string[this.ScanTable.Columns.Count];
            while (countIzmer < listBox1.Items.Count)
            {
                XmlNode NumberIzmer = xd.CreateElement("NumberIzmer");
                XmlAttribute attribute1 = xd.CreateAttribute("Nomer");
                attribute1.Value = Convert.ToString(countIzmer); // устанавливаем значение атрибута
                NumberIzmer.Attributes.Append(attribute1); // добавляем атрибут

                XmlNode CountStr = xd.CreateElement("CountStr");
                CountStr.InnerText = countScan[countIzmer].GetLength(0).ToString();
                NumberIzmer.AppendChild(CountStr);

                int m = countScan[countIzmer].GetLength(0);
                int n = countScan[countIzmer].GetLength(1);
                for (int i = 0; i < m; i++)
                {
                    XmlNode Str = xd.CreateElement("Str");
                    XmlAttribute attribute2 = xd.CreateAttribute("Nomer");
                    attribute2.Value = Convert.ToString(i); // устанавливаем значение атрибута
                    Str.Attributes.Append(attribute2); // добавляем атрибут
                    NumberIzmer.AppendChild(Str);

                    for (int j = 0; j < n; j++)
                    {
                        HeaderCells[j] = this.ScanTable.Columns[j].HeaderText;
                        XmlNode Cells1 = xd.CreateElement("Cells" + j);
                        XmlAttribute attribute3 = xd.CreateAttribute("TypeCell");
                        attribute3.Value = HeaderCells[j]; // устанавливаем значение атрибута
                        Cells1.Attributes.Append(attribute3); // добавляем атрибут
                        Cells1.InnerText = countScan[countIzmer][i, j];
                        Str.AppendChild(Cells1);
                        //xd.DocumentElement.AppendChild(Cells1);
                    }
                    //   xd.DocumentElement.AppendChild(Str);
                }
                xd.DocumentElement.AppendChild(NumberIzmer);
                countIzmer++;
            }



            fs.Close();         // Закрываем поток  
            xd.Save(filepath); // Сохраняем файл  
        }
        public void WriteXml(ref string filepath)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath, FileMode.Open);
            xd.Load(fs);

            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

            XmlNode Version = xd.CreateElement("Version"); // Версия программы
            Version.InnerText = version; // и значение
            Izmerenie.AppendChild(Version); // и указываем кому принадлежит

            XmlNode Nazvanie = xd.CreateElement("Nazvanie"); // Название вещества
            Nazvanie.InnerText = "Расчет градуировочного графика"; // и значение
            Izmerenie.AppendChild(Nazvanie); // и указываем кому принадлежит

            XmlNode Veshestvo = xd.CreateElement("Veshestvo"); // Название вещества
            Veshestvo.InnerText = Veshestvo1; // и значение
            Izmerenie.AppendChild(Veshestvo); // и указываем кому принадлежит

            XmlNode wavelength = xd.CreateElement("wavelength"); // Длина волны
            wavelength.InnerText = wavelength1; // и значение
            Izmerenie.AppendChild(wavelength); // и указываем кому принадлежит

            XmlNode WidthCuvet1 = xd.CreateElement("WidthCuvet"); // Ширина кюветы
            WidthCuvet1.InnerText = WidthCuvette; // и значение
            Izmerenie.AppendChild(WidthCuvet1); // и указываем кому принадлежит

            XmlNode BottomLine1 = xd.CreateElement("BottomLine"); // Нижняя граница
            BottomLine1.InnerText = BottomLine; // и значение
            Izmerenie.AppendChild(BottomLine1); // и указываем кому принадлежит

            XmlNode TopLine1 = xd.CreateElement("TopLine"); // Верхняя граница
            TopLine1.InnerText = TopLine; // и значение
            Izmerenie.AppendChild(TopLine1); // и указываем кому принадлежит

            XmlNode ND1 = xd.CreateElement("ND"); // НД
            ND1.InnerText = ND; // и значение
            Izmerenie.AppendChild(ND1); // и указываем кому принадлежит

            XmlNode Description1 = xd.CreateElement("Description"); // Примечание
            Description1.InnerText = Description; // и значение
            Izmerenie.AppendChild(Description1); // и указываем кому принадлежит

            XmlNode DateTime1 = xd.CreateElement("DateTime"); // дата создания градуировки
            DateTime1.InnerText = DateTime; // и значение
            Izmerenie.AppendChild(DateTime1); // и указываем кому принадлежит

            XmlNode DateTime1_1 = xd.CreateElement("DateTime1_1"); // Действительно до
            DateTime1_1.InnerText = label6.Text; // и значение
            Izmerenie.AppendChild(DateTime1_1); // и указываем кому принадлежит

            XmlNode DateTime1_1_1 = xd.CreateElement("DateTime1_1_1"); // Действительно до
            DateTime1_1_1.InnerText = numericUpDown1.Value.ToString(); // и значение
            Izmerenie.AppendChild(DateTime1_1_1); // и указываем кому принадлежит

            XmlNode Pogreshnost = xd.CreateElement("Pogreshnost"); // Действительно до
            Pogreshnost.InnerText = textBox3.Text; // и значение
            Izmerenie.AppendChild(Pogreshnost); // и указываем кому принадлежит

            XmlNode Ispolnitel1 = xd.CreateElement("Ispolnitel"); // Примечание
            Ispolnitel1.InnerText = Ispolnitel; // и значение
            Izmerenie.AppendChild(Ispolnitel1); // и указываем кому принадлежит

            XmlNode CountSeriya1 = xd.CreateElement("CountSeriyal"); // Примечание
            CountSeriya1.InnerText = CountSeriya; // и значение
            Izmerenie.AppendChild(CountSeriya1); // и указываем кому принадлежит

            XmlNode CountInSeriya1 = xd.CreateElement("CountInSeriyal"); // Примечание
            CountInSeriya1.InnerText = CountInSeriya; // и значение
            Izmerenie.AppendChild(CountInSeriya1); // и указываем кому принадлежит

            XmlNode edconctr1 = xd.CreateElement("edconctr");
            edconctr1.InnerText = edconctr;
            Izmerenie.AppendChild(edconctr1);
            XmlNode USE_CO_XML = xd.CreateElement("USE_CO_XML"); // Примечание
            if (USE_KO == true)
            {
                USE_CO_XML.InnerText = "true";
            }
            else
            {
                USE_CO_XML.InnerText = "false";
            }

            Izmerenie.AppendChild(USE_CO_XML); // и указываем кому принадлежит

            XmlNode TypeYravn1 = xd.CreateElement("TypeYravn"); // Тип уравнения
            if (radioButton1.Checked == true)
            {
                TypeYravn1.InnerText = "Линейное через 0"; // и значение
            }
            else
            {
                if (radioButton2.Checked == true)
                {

                    TypeYravn1.InnerText = "Линейное";
                }
                else
                {
                    TypeYravn1.InnerText = "Квадратичное";
                }
            }

            Izmerenie.AppendChild(TypeYravn1); // и указываем кому принадлежит

            XmlNode TypeIzmer1 = xd.CreateElement("TypeIzmer"); // Тип уравнения
            if (radioButton4.Checked == true)
            {
                TypeIzmer1.InnerText = "A (C) - градуировочное уравнение (стандарт)"; // и значение
            }
            else
            {
                TypeIzmer1.InnerText = "C (A) - расчетное уравнение (прибор)";
            }

            Izmerenie.AppendChild(TypeIzmer1); // и указываем кому принадлежит

            // ЗАбиваем запись в документ  
            xd.DocumentElement.AppendChild(Izmerenie);
            HeaderCells = new string[this.Table1.Columns.Count];
            Cells1 = new string[this.Table1.Rows.Count - 1, this.Table1.Columns.Count];

            XmlNode Zavisimoct1 = xd.CreateElement("SposobZadan"); // Примечание
            Zavisimoct1.InnerText = SposobZadan; // и значение
            Izmerenie.AppendChild(Zavisimoct1); // и указываем кому принадлежит
            if (SposobZadan != "Ввод коэффициентов")
            {
                for (int i = 0; i < this.Table1.Rows.Count - 1; i++)
                {
                    XmlNode Cells2 = xd.CreateElement("Stroka");

                    XmlAttribute attribute1 = xd.CreateAttribute("Nomer");
                    attribute1.Value = Convert.ToString(i); // устанавливаем значение атрибута
                    Cells2.Attributes.Append(attribute1); // добавляем атрибут
                    for (int j = 0; j < this.Table1.Columns.Count; j++)
                    {

                        Cells1[i, j] = Convert.ToString(this.Table1.Rows[i].Cells[j].Value);

                        HeaderCells[j] = this.Table1.Columns[j].HeaderText;
                        XmlNode HeaderCells1 = xd.CreateElement("Stolbec"); // Столбец
                        HeaderCells1.InnerText = Cells1[i, j]; // и значение
                        Cells2.AppendChild(HeaderCells1); // и указываем кому принадлежит
                        XmlAttribute attribute = xd.CreateAttribute("Header");
                        attribute.Value = HeaderCells[j]; // устанавливаем значение атрибута
                        HeaderCells1.Attributes.Append(attribute); // добавляем атрибут                    
                    }
                    xd.DocumentElement.AppendChild(Cells2);
                }

            }
            else
            {
                XmlNode k_0 = xd.CreateElement("k0"); // Примечание
                k_0.InnerText = AgroText0.Text; // и значение
                Izmerenie.AppendChild(k_0); // и указываем кому принадлежит
                XmlNode k_1 = xd.CreateElement("k1"); // Примечание
                k_1.InnerText = AgroText1.Text; // и значение
                Izmerenie.AppendChild(k_1); // и указываем кому принадлежит
                XmlNode k_2 = xd.CreateElement("k2"); // Примечание
                k_2.InnerText = AgroText2.Text; // и значение
                Izmerenie.AppendChild(k_2); // и указываем кому принадлежит
            }


            //ds.WriteXml(filepath);

            fs.Close();         // Закрываем поток  
            xd.Save(filepath); // Сохраняем файл  

        }
        public void SaveAs2()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as XML File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "QA2 файл|*.qa2";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                CreateXMLDocument2(ref filepath2);
                WriteXml2(ref filepath2, ref filepath);
                button3.Enabled = true;
                button9.Enabled = true;
                печатьToolStripMenuItem1.Enabled = true;
            }
        }
        private void CreateXMLDocument2(ref string filepath2)
        {

            filepath2 = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath2, Encoding.UTF8);

            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        public void WriteXml2(ref string filepath2, ref string filepath)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath2, FileMode.Open);
            xd.Load(fs);

            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

            XmlNode Nazvanie = xd.CreateElement("Nazvanie"); // Название вещества
            Nazvanie.InnerText = "Измерения"; // и значение
            Izmerenie.AppendChild(Nazvanie); // и указываем кому принадлежит


            XmlNode WidthCuvet1 = xd.CreateElement("WidthCuvet"); // Ширина кюветы
            WidthCuvet1.InnerText = Opt_dlin_cuvet.Text; // и значение
            Izmerenie.AppendChild(WidthCuvet1); // и указываем кому принадлежит

            XmlNode Description1 = xd.CreateElement("Description"); // Примечание
            Description1.InnerText = textBox8.Text; // и значение
            Izmerenie.AppendChild(Description1); // и указываем кому принадлежит

            XmlNode DateTime1 = xd.CreateElement("DateTime"); // дата создания градуировки
            DateTime1.InnerText = dateTimePicker2.Text; // и значение
            Izmerenie.AppendChild(DateTime1); // и указываем кому принадлежит

            XmlNode Pogreshnost = xd.CreateElement("Pogreshnost"); // Погрешность
            Pogreshnost.InnerText = textBox7.Text; // и значение
            Izmerenie.AppendChild(Pogreshnost); // и указываем кому принадлежит

            XmlNode F1 = xd.CreateElement("F1"); // F1
            F1.InnerText = F1Text.Text; // и значение
            Izmerenie.AppendChild(F1); // и указываем кому принадлежит

            XmlNode F2 = xd.CreateElement("F2"); // ДF2
            F2.InnerText = F2Text.Text; // и значение
            Izmerenie.AppendChild(F2); // и указываем кому принадлежит

            XmlNode Gradfilepath = xd.CreateElement("filepath");
            Gradfilepath.InnerText = filepath;
            Izmerenie.AppendChild(Gradfilepath);

            XmlNode USE_CO_XML = xd.CreateElement("USE_CO_XML"); // Примечание
            if (USE_KO == true)
            {
                USE_CO_XML.InnerText = "true";
            }
            else
            {
                USE_CO_XML.InnerText = "false";
            }

            Izmerenie.AppendChild(USE_CO_XML); // и указываем кому принадлежит
            XmlNode CountSeriya1 = xd.CreateElement("CountSeriyal"); // Примечание
            CountSeriya1.InnerText = Convert.ToString(NoCaIzm1); // и значение
            Izmerenie.AppendChild(CountSeriya1); // и указываем кому принадлежит

            XmlNode CountInSeriya1 = xd.CreateElement("CountInSeriyal"); // Примечание
            if (USE_KO != true)
            {
                CountInSeriya1.InnerText = Convert.ToString(Table2.Rows.Count - 1); // и значение
            }
            else {
                CountInSeriya1.InnerText = Convert.ToString(Table2.Rows.Count - 2); // и значение
            }
            Izmerenie.AppendChild(CountInSeriya1); // и указываем кому принадлежит

            // ЗАбиваем запись в документ  
            xd.DocumentElement.AppendChild(Izmerenie);
            HeaderCells = new string[this.Table2.Columns.Count];
            Cells1 = new string[this.Table2.Rows.Count - 1, this.Table2.Columns.Count];

            for (int i = 0; i < this.Table2.Rows.Count - 1; i++)
            {
                XmlNode Cells2 = xd.CreateElement("Stroka");

                XmlAttribute attribute1 = xd.CreateAttribute("Nomer");
                attribute1.Value = Convert.ToString(i); // устанавливаем значение атрибута
                Cells2.Attributes.Append(attribute1); // добавляем атрибут
                for (int j = 0; j < this.Table2.Columns.Count; j++)
                {

                    Cells1[i, j] = Convert.ToString(this.Table2.Rows[i].Cells[j].Value);

                    HeaderCells[j] = this.Table2.Columns[j].HeaderText;
                    XmlNode HeaderCells1 = xd.CreateElement("Stolbec"); // Столбец
                    if (Cells1[i, j] != "")
                    {
                        HeaderCells1.InnerText = Cells1[i, j]; // и значение
                    }
                    else
                    {
                        HeaderCells1.InnerText = "-";
                    }
                    Cells2.AppendChild(HeaderCells1); // и указываем кому принадлежит
                    XmlAttribute attribute = xd.CreateAttribute("Header");
                    attribute.Value = HeaderCells[j]; // устанавливаем значение атрибута
                    HeaderCells1.Attributes.Append(attribute); // добавляем атрибут                    
                }
                xd.DocumentElement.AppendChild(Cells2);
            }

            fs.Close();         // Закрываем поток  
            xd.Save(filepath2); // Сохраняем файл  

        }
        public void CreateXMLDocumentIzmerenieFR(ref string filepath)
        {
            filepath = saveFileDialog1.FileName;
            XmlTextWriter xtw = new XmlTextWriter(filepath, Encoding.UTF8);
            xtw.WriteStartDocument();
            xtw.WriteStartElement("Data_Izmerenie");
            xtw.WriteEndDocument();
            xtw.Close();
        }
        public void WriteXmlIzmerenieFR(ref string filepath)
        {
            XmlDocument xd = new XmlDocument();
            FileStream fs = new FileStream(filepath, FileMode.Open);
            xd.Load(fs);

            XmlNode Izmerenie = xd.CreateElement("Izmerenie");

            XmlNode Version = xd.CreateElement("Version"); // Версия программы
            Version.InnerText = version; // и значение
            Izmerenie.AppendChild(Version); // и указываем кому принадлежит
            XmlNode Ispolnitel1 = xd.CreateElement("Ispolnitel"); // Примечание
            Ispolnitel1.InnerText = Ispolnitel; // и значение
            Izmerenie.AppendChild(Ispolnitel1); // и указываем кому принадлежит
            XmlNode Description1 = xd.CreateElement("Description"); // Примечание
            Description1.InnerText = Description; // и значение
            Izmerenie.AppendChild(Description1); // и указываем кому принадлежит

            XmlNode DateTime1 = xd.CreateElement("DateTime"); // дата создания градуировки
            DateTime1.InnerText = DateTime; // и значение
            Izmerenie.AppendChild(DateTime1); // и указываем кому принадлежит
            int countIzmer = IzmerenieFR_Table.Rows.Count - 1;
            XmlNode countIzmer1 = xd.CreateElement("countIzmer1");
            countIzmer1.InnerText = Convert.ToString(countIzmer);
            Izmerenie.AppendChild(countIzmer1);
            xd.DocumentElement.AppendChild(Izmerenie);
            HeaderCells = new string[this.IzmerenieFR_Table.Columns.Count];
            Cells1 = new string[this.IzmerenieFR_Table.Rows.Count - 1, this.IzmerenieFR_Table.Columns.Count];
            for (int i = 0; i < this.IzmerenieFR_Table.Rows.Count - 1; i++)
            {
                XmlNode Cells2 = xd.CreateElement("Stroka");

                XmlAttribute attribute1 = xd.CreateAttribute("Nomer");
                attribute1.Value = Convert.ToString(i); // устанавливаем значение атрибута
                Cells2.Attributes.Append(attribute1); // добавляем атрибут
                for (int j = 0; j < this.IzmerenieFR_Table.Columns.Count; j++)
                {

                    Cells1[i, j] = Convert.ToString(this.IzmerenieFR_Table.Rows[i].Cells[j].Value);

                    HeaderCells[j] = this.IzmerenieFR_Table.Columns[j].HeaderText;
                    XmlNode HeaderCells1 = xd.CreateElement("Stolbec"); // Столбец
                    if (Cells1[i, j] != "")
                    {
                        HeaderCells1.InnerText = Cells1[i, j]; // и значение
                    }
                    else
                    {
                        HeaderCells1.InnerText = "-";
                    }
                    Cells2.AppendChild(HeaderCells1); // и указываем кому принадлежит
                    XmlAttribute attribute = xd.CreateAttribute("Header");
                    attribute.Value = HeaderCells[j]; // устанавливаем значение атрибута
                    HeaderCells1.Attributes.Append(attribute); // добавляем атрибут                    
                }
                xd.DocumentElement.AppendChild(Cells2);
            }

            fs.Close();         // Закрываем поток  
            xd.Save(filepath); // Сохраняем файл  

        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                if ((Table1.RowCount < 1) && SposobZadan == "По СО")
                {
                    MessageBox.Show("Создайте Градуировку");

                }
                else
                {
                    Save();
                }
            }
            else
            {
                if (Table1.RowCount > 1)
                {
                    Save1();
                }
                else
                {
                    MessageBox.Show("Создайте Измерение");
                }
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Izmerenie1 = true;
            if (tabControl2.SelectedIndex == 0)
            {
                Open();
            }
            else
            {
                Open1();
            }
        }

        public void WLADDSTR1()
        {
            if (USE_KO == true)
            {
                Table1.Rows.Add(0, Convert.ToDouble(0.000));

                for (int i = 1; i <= NoCaSer; i++)
                {
                    Table1.Rows.Add(i, textBoxCO[i - 1].Text);


                }

                Table1.CurrentCell = this.Table1[3, 0];

            }
            else
            {
                for (int i = 1; i <= NoCaSer; i++)
                {
                    Table1.Rows.Add(i, textBoxCO[i - 1].Text);


                }

                Table1.CurrentCell = this.Table1[3, 0];
            }
            for (int i = 1; i <= NoCaIzm; i++)
            {
                if (USE_KO == false)
                {
                    Table1.Rows[NoCaSer].Cells["A;Ser (" + i].ReadOnly = true;
                }
                else
                {
                    Table1.Rows[NoCaSer + 1].Cells["A;Ser (" + i].ReadOnly = true;
                }
            }

            if (USE_KO == false)
            {
                Table1.Rows[NoCaSer].Cells["NoCo"].ReadOnly = true;
                Table1.Rows[NoCaSer].Cells["Concetr"].ReadOnly = true;
                Table1.Rows[NoCaSer].Cells["Asred"].ReadOnly = true;
            }
            else
            {
                Table1.Rows[NoCaSer + 1].Cells["NoCo"].ReadOnly = true;
                Table1.Rows[NoCaSer + 1].Cells["Concetr"].ReadOnly = true;
                Table1.Rows[NoCaSer + 1].Cells["Asred"].ReadOnly = true;
            }

            button11.Enabled = true;
        }

        public void WLADDSTRAgro1()
        {
            if (USE_KO == true)
            {
                AgroTable1.Rows.Add(0, Convert.ToDouble(0.000));

                for (int i = 1; i <= NoCaSer; i++)
                {
                    AgroTable1.Rows.Add(i, textBoxCO[i - 1].Text);


                }

                AgroTable1.CurrentCell = this.AgroTable1[3, 0];

            }
            else
            {
                for (int i = 1; i <= NoCaSer; i++)
                {
                    AgroTable1.Rows.Add(i, textBoxCO[i - 1].Text);


                }

                AgroTable1.CurrentCell = this.AgroTable1[3, 0];
            }
            for (int i = 1; i <= NoCaIzm; i++)
            {
                if (USE_KO == false)
                {
                    AgroTable1.Rows[NoCaSer].Cells["A;Ser (" + i].ReadOnly = true;
                }
                else
                {
                    AgroTable1.Rows[NoCaSer + 1].Cells["A;Ser (" + i].ReadOnly = true;
                }
            }

            if (USE_KO == false)
            {
                AgroTable1.Rows[NoCaSer].Cells["AgroNoCo"].ReadOnly = true;
                AgroTable1.Rows[NoCaSer].Cells["AgroConcetr"].ReadOnly = true;
                AgroTable1.Rows[NoCaSer].Cells["AgroASred"].ReadOnly = true;
            }
            else
            {
                AgroTable1.Rows[NoCaSer + 1].Cells["AgroNoCo"].ReadOnly = true;
                AgroTable1.Rows[NoCaSer + 1].Cells["AgroConcetr"].ReadOnly = true;
                AgroTable1.Rows[NoCaSer + 1].Cells["AgroASred"].ReadOnly = true;
            }

            button11.Enabled = true;
        }
        public void WLREMOVESTR1()
        {
            Table1.Rows.Clear();

        }
        public void WLREMOVESTRAgro1()
        {
            AgroTable1.Rows.Clear();

        }
        public void WLREMOVE2()
        {
            while (true)
            {
                int i = Table2.Columns.Count - 1;//С какого столбца начать
                if (Table2.Columns[i].Name == "Obrazec")
                    break;
                Table2.Columns.RemoveAt(i);
            }

        }

        public void WLADD2()
        {
            if (NoCaIzm1 >= 2)
            {
                for (int i = 1; i <= NoCaIzm1; i++)
                {

                    DataGridViewTextBoxColumn firstColumn2 =
                    new DataGridViewTextBoxColumn();
                    firstColumn2.HeaderText = "A; Сер." + i;
                    firstColumn2.Name = "A;Ser" + i;
                    firstColumn2.ValueType = Type.GetType("System.Double");
                    Table2.Columns.Add(firstColumn2);
                    DataGridViewTextBoxColumn firstColumn3 =
                    new DataGridViewTextBoxColumn();
                    firstColumn3.HeaderText = "C, " + edconctr + "; Сер." + i;
                    firstColumn3.Name = "C,edconctr;Ser." + i;
                    firstColumn3.ValueType = Type.GetType("System.Double");
                    Table2.Columns.Add(firstColumn3);
                    // Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A; Сер" + i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
                    firstColumn3.ReadOnly = true;
                    firstColumn3.Width = 50;
                    firstColumn2.Width = 50;
                }
            }
            else
            {

                DataGridViewTextBoxColumn firstColumn2_1 =
                        new DataGridViewTextBoxColumn();
                firstColumn2_1.HeaderText = "A; Сер." + 1;
                firstColumn2_1.Name = "A;Ser" + 1;
                firstColumn2_1.ValueType = Type.GetType("System.Double");
                Table2.Columns.Add(firstColumn2_1);
                DataGridViewTextBoxColumn firstColumn3_1 =
                new DataGridViewTextBoxColumn();
                firstColumn3_1.HeaderText = "C, " + edconctr + "; Сер." + 1;
                firstColumn3_1.Name = "C,edconctr;Ser." + 1;
                firstColumn3_1.ValueType = Type.GetType("System.Double");
                Table2.Columns.Add(firstColumn3_1);
                firstColumn3_1.ReadOnly = true;
                firstColumn3_1.Width = 50;
                firstColumn2_1.Width = 50;
            }
            if (selet_rezim == 2)
            {
                DataGridViewTextBoxColumn firstColumn4 =
                new DataGridViewTextBoxColumn();
                firstColumn4.HeaderText = "Cср, " + edconctr;
                firstColumn4.Name = "Ccr";
                firstColumn4.ValueType = Type.GetType("System.Double");
                Table2.Columns.Add(firstColumn4);
                firstColumn4.ReadOnly = true;
                DataGridViewTextBoxColumn firstColumn5 =
                new DataGridViewTextBoxColumn();
                firstColumn5.HeaderText = "d, %";
                firstColumn5.Name = "d%";
                firstColumn5.ValueType = Type.GetType("System.Double");
                firstColumn5.ReadOnly = true;
                Table2.Columns.Add(firstColumn5);
                firstColumn4.Width = 100;
                firstColumn5.Width = 50;
            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
            RegistryKey hkcr = Registry.ClassesRoot;
            RegistryKey excelKey = hkcr.OpenSubKey("Excel.Application");
            bool excelInstalled = excelKey == null ? false : true;
            if (excelInstalled == true)
            {
                if (selet_rezim == 2)
                {
                    if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
                    {
                        SaveExcel();
                    }
                    else
                    {
                        if (tabControl2.SelectedIndex != 0 && SposobZadan == "По СО")
                        {
                            SaveExcel1();
                        }
                    }
                }
                else
                {
                    if (selet_rezim == 1)
                    {
                        IzmerenieFR_TableSaveExcel();
                    }
                    else
                    {
                        if (selet_rezim == 6)
                        {
                            if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
                            {
                                SaveExcel();
                            }
                            else
                            {
                                if (tabControl2.SelectedIndex != 0 && SposobZadan == "По СО")
                                {
                                    SaveExcel1();
                                }
                            }
                        }
                        else
                        {
                            if (selet_rezim == 5)
                            {
                                SaveExcelScan();
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Внимание!! Экспорт в Ecxel не возможен! Отсутствует Excel!");
            }
        }
        public void SaveExcelScan()
        {
            bool doNotWrite = false;
            for (int j = 0; j < ScanTable.Rows.Count - 1; j++)
            {

                for (int i = 3; i < ScanTable.Rows[j].Cells.Count; i++)
                {
                    if (ScanTable.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                ExportToExcelScan();
            }
        }

        public void IzmerenieFR_TableSaveExcel()
        {
            bool doNotWrite = false;
            for (int j = 0; j < IzmerenieFR_Table.Rows.Count - 1; j++)
            {

                for (int i = 3; i < IzmerenieFR_Table.Rows[j].Cells.Count; i++)
                {
                    if (IzmerenieFR_Table.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                ExportToExcelIzmerenieFR();
            }
        }
        public void SaveExcel()
        {
            bool doNotWrite = false;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {

                for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                {
                    if (Table1.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                ExportToExcel();
            }
        }
        public void ExportToExcelScan()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                //Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);

                exApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < this.ScanTable.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = this.ScanTable.Columns[i - 1].HeaderText;
                }
                //Thread.Sleep(500);
                for (int i = 0; i < this.ScanTable.Rows.Count; i++)
                {
                    // Thread.Sleep(200);
                    for (int j = 0; j < this.ScanTable.Columns.Count; j++)
                    {
                        exApp.Cells[i + 2, j + 1] = this.ScanTable.Rows[i].Cells[j].Value;
                    }
                }

                exApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                exApp.ActiveWorkbook.Saved = true;
                exApp.Visible = true;
            }
        }

        public void ExportToExcel()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                //Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);

                exApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < this.Table1.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = this.Table1.Columns[i - 1].HeaderText;
                }
                Thread.Sleep(500);
                for (int i = 0; i < this.Table1.Rows.Count; i++)
                {
                    Thread.Sleep(2000);
                    for (int j = 0; j < this.Table1.Columns.Count; j++)
                    {
                        exApp.Cells[i + 2, j + 1] = this.Table1.Rows[i].Cells[j].Value;
                    }
                }

                exApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                exApp.ActiveWorkbook.Saved = true;
                exApp.Visible = true;
                //  exApp.Quit();

            }
        }
        public void ExportToExcelIzmerenieFR()
        {
            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                //Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);

                exApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < this.IzmerenieFR_Table.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = this.IzmerenieFR_Table.Columns[i - 1].HeaderText;
                }
                Thread.Sleep(500);
                for (int i = 0; i < this.IzmerenieFR_Table.Rows.Count; i++)
                {
                    Thread.Sleep(2000);
                    for (int j = 0; j < this.IzmerenieFR_Table.Columns.Count; j++)
                    {
                        string value = Convert.ToString(this.IzmerenieFR_Table.Rows[i].Cells[j].Value);
                        value = value.Replace(",", ".");
                        exApp.Cells[i + 2, j + 1] = value;

                    }
                }

                exApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                exApp.ActiveWorkbook.Saved = true;
                exApp.Visible = true;
                //  exApp.Quit();

            }
        }
        public void SaveExcel1()
        {

            bool doNotWrite = false;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {

                for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                {
                    if (Table2.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                ExportToExcel2();
            }


        }
        public void ExportToExcel2()
        {

            saveFileDialog1.InitialDirectory = "C";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application exApp = new Microsoft.Office.Interop.Excel.Application();
                //Excel.Application exApp = new Excel.Application();
                exApp.Application.Workbooks.Add(Type.Missing);

                exApp.Columns.ColumnWidth = 20;
                for (int i = 1; i < this.Table2.Columns.Count + 1; i++)
                {
                    exApp.Cells[1, i] = this.Table2.Columns[i - 1].HeaderText;
                }
                Thread.Sleep(500);
                for (int i = 0; i < this.Table2.Rows.Count; i++)
                {
                    Thread.Sleep(2000);
                    for (int j = 0; j < this.Table2.Columns.Count; j++)
                    {
                        exApp.Cells[i + 2, j + 1] = this.Table2.Rows[i].Cells[j].Value;
                    }
                }

                exApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                exApp.ActiveWorkbook.Saved = true;
                // exApp.Quit();
                exApp.Visible = true;
            }

        }
        int countButtonClick = 1;
        private void button10_Click(object sender, EventArgs e)
        {
            if (selet_rezim == 2)
            {
                if (tabControl2.SelectedIndex == 0)
                {
                    NewGrad(ref CountSeriya, ref CountInSeriya);
                }
                else
                {
                    NewIzmer();
                }
            }
            else
            {
                if (selet_rezim == 6)
                {
                    if (tabControl2.SelectedIndex == 0)
                    {
                        NewGrad(ref CountSeriya, ref CountInSeriya);
                    }
                    else
                    {
                        ///NewIzmerenie();
                    }
                }
                else
                {
                    if (selet_rezim == 5)
                    {

                    }
                }
            }

        }
        public void NewGrad(ref string CountSeriya, ref string CountInSeriya)
        {
            CountSeriya1 = CountSeriya;
            CountInSeriya1 = CountInSeriya;
            NewGraduirovka _NewGraduirovka = new NewGraduirovka(this);
            _NewGraduirovka.ShowDialog();
            _NewGraduirovka.button1.Click += (NewGraduirovka, eSlave) =>
            {
                Veshestvo1 = _NewGraduirovka.Veshestvo.Text;
                wavelength1 = _NewGraduirovka.WL_grad.Text;
                WidthCuvette = _NewGraduirovka.Opt_dlin_cuvet.Text;
                BottomLine = _NewGraduirovka.Down.Text;
                TopLine = _NewGraduirovka.Up.Text;
                ND = _NewGraduirovka.ND.Text;
                Description = _NewGraduirovka.Description.Text;
                DateTime = _NewGraduirovka.dateTimePicker1.Text;
                Ispolnitel = _NewGraduirovka.Ispolnitel.Text;
                textBox3.Text = _NewGraduirovka.textBox4.Text;
                CountSeriya1 = _NewGraduirovka.numericUpDown3.Text;
                CountInSeriya1 = _NewGraduirovka.numericUpDown4.Text;
                edconctr = _NewGraduirovka.Ed.Text;
                Days = Convert.ToInt32(_NewGraduirovka.numericUpDown1.Value);
                label6.Text = dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy");
                if (_NewGraduirovka.radioButton7.Checked == true)
                {
                    this.AgroText0.Text = string.Format("{0:0.0000}", _NewGraduirovka.k0Text.Text);
                    this.AgroText1.Text = string.Format("{0:0.0000}", _NewGraduirovka.k1Text.Text);
                    this.AgroText2.Text = string.Format("{0:0.0000}", _NewGraduirovka.k2Text.Text);


                }
                else
                {
                    if (_NewGraduirovka.radioButton6.Checked == true)
                    {
                        //  this.textBox4.Text = string.Format("{0:0.0000}", 0);
                        //   this.textBox5.Text = string.Format("{0:0.0000}", 0);
                        ///   this.textBox6.Text = string.Format("{0:0.0000}", 0);
                    }
                }

                if (_NewGraduirovka.radioButton4.Checked == true)
                {
                    Zavisimoct = "A(C)";
                    radioButton4.Checked = true;
                    label14.Text = "A(C)";
                    /*             textBox4.Text = "";
                                 textBox5.Text = "";
                                 textBox6.Text = "";*/
                }
                else
                {
                    Zavisimoct = "C(A)";
                    radioButton5.Checked = true;
                    label14.Text = "C(A)";
                    /*  textBox4.Text = "";
                      textBox5.Text = "";
                      textBox6.Text = "";*/
                }
                if (_NewGraduirovka.radioButton1.Checked == true)
                {
                    aproksim = "Линейная через 0";
                }
                else
                {
                    if (_NewGraduirovka.radioButton2.Checked == true)
                    {
                        aproksim = "Линейная";
                    }
                    else
                    {
                        aproksim = "Квадратичная";
                    }
                }
                if (_NewGraduirovka.radioButton6.Checked == true)
                {
                    SposobZadan = "По СО";
                }
                else
                {
                    SposobZadan = "Ввод коэффициентов";
                    AgroText0.Text = _NewGraduirovka.k0Text.Text;
                    AgroText1.Text = _NewGraduirovka.k1Text.Text;
                    AgroText2.Text = _NewGraduirovka.k2Text.Text;
                    k0 = Convert.ToDouble(AgroText0.Text);
                    k1 = Convert.ToDouble(AgroText1.Text);
                    k2 = Convert.ToDouble(AgroText2.Text);
                    functionA();
                }
                if (_NewGraduirovka.USE_KO.Checked == true)
                {
                    USE_KO = true;


                }
                else
                {
                    USE_KO = false;
                }

            };
            CountSeriya = CountSeriya1;
            CountInSeriya = CountInSeriya1;
            dateTimePicker1.Text = DateTime;
            tabControl2.SelectTab(tabPage3);
            /*    chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();*/
            //       Table1.Columns.Add("X*X", "Concetr*Concetr");
            //Table1.Columns.Add("X*Y", "Asred*Concetr");
            //CalibrovkaGrad();
            // MessageBox.Show(Convert.ToString(Days));
            while (true)
            {
                int i = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns.Count == 3 + NoCaIzm)
                    break;
                Table1.Columns.RemoveAt(i);
            }
            if (Convert.ToInt32(CountInSeriya) < 3)
            {
                radioButton3.Enabled = false;
            }
            else
            {
                if (Convert.ToInt32(CountInSeriya) < 2)
                {
                    radioButton2.Enabled = false;
                    radioButton3.Enabled = false;
                }
                else
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                }
            }
            /*       RR.Text = "";
                   SKO.Text = "";
                   label21.Text = "";
                   label22.Text = "";*/
            if (SposobZadan == "По СО")
            {
                /*         textBox4.Text = "";
                         textBox5.Text = "";
                         textBox6.Text = "";*/
            }
            Table1.Rows[Table1.Rows.Count - 1].Cells[0].Value = "";
            if (SposobZadan == "Ввод коэффициентов")
            {
                functionA();
            }

        }
        public void NewIzmer()
        {
            Parametr1 _Parametr1 = new Parametr1(this);
            _Parametr1.ShowDialog();
            if (ComPort == true)
            {
                button14.Enabled = true;
            }
            else
            {
                button14.Enabled = false;
            }

        }


        public void Table2_UseCo()
        {
            double CCR = 0.0;

            double maxEl;
            double minEl;
            double serValue = 0;
            int cellnull = 0;
            int count = 0;
            for (int i1 = 0; i1 < Table2.RowCount - 1; i1++)
            {
                serValue = 0;
                El = new double[NoCaIzm1];

                double SredValue = 0;
                for (int i = 1; i <= NoCaIzm1; i++)
                {
                    if (Table2.Rows[i1].Cells["A;Ser" + i].Value == null)
                    {
                        cellnull++;
                    }
                    else
                    {
                        if (aproksim == "Линейная через 0")
                        {

                            if (Table2.Rows[0].Cells["A;Ser" + i].Value.ToString() != "" && Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString() != "")
                            {
                                if ((Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString()) > Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString())) && count == 0)
                                {
                                    if (count == 0)
                                    {
                                        count++;
                                        MessageBox.Show("Оптическая плотность контрольногго образца не может быть больше иззмеряемого!");
                                    }

                                }

                                serValue = (Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString()) - Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString())) / Convert.ToDouble(AgroText1.Text);
                            }
                            else
                            {

                                serValue = 0;
                                if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() == null)
                                {
                                    MessageBox.Show("Измерьте Контрольный образец!");
                                    return;


                                }
                            }
                        }
                        if (aproksim == "Линейная")
                        {
                            if (Table2.Rows[0].Cells["A;Ser" + i].Value.ToString() != "" && Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i].Value.ToString() != "")
                            {
                                if ((Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString()) > Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString())) && count == 0)
                                {
                                    if (count == 0)
                                    {
                                        count++;
                                        MessageBox.Show("Оптическая плотность контрольногго образца не может быть больше иззмеряемого!");
                                    }
                                }
                                serValue = ((Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString()) - Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / Convert.ToDouble(AgroText1.Text);
                            }
                            else
                            {

                                serValue = 0;
                                if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() == null)
                                {
                                    MessageBox.Show("Измерьте Контрольный образец!");
                                    return;


                                }
                            }

                        }
                        if (aproksim == "Квадратичная")
                        {
                            if (Table2.Rows[0].Cells["A;Ser" + i].Value.ToString() != "" && Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString() != "")
                            {
                                if ((Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString()) > Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString())) && count == 0)
                                {
                                    if (count == 0)
                                    {
                                        count++;
                                        MessageBox.Show("Оптическая плотность контрольногго образца не может быть больше иззмеряемого!");
                                    }
                                }
                                serValue = ((Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString()) - Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                            }
                            else
                            {
                                serValue = 0;
                                if (Table2.Rows[0].Cells["A;Ser" + i].Value.ToString() == null)
                                {
                                    MessageBox.Show("Измерьте Контрольый образец!");
                                    return;


                                }
                            }
                        }
                        double CValue1 = Convert.ToDouble(F1Text.Text);
                        double CValue2 = Convert.ToDouble(F2Text.Text);

                        if (serValue >= 0)
                        {
                            Table2.Rows[i1].Cells["C,edconctr;Ser." + i].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                            SredValue += Convert.ToDouble(Table2.Rows[i1].Cells["C,edconctr;Ser." + i].Value.ToString());
                        }
                        else
                        {
                            Table2.Rows[i1].Cells["C,edconctr;Ser." + i].Value = "";
                        }
                        if (Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i].Value.ToString()) > Convert.ToDouble(Table2.Rows[i1].Cells["A;Ser" + i].Value.ToString()))
                        {
                            if (selet_rezim == 2)
                            {
                                Table2.Rows[i1].Cells["Ccr"].Value = "";
                            }

                        }
                        else {
                            if (selet_rezim == 2)
                            {
                                CCR = SredValue / NoCaIzm1;
                                if (Convert.ToDouble(textBox7.Text) >= 1)
                                {

                                    Table2.Rows[i1].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.0000}", ((CCR * Convert.ToDouble(textBox7.Text))) / 100);
                                }
                                else
                                {

                                    Table2.Rows[i1].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                                }

                                //Table2.Rows[j].Cells["d%"].Value = El.Max();
                                if (Table2.Rows[i1].Cells["C,edconctr;Ser." + i].Value == null)
                                {
                                    cellnull++;
                                }
                                else
                                {
                                    if (Table2.Rows[i1].Cells["C,edconctr;Ser." + i].Value.ToString() != "")
                                    {
                                        El[i - 1] = Convert.ToDouble(Table2.Rows[i1].Cells["C,edconctr;Ser." + i].Value.ToString());
                                    }
                                }
                            }
                        }




                    }
                }
                if (selet_rezim == 2)
                {
                    Array.Sort(El);
                    maxEl = El[El.Length - 1];
                    minEl = El[0];
                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                    double b = a;


                    if (minEl == 0)
                    {
                        Table2.Rows[i1].Cells["d%"].Value = 0.0000;
                    }
                    else
                    {
                        Table2.Rows[i1].Cells["d%"].Value = string.Format("{0:0.0}", b);

                    }
                }
            }
            for (int i1 = 1; i1 <= NoCaIzm1; i1++)
            {
                Table2.Rows[0].Cells["C,edconctr;Ser." + i1].Value = "";
                if (selet_rezim == 2)
                {
                    Table2.Rows[0].Cells["Ccr"].Value = "";
                    Table2.Rows[0].Cells["d%"].Value = "";
                }
            }


        }
        private void параметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                NewGrad(ref CountSeriya, ref CountInSeriya);
            }
            else
            {
                NewIzmer();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (selet_rezim == 2)
            {
                if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
                {
                    SaveToPdf();
                }
                else
                {
                    if (tabControl2.SelectedIndex == 0 && SposobZadan != "По СО")
                    {
                        SaveToPdf1();
                    }
                    else
                    {
                        SaveTpPdf2();
                    }
                }
            }
            else
            {
                if (selet_rezim == 1)
                {
                    IzmerenieFRSavePDF();
                }
                else
                {
                    if (selet_rezim == 6)
                    {
                        if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
                        {
                            SaveToPdf();
                        }
                        else
                        {
                            if (tabControl2.SelectedIndex == 0 && SposobZadan != "По СО")
                            {
                                SaveToPdf1();
                            }
                            else
                            {
                                SaveTpPdf2();
                            }
                        }
                    }
                }
            }
        }
        public void IzmerenieFRSavePDF()
        {
            bool doNotWrite = false;
            for (int j = 0; j < IzmerenieFR_Table.Rows.Count - 1; j++)
            {

                for (int i = 3; i < IzmerenieFR_Table.Rows[j].Cells.Count; i++)
                {
                    if (IzmerenieFR_Table.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                IzmerenieFRExportToPdf();
            }
        }
        public void SaveToPdf()
        {
            bool doNotWrite = false;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {

                for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                {
                    if (Table1.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                ExportToPDF1();
            }
        }
        public void SaveToPdf1()
        {
            ExportToPDF1();
        }

        public void ExportToPDF1()
        {
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            string head = @"Расчет линейного градуировочного графика";
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 18f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font fontBold1 = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font1 = new iTextSharp.text.Font(baseFont, 5f, iTextSharp.text.Font.BOLD);
            PdfPTable pdfTable = new PdfPTable(Table1.ColumnCount);
            PdfPTable pdfTable1 = new PdfPTable(Table1.ColumnCount - 3 - NoCaIzm);
            if (NoCaIzm <= 3)
            {
                //Creating iTextSharp Table from the DataTable data
                pdfTable = new PdfPTable(Table1.ColumnCount);
                pdfTable.DefaultCell.Padding = 5;
                pdfTable.WidthPercentage = 100;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfTable.DefaultCell.BorderWidth = 1;
            }
            else
            {
                if (NoCaIzm > 3 && NoCaIzm <= 7)
                {
                    pdfTable = new PdfPTable(3 + NoCaIzm);
                    pdfTable.DefaultCell.Padding = 5;
                    pdfTable.WidthPercentage = 100;
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable.DefaultCell.BorderWidth = 1;
                    pdfTable1 = new PdfPTable(Table1.ColumnCount - 3 - NoCaIzm);
                    pdfTable1.DefaultCell.Padding = 5;
                    pdfTable1.WidthPercentage = 20;
                    pdfTable1.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable1.DefaultCell.BorderWidth = 1;
                }
                else
                {
                    pdfTable = new PdfPTable(3 + 5);
                    pdfTable.DefaultCell.Padding = 5;
                    pdfTable.WidthPercentage = 100;
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable.DefaultCell.BorderWidth = 1;
                    pdfTable1 = new PdfPTable(Table1.ColumnCount - 3 - 5);
                    pdfTable1.DefaultCell.Padding = 5;
                    pdfTable1.WidthPercentage = 100;
                    pdfTable1.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable1.DefaultCell.BorderWidth = 1;
                }
            }



            // iTextSharp.text.Font fontLeft = new iTextSharp.text.Font(baseFont, 9f, iTextSharp.text.Font.NORMAL);
            if (SposobZadan == "По СО")
            {

                if (NoCaIzm <= 3)
                {
                    PdfPCell cell;
                    for (int i = 0; i < Table1.ColumnCount; i++)
                    {
                        cell = new PdfPCell(new Phrase(Table1.Columns[i].HeaderText, fontBold1));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                        cell.BorderWidth = 1;
                        cell.Padding = 1;
                        cell.PaddingBottom = 5;
                        pdfTable.AddCell(cell);
                    }
                    for (int j = 0; j < Table1.Rows.Count; j++)
                    {
                        for (int i = 0; i < Table1.ColumnCount; i++)
                        {
                            pdfTable.AddCell(new Phrase(Convert.ToString(Table1.Rows[j].Cells[i].Value), font));
                        }
                    }
                }
                else
                {
                    if (NoCaIzm > 3 && NoCaIzm <= 7)
                    {
                        PdfPCell cell1;
                        PdfPCell cell;
                        int kIzmer1 = 0;
                        for (int i = 0; i < 3 + NoCaIzm; i++)
                        {
                            cell = new PdfPCell(new Phrase(Table1.Columns[kIzmer1].HeaderText, fontBold1));
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                            //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                            cell.BorderWidth = 1;
                            cell.Padding = 1;
                            cell.PaddingBottom = 5;
                            pdfTable.AddCell(cell);
                            kIzmer1++;
                        }
                        kIzmer1 = 0;
                        for (int j = 0; j < Table1.Rows.Count; j++)
                        {
                            for (int i = 0; i < 3 + NoCaIzm; i++)
                            {
                                pdfTable.AddCell(new Phrase(Convert.ToString(Table1.Rows[j].Cells[kIzmer1].Value), font));
                                kIzmer1++;
                            }
                            kIzmer1 = 0;
                        }
                        int kIzmer = 3 + NoCaIzm;
                        for (int i = 0; i < Table1.ColumnCount - 3 - NoCaIzm; i++)
                        {
                            cell1 = new PdfPCell(new Phrase(Table1.Columns[kIzmer].HeaderText, fontBold1));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                            //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                            cell1.BorderWidth = 1;
                            cell1.Padding = 1;
                            cell1.PaddingBottom = 5;
                            pdfTable1.AddCell(cell1);
                            kIzmer++;
                        }
                        kIzmer = 3 + NoCaIzm;
                        for (int j = 0; j < Table1.Rows.Count; j++)
                        {
                            for (int i = 0; i < Table1.ColumnCount - 3 - NoCaIzm; i++)
                            {
                                pdfTable1.AddCell(new Phrase(Convert.ToString(Table1.Rows[j].Cells[kIzmer].Value), font));
                                kIzmer++;
                            }
                            kIzmer = 3 + NoCaIzm;
                        }
                    }
                    else
                    {
                        PdfPCell cell1;
                        PdfPCell cell;
                        int kIzmer1 = 0;
                        for (int i = 0; i < 3 + 5; i++)
                        {
                            cell = new PdfPCell(new Phrase(Table1.Columns[kIzmer1].HeaderText, fontBold1));
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                            //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                            cell.BorderWidth = 1;
                            cell.Padding = 1;
                            cell.PaddingBottom = 5;
                            pdfTable.AddCell(cell);
                            kIzmer1++;
                        }
                        kIzmer1 = 0;
                        for (int j = 0; j < Table1.Rows.Count; j++)
                        {
                            for (int i = 0; i < 3 + 5; i++)
                            {
                                pdfTable.AddCell(new Phrase(Convert.ToString(Table1.Rows[j].Cells[kIzmer1].Value), font));
                                kIzmer1++;
                            }
                            kIzmer1 = 0;
                        }
                        int kIzmer = 3 + 5;
                        for (int i = 0; i < Table1.ColumnCount - 3 - 5; i++)
                        {
                            cell1 = new PdfPCell(new Phrase(Table1.Columns[kIzmer].HeaderText, fontBold1));
                            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                            //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                            cell1.BorderWidth = 1;
                            cell1.Padding = 1;
                            cell1.PaddingBottom = 5;
                            pdfTable1.AddCell(cell1);
                            kIzmer++;
                        }
                        kIzmer = 3 + 5;
                        for (int j = 0; j < Table1.Rows.Count; j++)
                        {
                            for (int i = 0; i < Table1.ColumnCount - 3 - 5; i++)
                            {
                                pdfTable1.AddCell(new Phrase(Convert.ToString(Table1.Rows[j].Cells[kIzmer].Value), font));
                                kIzmer++;
                            }
                            kIzmer = 3 + 5;
                        }
                    }
                }

            }


            var chartimage = new MemoryStream();
            chart1.SaveImage(chartimage, ChartImageFormat.Png);
            iTextSharp.text.Image Chart_Image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
            Chart_Image.ScalePercent(70f);
            iTextSharp.text.Rectangle orient = PageSize.A4;
            float margintop = 20;
            float marginleft = 25;
            float marginright = 25;
            float marginbottom = 5;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Document doc = new Document(orient, marginleft, marginright, margintop, marginbottom);

                PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                doc.Open();
                //iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("Image.jpeg");

                Paragraph welcomeParagraph = new Paragraph("Расчет линейного градуировочного графика\n", fontBold);
                welcomeParagraph.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                Paragraph Veshestvo2 = new Paragraph("Вещество: " + Veshestvo1, font);
                Paragraph wavelength2 = new Paragraph("Длина волны: " + wavelength1, font);
                Paragraph WidthCuvette2 = new Paragraph("Длина кюветы: " + WidthCuvette, font); ;
                Paragraph BottomLine2 = new Paragraph("Нижняя граница обнаружения: " + BottomLine, font);
                Paragraph TopLine2 = new Paragraph("Верхняя граница обнаружения: " + TopLine, font);
                Paragraph ND2 = new Paragraph("НД: " + ND, font);
                Paragraph Description2 = new Paragraph("Примечание: " + Description, font);
                Paragraph DateTime2 = new Paragraph("Дата: " + DateTime, font);
                Paragraph Ispolnitel2 = new Paragraph("Исполнитель: " + Ispolnitel, font);
                Paragraph GradYrav = new Paragraph("Градуировочное уравнение: " + label14.Text, font);
                Paragraph Table1 = new Paragraph("Таблица исходных данных\n\n", font);

                Paragraph InformationAboutPribor = new Paragraph("Информация о приборе\n", font);
                var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                const string model = @"pribor/model";
                var model_var = Path.Combine(applicationDirectory, model);
                string SerNomer_Text = @"pribor/SerNomer";
                var SerNomer_Text_var = Path.Combine(applicationDirectory, SerNomer_Text);

                string InventarNomer_Text = @"pribor/InventarNomer";
                var InventarNomer_Text_var = Path.Combine(applicationDirectory, InventarNomer_Text);

                string SrokIstech_Text = @"pribor/SrokIstech";
                var SrokIstech_Text_var = Path.Combine(applicationDirectory, SrokIstech_Text);

                string Poveren_Text = @"pribor/Poveren";
                var Poveren_Text_var = Path.Combine(applicationDirectory, Poveren_Text);
                StreamReader fs = new StreamReader(model_var);
                Paragraph Model = new Paragraph("Модель\n" + fs.ReadLine(), font);
                fs.Close();

                StreamReader fs1 = new StreamReader(SerNomer_Text_var);
                Paragraph SerNomer = new Paragraph("Серийный номер\n" + fs1.ReadLine(), font);
                fs1.Close();

                StreamReader fs2 = new StreamReader(InventarNomer_Text_var);
                Paragraph InventarNomer = new Paragraph("Инвентарный номер\n" + fs2.ReadLine(), font);
                fs2.Close();

                StreamReader fs3 = new StreamReader(Poveren_Text_var);
                DateTime data = Convert.ToDateTime(fs3.ReadLine());
                // data.Date.ToString("d.mm.yyyy"); 
                //  MessageBox.Show(Convert.ToString(data));   
                data = data.AddYears(1);
                fs3.Close();
                Paragraph Poveren = new Paragraph("Поверка действительна до\n" + data.Date.ToString("dd.MM.yyyy"), font);

                Paragraph Statistica = new Paragraph("Статистика: " + RR.Text + "\n                         " + SKO.Text + "\n                         " + label21.Text + "\n                         " + label22.Text, font);

                PdfPTable Information = new PdfPTable(6);
                PdfPCell Informationcell = new PdfPCell(Model);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(SerNomer);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(Poveren);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(InventarNomer);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);



                PdfPTable table = new PdfPTable(5);
                PdfPCell cell = new PdfPCell(Veshestvo2);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell();
                cell.BorderWidth = 0;
                table.AddCell(cell);

                cell = new PdfPCell(ND2);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell(wavelength2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                cell = new PdfPCell(WidthCuvette2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                cell = new PdfPCell(BottomLine2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                cell = new PdfPCell(TopLine2);
                cell.BorderWidth = 0;
                cell.Colspan = 5;
                table.AddCell(cell);

                Paragraph welcomeParagraph1 = new Paragraph("\n", fontBold);

                PdfPTable table1 = new PdfPTable(5);
                PdfPCell cell1 = new PdfPCell(Chart_Image);
                cell1.BorderWidth = 0;
                cell1.Colspan = 5;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(welcomeParagraph1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 5;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(DateTime2);
                cell1.BorderWidth = 0;
                cell1.Colspan = 5;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Ispolnitel2);
                cell1.BorderWidth = 0;
                cell1.Colspan = 5;
                table1.AddCell(cell1);




                doc.Add(welcomeParagraph);
                doc.Add(welcomeParagraph1);
                doc.Add(table);
                //  doc.Add(Veshestvo2);
                //  doc.Add(wavelength2);
                // doc.Add(WidthCuvette2);
                // doc.Add(BottomLine2);
                //  doc.Add(TopLine2);
                doc.Add(Description2);
                doc.Add(Statistica);
                doc.Add(welcomeParagraph1);
                doc.Add(InformationAboutPribor);

                doc.Add(Information);
                // doc.Add(welcomeParagraph1);
                if (SposobZadan == "По СО")
                {
                    doc.Add(Table1);

                    //  doc.Add(welcomeParagraph1);
                    if (NoCaIzm <= 3)
                    {
                        doc.Add(pdfTable);
                    }
                    else
                    {
                        if (NoCaIzm > 3 && NoCaIzm <= 7)
                        {
                            doc.Add(pdfTable);
                            doc.Add(welcomeParagraph1);
                            doc.Add(pdfTable1);
                        }
                        else
                        {
                            doc.Add(pdfTable);
                            doc.Add(welcomeParagraph1);
                            doc.Add(pdfTable1);
                        }
                    }
                }
                doc.Add(welcomeParagraph1);

                doc.Add(GradYrav);
                doc.Add(welcomeParagraph1);
                //    doc.Add(Chart_Image);
                //  doc.Add(welcomeParagraph1);
                doc.Add(table1);
                //  doc.Add(ND2);

                // doc.Add(DateTime2);
                // doc.Add(Ispolnitel2);

                doc.Close();
                /*   string filename = Application.StartupPath;
                   filename = Path.GetFullPath(Path.Combine(filename, ".\\Test.pdf"));
                   wbrPdf.Navigate(filename);*/
                filename = sfd.FileName;

            }

            /*   Spire.Pdf.PdfDocument pdfdocument = new Spire.Pdf.PdfDocument();
               pdfdocument.LoadFromFile(filename);        
               pdfdocument.PrintDocument.Print();
               pdfdocument.Dispose();
               */
        }

        public void IzmerenieFRExportToPdf()
        {
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 18f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font fontBold1 = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font1 = new iTextSharp.text.Font(baseFont, 5f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Rectangle orient = PageSize.A4;
            float margintop = 20;
            float marginleft = 25;
            float marginright = 25;
            float marginbottom = 5;
            PdfPTable pdfTable = new PdfPTable(IzmerenieFR_Table.ColumnCount);
            pdfTable = new PdfPTable(IzmerenieFR_Table.ColumnCount);
            pdfTable.DefaultCell.Padding = 5;
            pdfTable.WidthPercentage = 100;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable.DefaultCell.BorderWidth = 1;
            PdfPCell cell;
            for (int i = 0; i < IzmerenieFR_Table.ColumnCount; i++)
            {
                cell = new PdfPCell(new Phrase(IzmerenieFR_Table.Columns[i].HeaderText, fontBold1));
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                cell.BorderWidth = 1;
                cell.Padding = 1;
                cell.PaddingBottom = 5;
                pdfTable.AddCell(cell);
            }
            for (int j = 0; j < IzmerenieFR_Table.Rows.Count; j++)
            {
                for (int i = 0; i < IzmerenieFR_Table.ColumnCount; i++)
                {
                    pdfTable.AddCell(new Phrase(Convert.ToString(IzmerenieFR_Table.Rows[j].Cells[i].Value), font));
                }
            }
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Document doc = new Document(orient, marginleft, marginright, margintop, marginbottom);
                PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                doc.Open();
                Paragraph welcomeParagraph1 = new Paragraph("\n", fontBold);
                Paragraph welcomeParagraph = new Paragraph("Измерения в фотометрическом режиме\n", fontBold);
                welcomeParagraph.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                Paragraph Description2 = new Paragraph("Примечание: " + Description, font);
                Paragraph DateTime2 = new Paragraph("Дата: " + DateTime, font);
                Paragraph Ispolnitel2 = new Paragraph("Исполнитель: " + Ispolnitel, font);

                Paragraph InformationAboutPribor = new Paragraph("Информация о приборе:\n", font);
                var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                const string model = @"pribor/model";
                var model_var = Path.Combine(applicationDirectory, model);
                string SerNomer_Text = @"pribor/SerNomer";
                var SerNomer_Text_var = Path.Combine(applicationDirectory, SerNomer_Text);

                string InventarNomer_Text = @"pribor/InventarNomer";
                var InventarNomer_Text_var = Path.Combine(applicationDirectory, InventarNomer_Text);

                string SrokIstech_Text = @"pribor/SrokIstech";
                var SrokIstech_Text_var = Path.Combine(applicationDirectory, SrokIstech_Text);

                string Poveren_Text = @"pribor/Poveren";
                var Poveren_Text_var = Path.Combine(applicationDirectory, Poveren_Text);
                StreamReader fs = new StreamReader(model_var);
                Paragraph Model = new Paragraph("Модель\n" + fs.ReadLine(), font);
                fs.Close();

                StreamReader fs1 = new StreamReader(SerNomer_Text_var);
                Paragraph SerNomer = new Paragraph("Серийный номер\n" + fs1.ReadLine(), font);
                fs1.Close();

                StreamReader fs2 = new StreamReader(InventarNomer_Text_var);
                Paragraph InventarNomer = new Paragraph("Инвентарный номер\n" + fs2.ReadLine(), font);
                fs2.Close();

                StreamReader fs3 = new StreamReader(Poveren_Text_var);
                DateTime data = Convert.ToDateTime(fs3.ReadLine());
                // data.Date.ToString("d.mm.yyyy"); 
                //  MessageBox.Show(Convert.ToString(data));   
                data = data.AddYears(1);
                fs3.Close();
                Paragraph Poveren = new Paragraph("Поверка действительна до\n" + data.Date.ToString("dd.MM.yyyy"), font);
                PdfPTable Information = new PdfPTable(6);
                PdfPCell Informationcell = new PdfPCell(Model);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(SerNomer);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(Poveren);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(InventarNomer);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);
                Paragraph Table1 = new Paragraph("Таблица исходных данных\n\n", font);

                doc.Add(welcomeParagraph);
                doc.Add(welcomeParagraph1);
                doc.Add(Description2);
                doc.Add(welcomeParagraph1);
                doc.Add(InformationAboutPribor);
                doc.Add(welcomeParagraph1);
                doc.Add(Information);
                doc.Add(welcomeParagraph1);
                doc.Add(Table1);
                doc.Add(welcomeParagraph1);
                doc.Add(pdfTable);
                doc.Add(welcomeParagraph1);
                doc.Add(DateTime2);
                doc.Add(welcomeParagraph1);
                doc.Add(Ispolnitel2);
                doc.Add(welcomeParagraph1);
                doc.Close();
            }
        }

        public void WLADDSTR2()
        {
            count = 0;
            if (USE_KO == false)
            {
                if (NoCaSer1 > 1)
                {
                    for (int i = 1; i <= NoCaSer1; i++)
                    {
                        Table2.Rows.Add(i);
                        Table2.Rows[count].Cells["Column1"].Value = count + 1;
                        count++;
                    }
                }
                else
                {
                    Table2.Rows.Add(1);
                    Table2.Rows[count].Cells["Column1"].Value = count + 1;
                    count++;
                    Table2.Rows.Add(1);
                }
                for (int i = 0; i < Table2.RowCount - 1; i++)
                {

                    if (Table2.Rows[i].Cells["Column1"].Value == null)
                    {
                        Table2.Rows.RemoveAt(i);
                        i--;
                    }
                }
            }
            else
            {

                if (NoCaSer1 > 1)
                {
                    Table2.Rows.Add(0, "Контрольный", string.Format("{0:0.0000}", 0));
                    Table2.Rows[count].Cells["Column1"].Value = count;
                    count++;
                    for (int i = 1; i <= NoCaSer1; i++)
                    {
                        Table2.Rows.Add(i);
                        Table2.Rows[count].Cells["Column1"].Value = count;
                        count++;
                    }
                }
                else
                {
                    Table2.Rows.Add(0, "Контрольный", string.Format("{0:0.0000}", 0));
                    Table2.Rows[count].Cells["Column1"].Value = count;
                    count++;
                    Table2.Rows.Add(1, "");
                    Table2.Rows[count].Cells["Column1"].Value = count;
                    Table2.Rows.Add(1);
                }
                for (int i = 0; i < Table2.RowCount - 1; i++)
                {

                    if (Table2.Rows[i].Cells["Column1"].Value == null)
                    {
                        Table2.Rows.RemoveAt(i);
                        i--;
                    }
                }
            }
            //Table2.Rows.Add();
            Table2.CurrentCell = this.Table2[2, 0];
            for (int i = 1; i <= NoCaIzm1; i++)
            {
                Table2.Rows[0].Cells["C,edconctr;Ser." + i].ReadOnly = true;
                if (selet_rezim == 2)
                {
                    Table2.Rows[0].Cells["Ccr"].ReadOnly = true;
                    Table2.Rows[0].Cells["d%"].ReadOnly = true;
                }
            }
            Table2.Rows[Table2.RowCount - 1].ReadOnly = true;
            button11.Enabled = true;

        }

        public void WLREMOVESTR2()
        {
            Table2.Rows.Clear();

        }
        public void IzmerenieFR_RowsRemove2()
        {
            IzmerenieFR_Table.Rows.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string str = "";
            for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
            {
                str += PrinterSettings.InstalledPrinters[i] + "\n";
            }

            if (str != "")
            {
                switch (selet_rezim)
                {
                    case 2:
                        if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
                        {
                            PrintDoc();
                        }
                        else
                        {
                            if (tabControl2.SelectedIndex == 0 && SposobZadan != "По СО")
                            {
                                PrintDoc1();
                            }
                            else
                            {
                                PrintDoc2();
                            }
                        }
                        break;
                    case 1:
                        IzmerenieFR_TablePrintDoc();
                        break;
                    case 6:
                        if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
                        {
                            PrintDoc();
                        }
                        else
                        {
                            if (tabControl2.SelectedIndex == 0 && SposobZadan != "По СО")
                            {
                                PrintDoc1();
                            }
                            else
                            {
                                PrintDoc2();
                            }
                        }
                        break;
                    case 5:
                        this.TopMost = false;
                        PrintScan();
                        break;
                }

            }
            else
            {
                MessageBox.Show("Внимание! Принтер не найден! Подключите принтер!");
            }
        }
        public void PrintDoc()
        {
            bool doNotWrite = false;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {

                for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                {
                    if (Table1.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                printPreviewTable1.Document = printTable1;
                printPreviewTable1.ShowDialog();
            }
        }
        private int pagesCount;
        //  PaperSize paperSize = new PaperSize("papersize", 2100, 5);
        int prinPage = 0;
        int strcountScan = 0;
        int realwidth = 0;
        int realheight = 0;
        int width = 0;
        int height1 = 0;
        public void PrintScan()
        {
            bool doNotWrite = false;
            for (int j = 0; j < ScanTable.Rows.Count - 1; j++)
            {

                for (int i = 0; i < ScanTable.Rows[j].Cells.Count; i++)
                {
                    if (ScanTable.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                prinPage = 0;
                strcountScan = 0;
                realwidth = 0;
                realheight = 0;
                width = 0;
                height1 = 0;
                PrintScanTable.Document = ScanTablePrint;

                PrintScanTable.ShowDialog();

            }
        }

        public void IzmerenieFR_TablePrintDoc()
        {
            bool doNotWrite = false;
            for (int j = 0; j < IzmerenieFR_Table.Rows.Count - 1; j++)
            {

                for (int i = 3; i < IzmerenieFR_Table.Rows[j].Cells.Count; i++)
                {
                    if (IzmerenieFR_Table.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                IzmerenieFRprintPreviewTable1.Document = IzmerenieFRprintTable1;
                // IzmerenieFRprintTable1.DefaultPageSettings.PaperSize = paperSize;
                IzmerenieFRprintPreviewTable1.ShowDialog();
            }
        }
        public void PrintDoc1()
        {
            printPreviewTable1.Document = printTable1;
            printPreviewTable1.ShowDialog();
        }
        public void PrintDoc2()
        {
            bool doNotWrite = false;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {

                for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                {
                    if (Table2.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                if (Table2.Rows.Count >= 1)
                {
                    printPreviewTable2.Document = printTable2;
                    printPreviewTable2.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Создайте таблицу измерений!");
                }
            }
        }
        int height;

        public void ScanTablePrint_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (prinPage <= 0)
            {
                e.Graphics.DrawString("Измерение в режиме сканирования\n\n",
                new System.Drawing.Font("Times New Roman", 20, FontStyle.Bold), Brushes.Black, 100, 50);
                e.Graphics.DrawString("График сканирования\n\n",
                  new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 100, 100);
                height = ScanChart.Height;
                Bitmap bmp = new Bitmap(ScanChart.Width, ScanChart.Height);
                ScanChart.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, ScanChart.Width, ScanChart.Height));
                e.Graphics.DrawImage(bmp, 25, 130);
                height = height + 160;
                e.Graphics.DrawString("Таблица сканирования\n\n",
                new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 100, height);


                realwidth = ScanTable.Columns[0].Width + 5;
                realheight = height + 35;
                width = 100;
                height1 = ScanTable.Rows[0].Height + 5;
                for (int z = 0; z < ScanTable.Columns.Count; z++)
                {
                    e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height1);
                    e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height1);
                    e.Graphics.DrawString(ScanTable.Columns[z].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, realwidth, realheight);
                    realwidth = realwidth + width;
                }
                realwidth = ScanTable.Columns[0].Width + 5;
                realheight = realheight + 20;

                while (strcountScan < ScanTable.Rows.Count - 1)
                {
                    for (int j = 0; j < ScanTable.Columns.Count; j++)
                    {
                        e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height1);
                        e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height1);
                        e.Graphics.DrawString(ScanTable.Rows[strcountScan].Cells[j].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, realwidth, realheight);
                        realwidth = realwidth + width;

                    }
                    realwidth = ScanTable.Columns[0].Width + 5;
                    realheight = realheight + 20;

                    if (realheight > e.MarginBounds.Height)
                    {
                        height = 100;
                        e.HasMorePages = true;
                        //   strcountScan++;
                        prinPage++;
                        return;
                    }
                    else
                    {
                        e.HasMorePages = false;

                    }
                    // strcountScan++;

                    strcountScan++;
                }
            }
            else {
                realwidth = ScanTable.Columns[0].Width + 5;
                realheight = 20;
                width = 100;
                height1 = ScanTable.Rows[0].Height + 5;
                for (int z = 0; z < ScanTable.Columns.Count; z++)
                {
                    e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height1);
                    e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height1);
                    e.Graphics.DrawString(ScanTable.Columns[z].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, realwidth, realheight);
                    realwidth = realwidth + width;
                }
                realwidth = ScanTable.Columns[0].Width + 5;
                realheight = realheight + 20;

                while (strcountScan < ScanTable.Rows.Count - 1)
                {
                    for (int j = 0; j < ScanTable.Columns.Count; j++)
                    {
                        e.Graphics.FillRectangle(Brushes.AliceBlue, realwidth, realheight, width, height1);
                        e.Graphics.DrawRectangle(Pens.Black, realwidth, realheight, width, height1);
                        e.Graphics.DrawString(ScanTable.Rows[strcountScan].Cells[j].Value.ToString(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, realwidth, realheight);
                        realwidth = realwidth + width;

                    }
                    realwidth = ScanTable.Columns[0].Width + 5;
                    realheight = realheight + 20;

                    if (realheight > e.MarginBounds.Height)
                    {
                        height = 100;
                        e.HasMorePages = true;
                        //   strcountScan++;
                        prinPage++;
                        return;
                    }
                    else
                    {
                        e.HasMorePages = false;

                    }
                    strcountScan++;
                }
            }
            
        }

    
        public void printTableScan(PrintPageEventArgs e)
        {
            int height = ScanChart.Height + 160;
            
        }
        private void IzmerenieFRprintTable1_PrintPage(object sender, PrintPageEventArgs e)
        {
            /* int charactersOnPage = 0;
             int linesPerPage = 0;

             e.Graphics.MeasureString(stringToPrint, this.Font,
        e.MarginBounds.Size, StringFormat.GenericTypographic,
        out charactersOnPage, out linesPerPage);
             e.Graphics.DrawString(stringToPrint, this.Font, Brushes.Black,
         e.MarginBounds, StringFormat.GenericTypographic);*/

            
            e.Graphics.DrawString("Измерение в фототметрическом режиме\n\n", new System.Drawing.Font("Times New Roman", 20, FontStyle.Bold), Brushes.Black, 100, 50);
            e.Graphics.DrawString("Примечание:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 110);
            e.Graphics.DrawString(Description, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 155, 110);
            e.Graphics.DrawString("Информация о приборе:\n", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(25, 130));
            var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            const string model = @"pribor/model";
            var filePathToOpen = Path.Combine(applicationDirectory, model);



            StreamReader fs = new StreamReader(filePathToOpen);
            e.Graphics.DrawString("Модель: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(200, 150));
            e.Graphics.DrawString(fs.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(310, 150));
            fs.Close();


            const string SerNomer_Text = @"pribor/SerNomer";
            filePathToOpen = Path.Combine(applicationDirectory, SerNomer_Text);

            StreamReader fs1 = new StreamReader(filePathToOpen);
            e.Graphics.DrawString("Серийный номер: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(530, 150));
            e.Graphics.DrawString(fs1.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(700, 150));
            fs1.Close();


            const string InventarNomer_Text = @"pribor/InventarNomer";
            filePathToOpen = Path.Combine(applicationDirectory, InventarNomer_Text);

            StreamReader fs2 = new StreamReader(filePathToOpen);
            e.Graphics.DrawString("Инвентарный номер: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(500, 170));
            e.Graphics.DrawString(fs2.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(705, 170));
            fs2.Close();

            string Poveren_Text = @"pribor/Poveren";
            filePathToOpen = Path.Combine(applicationDirectory, Poveren_Text);

            StreamReader fs3 = new StreamReader(filePathToOpen);
            DateTime data = Convert.ToDateTime(fs3.ReadLine());
            // data.Date.ToString("d.mm.yyyy"); 
            //  MessageBox.Show(Convert.ToString(data));   
            data = data.AddYears(1);
            fs3.Close();
            e.Graphics.DrawString("Поверка действительна до: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(60, 170));
            e.Graphics.DrawString(data.Date.ToString("dd.MM.yyyy"), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(315, 170));
            e.Graphics.DrawString("Таблица исходных данных", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 200);
            IzmerenieFRPrintViewer1(sender, e);

            e.Graphics.DrawString("Дата:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY);
            e.Graphics.DrawString(DateTime, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 80, cordY);
            e.Graphics.DrawString("Исполнитель:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY + 30);
            e.Graphics.DrawString(Ispolnitel + "   _______________________", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 160, cordY + 30);
            /* stringToPrint = stringToPrint.Substring(charactersOnPage);

             // Check to see if more pages are to be printed.
             e.HasMorePages = (stringToPrint.Length > 0);*/
            e.HasMorePages = false;
        }
        
        private void printDocument1_PrintPage_2(object sender, PrintPageEventArgs e)
        {
            /* int charactersOnPage = 0;
             int linesPerPage = 0;*/

            e.Graphics.DrawString("Расчет линейного градуировочного графика\n\n", new System.Drawing.Font("Times New Roman", 20, FontStyle.Bold), Brushes.Black, 100, 50);
            e.Graphics.DrawString("Вещесво:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 110);
            e.Graphics.DrawString(Veshestvo1, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 115, 110);
            e.Graphics.DrawString("Длина волны:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 130);
            e.Graphics.DrawString(wavelength1, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 155, 130);
            e.Graphics.DrawString("Длина кюветы:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 150);
            e.Graphics.DrawString(WidthCuvette, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 190, 150);
            e.Graphics.DrawString("Границы обнаружения:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 170);
            e.Graphics.DrawString("Нижняя:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 250, 170);
            e.Graphics.DrawString(BottomLine, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 330, 170);
            e.Graphics.DrawString("Верхняя:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 450, 170);
            e.Graphics.DrawString(TopLine, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 540, 170);
            e.Graphics.DrawString("НД:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 450, 110);
            e.Graphics.DrawString(ND, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 500, 110);
            e.Graphics.DrawString("Примечание:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 190);
            e.Graphics.DrawString(Description, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 155, 190);
            e.Graphics.DrawString("Статистика:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 210);
            e.Graphics.DrawString(RR.Text + "                                               " + SKO.Text + "\n" + label21.Text + "          " + label22.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 140, 210);


            e.Graphics.DrawString("Информация о приборе:\n", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(25, 260));
            var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            const string model = @"pribor/model";
            var filePathToOpen = Path.Combine(applicationDirectory, model);

            

            StreamReader fs = new StreamReader(filePathToOpen);
            e.Graphics.DrawString("Модель: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(200, 280));
            e.Graphics.DrawString(fs.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(310, 280));
            fs.Close();


            const string SerNomer_Text = @"pribor/SerNomer";
            filePathToOpen = Path.Combine(applicationDirectory, SerNomer_Text);          

            StreamReader fs1 = new StreamReader(filePathToOpen);
            e.Graphics.DrawString("Серийный номер: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(530, 280));
            e.Graphics.DrawString(fs1.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(700, 280));
            fs1.Close();


            const string InventarNomer_Text = @"pribor/InventarNomer";
            filePathToOpen = Path.Combine(applicationDirectory, InventarNomer_Text);

            StreamReader fs2 = new StreamReader(filePathToOpen);
            e.Graphics.DrawString("Инвентарный номер: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(500, 300));
            e.Graphics.DrawString(fs2.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(705, 300));
            fs2.Close();

            string Poveren_Text = @"pribor/Poveren";
            filePathToOpen = Path.Combine(applicationDirectory, Poveren_Text);

            StreamReader fs3 = new StreamReader(filePathToOpen);
            DateTime data = Convert.ToDateTime(fs3.ReadLine());
            // data.Date.ToString("d.mm.yyyy"); 
            //  MessageBox.Show(Convert.ToString(data));   
            data = data.AddYears(1);
            fs3.Close();
            e.Graphics.DrawString("Поверка действительна до: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(60, 300));
            e.Graphics.DrawString(data.Date.ToString("dd.MM.yyyy"), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(315, 300));

            // e.Graphics.DrawString("Градуировочное уравнение: " + label14.Text, new System.Drawing.Font("C:\\Windows\\Fonts\\georgia.ttf", 12, FontStyle.Bold), Brushes.Black, new Point(50, 430));
            if (SposobZadan == "По СО")
            {
                e.Graphics.DrawString("Таблица исходных данных", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 330);
                if (NoCaIzm <= 3)
                {
                    Table1PrintViewer1(sender, e);
                }
                else
                {
                    if (NoCaIzm > 3 && NoCaIzm <= 7)
                    {
                        Table1PrintViewer2(sender, e);
                    }
                    else
                    {
                        Table1PrintViewer3(sender, e);
                    }
                }
            }
            else
            {
                cordY = 360;
            }


            e.Graphics.DrawString("Градуировочное уравнение:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY + 30);
            e.Graphics.DrawString(label14.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 285, cordY + 30);
            int height = chart1.Height;
            Bitmap bmp = new Bitmap(chart1.Width, chart1.Height);
            chart1.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, chart1.Width, chart1.Height));
            e.Graphics.DrawImage(bmp, 25, cordY + 60);
            cordY = cordY + chart1.Height + 60;
            e.Graphics.DrawString("Дата:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY);
            e.Graphics.DrawString(DateTime, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 80, cordY);
            e.Graphics.DrawString("Исполнитель:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY + 30);
            e.Graphics.DrawString(Ispolnitel + "   _______________________", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 160, cordY + 30);
            
            //  Paragraph Ispolnitel2 = new Paragraph("Исполнитель: " + Ispolnitel, font);

       //     stringToPrint = stringToPrint.Substring(charactersOnPage);

            // Check to see if more pages are to be printed.
           /* e.HasMorePages = (stringToPrint.Length > 0);*/


        }
        public void IzmerenieFRPrintViewer1(object sender, PrintPageEventArgs e)
        {
            int itemperpage = 0;
            int totalnumber = 0;
            int height = 230;
            int width = 25;

            Pen p = new Pen(Brushes.Black, 2.5f);
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[0].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[0].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[0].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[1].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[1].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[1].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[1].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[2].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[2].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[2].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[2].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[3].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[3].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[3].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[3].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[3].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[4].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[4].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[4].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[4].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[4].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[5].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[5].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[5].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[5].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[5].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[6].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[6].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            e.Graphics.DrawString(IzmerenieFR_Table.Columns[6].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[6].Width + 5, IzmerenieFR_Table.Rows[0].Height * 2));
            width = width + IzmerenieFR_Table.Columns[6].Width + 5;
            height = height + IzmerenieFR_Table.Rows[0].Height * 2;
            width = 25;
            int height1 = height;
            int width1_1 = width;
            while (totalnumber < IzmerenieFR_Table.Rows.Count - 1)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[0].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[0].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[0].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[0].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;

                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[1].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[1].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[1].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[1].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;

                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[2].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[2].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[2].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[2].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;


                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[3].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[3].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[3].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[3].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[1].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;


                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[4].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[4].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[4].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[4].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[4].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;


                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[5].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[5].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[5].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[5].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[5].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;



                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[6].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, IzmerenieFR_Table.Columns[6].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                e.Graphics.DrawString(IzmerenieFR_Table.Rows[totalnumber].Cells[6].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, IzmerenieFR_Table.Columns[6].Width + 5, IzmerenieFR_Table.Rows[totalnumber].Height));
                // width = width + IzmerenieFR_Table.Columns[6].Width;
                height += IzmerenieFR_Table.Rows[totalnumber].Height;
                totalnumber++;
                if (itemperpage < 20)
                {
                    itemperpage += 1;
                    e.HasMorePages = false;
                }
                else
                {
                    itemperpage = 0;
                    e.HasMorePages = true;
                    return;
                }
                
            }
            height = height1;
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;
            
            height = height1;
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;

            height = height1;
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;

            height = height1;
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;

            height = height1;
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;
           
            height = height1;
            width = width + IzmerenieFR_Table.Columns[0].Width + 5;
            
            cordY = height + 10;
        

        }
        ///Если меньше или равно 3
        public void Table1PrintViewer1(object sender, PrintPageEventArgs e)
        {
            int height = 360;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[1].Width;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[2].Width + 5;
            for (int i = 3; i <= Table1.Columns.Count - NoCaIzm; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[i].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[i].Width + 10;
            }
            for (int i = Table1.Columns.Count - NoCaIzm + 1; i < Table1.Columns.Count; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[i].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[i].Width + 10;
                // height = height + Table1.Rows[i].Height;
            }
            height = height + Table1.Rows[0].Height * 2;
            width = 25;
            int height1 = height;
            int width1_1 = width;

            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                // width = width + Table1.Columns[0].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[0].Width + 5;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                // width = width + Table1.Columns[1].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[1].Width;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                if (Table1.Rows[j].Cells[2].Value != null)
                {
                    e.Graphics.DrawString(Table1.Rows[j].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                // width = width + Table1.Columns[2].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[2].Width + 5;
            int width1 = width;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 3; i <= Table1.Columns.Count - NoCaIzm; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[i].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[i].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[i].Width + 10;
                    width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
            }

            height = height1;
            width1 = width1_1;
            width = width1;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = Table1.Columns.Count - NoCaIzm + 1; i < Table1.Columns.Count; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[i].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[i].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[i].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[i].Width + 10;
                }
                cordY = height;
                height += Table1.Rows[j].Height;
                width = width1;
            }


        }
        ///Если больше 3 и меньше или равно 7
        public void Table1PrintViewer2(object sender, PrintPageEventArgs e)
        {
            int height = 360;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[1].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[1].Width;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[2].Width + 5;
            int k = 3;
            for (int i = 0; i < NoCaIzm; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
            }
            height = height + Table1.Rows[0].Height * 2;
            /* Формируем значения */
            width = 25;
            int height1 = height;
            int width1_1 = width;


            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                // width = width + Table1.Columns[0].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[0].Width + 5;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                // width = width + Table1.Columns[1].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[1].Width;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                if (Table1.Rows[j].Cells[2].Value != null)
                {
                    e.Graphics.DrawString(Table1.Rows[j].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                // width = width + Table1.Columns[2].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[2].Width + 5;
            int width1 = width;
            k = 3;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < NoCaIzm; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
                k = 3;
            }
            /*Cancel*/
            height = height + 10;
            width = 25;
            k = NoCaIzm + 3;
            for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
                // height = height + Table1.Rows[i].Height;
            }

            /*Формируем вторую часть значений*/

            height = height + Table1.Rows[0].Height * 2;
            width1 = 25;
            width = 25;
            k = NoCaIzm + 3;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                }
                cordY = height;
                height += Table1.Rows[j].Height;
                width = width1;
                k = NoCaIzm + 3;
            }
            /*Cancel*/

        }


        /*Если больше 7*/
        public void Table1PrintViewer3(object sender, PrintPageEventArgs e)
        {
            int height = 360;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[1].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[1].Height * 2));
            e.Graphics.DrawString(Table1.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[1].Width;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            e.Graphics.DrawString(Table1.Columns[2].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[0].Height * 2));
            width = width + Table1.Columns[2].Width + 5;
            int k = 3;
            for (int i = 0; i < 7; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
            }


            height = height + Table1.Rows[0].Height * 2;
            /* Формируем значения */
            width = 25;
            int height1 = height;
            int width1_1 = width;


            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[0].Width + 5, Table1.Rows[j].Height));
                // width = width + Table1.Columns[0].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[0].Width + 5;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                e.Graphics.DrawString(Table1.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[1].Width, Table1.Rows[j].Height));
                // width = width + Table1.Columns[1].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[1].Width;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                if (Table1.Rows[j].Cells[2].Value != null)
                {
                    e.Graphics.DrawString(Table1.Rows[j].Cells[2].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[2].Width + 5, Table1.Rows[j].Height));
                }
                // width = width + Table1.Columns[2].Width;
                height += Table1.Rows[j].Height;
            }
            height = height1;
            width = width + Table1.Columns[2].Width + 5;
            int width1 = width;
            k = 3;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < 7; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
                k = 3;
            }
            /*Cancel*/

            height = height + 10;
            width = 25;
            k = 10;
            //k = 11;
            for (int i = 0; i < NoCaIzm - 7; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
            }

            for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                e.Graphics.DrawString(Table1.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[0].Height * 2));
                width = width + Table1.Columns[k].Width + 10;
                k++;
                // height = height + Table1.Rows[i].Height;
            }

            /*Формируем вторую часть значений*/

            height = height + Table1.Rows[0].Height * 2;
            height1 = height;
            width1 = 25;
            width = 25;
            k = 10;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < NoCaIzm - 7; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table1.Rows[j].Height;
                width = width1;
                k = 10;
            }
            width = (Table1.Columns[10].Width + 10) * (NoCaIzm - 7) + width;
            width1 = width;
            height = height1;
            k = 10 + NoCaIzm - 7;
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                for (int i = 0; i < Table1.Columns.Count - NoCaIzm - 3; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    if (Table1.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table1.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table1.Columns[k].Width + 10, Table1.Rows[j].Height));
                    }
                    width = width + Table1.Columns[k].Width + 10;
                    k++;
                }
                cordY = height;
                height += Table1.Rows[j].Height;
                width = width1;
                k = 10 + NoCaIzm - 7;
            }
            /*Cancel*/

        }
        private void printTable2_PrintPage(object sender, PrintPageEventArgs e)
        {
            cordY = 480;
            e.Graphics.DrawString("Протокол выполнения измерений\n\n", new System.Drawing.Font("Times New Roman", 20, FontStyle.Bold), Brushes.Black, 100, 50);
            e.Graphics.DrawString("Имя файла:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 110);
            e.Graphics.DrawString(filepath2, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 130, 110);
            e.Graphics.DrawString("Описание:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 130);
            e.Graphics.DrawString(textBox8.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 130, 130);
            e.Graphics.DrawString("Дата:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 150);
            e.Graphics.DrawString(dateTimePicker2.Value.ToString("dd.MM.yyyy"), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 120, 150);
            e.Graphics.DrawString("Длина волны:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 250, 150);
            e.Graphics.DrawString(wavelength1, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 390, 150);
            e.Graphics.DrawString("Погрешность методики: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 470, 150);
            e.Graphics.DrawString(textBox7.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 700, 150);
            e.Graphics.DrawString("Оптическая длина кюветы:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 180);
            e.Graphics.DrawString(Opt_dlin_cuvet.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 320, 180);
            e.Graphics.DrawString("F1 = ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 420, 180);
            e.Graphics.DrawString(F1Text.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 470, 180);
            e.Graphics.DrawString("F2 = ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 580, 180);
            e.Graphics.DrawString(F2Text.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 630, 180);
            //e.Graphics.DrawString("Таблица исходных данных", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 230);
            e.Graphics.DrawString("Градуировка:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 210);
            //   e.Graphics.DrawString(textBox8.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 130, 260);
            e.Graphics.DrawString("Имя файла:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 230);
            e.Graphics.DrawString(filepath, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 170, 230);
            e.Graphics.DrawString("Описание:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 260);
            e.Graphics.DrawString(Description, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 170, 260);
            e.Graphics.DrawString("Дата:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 290);
            e.Graphics.DrawString(DateTime, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 120, 290);
            e.Graphics.DrawString("Действительна до: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 230, 290);
            e.Graphics.DrawString(dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy"), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 405, 290);
            e.Graphics.DrawString("Погрешность методики:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 505, 290);
            e.Graphics.DrawString(textBox3.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 730, 290);
            e.Graphics.DrawString("Градуировочное уравнение:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 320);
            e.Graphics.DrawString(label14.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 330, 320);
            e.Graphics.DrawString("НД:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 60, 350);
            e.Graphics.DrawString(ND, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 100, 350);
            e.Graphics.DrawString("Статистика:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 380);
            e.Graphics.DrawString(RR.Text + "                                               " + SKO.Text + "\n" + label21.Text + "          " + label22.Text, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 140, 380);
            e.Graphics.DrawString("Информация о приборе:\n", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(25, 430));

            var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            const string model = @"pribor/model";
            var model_var = Path.Combine(applicationDirectory, model);
            string SerNomer_Text = @"pribor/SerNomer";
            var SerNomer_Text_var = Path.Combine(applicationDirectory, SerNomer_Text);

            string InventarNomer_Text = @"pribor/InventarNomer";
            var InventarNomer_Text_var = Path.Combine(applicationDirectory, InventarNomer_Text);

            string SrokIstech_Text = @"pribor/SrokIstech";
            var SrokIstech_Text_var = Path.Combine(applicationDirectory, SrokIstech_Text);

            string Poveren_Text = @"pribor/Poveren";
            var Poveren_Text_var = Path.Combine(applicationDirectory, Poveren_Text);
            StreamReader fs = new StreamReader(model_var);
            e.Graphics.DrawString("Модель: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(60, 450));
            e.Graphics.DrawString(fs.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(140, 450));
            fs.Close();

            StreamReader fs1 = new StreamReader(SerNomer_Text_var);
            e.Graphics.DrawString("Серийный номер: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(500, 450));
            e.Graphics.DrawString(fs1.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(700, 450));
            fs1.Close();

            StreamReader fs2 = new StreamReader(InventarNomer_Text_var);
            e.Graphics.DrawString("Инвентарный номер: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(500, 470));
            e.Graphics.DrawString(fs2.ReadLine(), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(705, 470));
            fs2.Close();

            StreamReader fs3 = new StreamReader(Poveren_Text_var);
            DateTime data = Convert.ToDateTime(fs3.ReadLine());
            // data.Date.ToString("d.mm.yyyy"); 
            //  MessageBox.Show(Convert.ToString(data));   
            data = data.AddYears(1);
            fs3.Close();
            e.Graphics.DrawString("Поверка действительна до: ", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new Point(60, 470));
            e.Graphics.DrawString(data.Date.ToString("dd.MM.yyyy"), new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new Point(315, 470));
            e.Graphics.DrawString("Данные измерений:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, 500);
            if (NoCaIzm1 <= 3)
            {
                Table2PrintViewer1(sender, e);
            }
            else
            {
                if (NoCaIzm1 > 3 && NoCaIzm1 <= 7)
                {
                    Table2PrintViewer2(sender, e);
                }
                else
                {
                    Table2PrintViewer3(sender, e);
                }
            }
            e.Graphics.DrawString("Измерения выполнил(а):", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY + 60);
            e.Graphics.DrawString(" ____________________________________________________", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 250, cordY + 60);


           /* e.Graphics.DrawString("Исполнитель:", new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, 25, cordY + 30);
            e.Graphics.DrawString(Ispolnitel, new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, 160, cordY + 30);*/

        }
        ///Если меньше или равно 3
        public void Table2PrintViewer1(object sender, PrintPageEventArgs e)
        {
            int height = 550;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            e.Graphics.DrawString(Table2.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            width = width + Table2.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[0].Height * 2));
            e.Graphics.DrawString(Table2.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[0].Height * 2));
            width = width + Table2.Columns[1].Width;

            for (int i = 2; i <= Table2.Columns.Count - NoCaIzm1; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[i].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[i].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[i].Width + 10;
            }
            for (int i = Table2.Columns.Count - NoCaIzm1 + 1; i < Table2.Columns.Count; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[i].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[i].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[i].Width + 10;
                // height = height + Table2.Rows[i].Height;
            }
            height = height + Table2.Rows[0].Height * 2;
            width = 25;
            int height1 = height;
            int width1_1 = width;

            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                e.Graphics.DrawString(Table2.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                // width = width + Table2.Columns[0].Width;
                height += Table2.Rows[j].Height;
            }
            height = height1;
            width = width + Table2.Columns[0].Width + 5;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                if (Table2.Rows[j].Cells[1].Value != null)
                {
                    e.Graphics.DrawString(Table2.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[j].Height));

                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                }
                // width = width + Table2.Columns[1].Width;
                height += Table2.Rows[j].Height;
            }
            height = height1;
            width = width + Table2.Columns[1].Width;

            int width1 = width;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = 2; i <= Table2.Columns.Count - NoCaIzm1; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[i].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[i].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[i].Width + 10;
                    width1_1 = width;
                }
                height += Table2.Rows[j].Height;
                width = width1;
            }

            height = height1;
            width1 = width1_1;
            width = width1;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = Table2.Columns.Count - NoCaIzm1 + 1; i < Table2.Columns.Count; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[i].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[i].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[i].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[i].Width + 10;
                }
                cordY = height;
                height += Table2.Rows[j].Height;
                width = width1;
            }


        }
        ///Если больше 3 и меньше или равно 7
        public void Table2PrintViewer2(object sender, PrintPageEventArgs e)
        {
            int height = 550;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            e.Graphics.DrawString(Table2.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            width = width + Table2.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[1].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[0].Height * 2));
            e.Graphics.DrawString(Table2.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[0].Height * 2));
            width = width + Table2.Columns[1].Width;

            int k = 2;
            for (int i = 0; i < NoCaIzm1 * 2; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[k].Width + 10;
                k++;
            }
            height = height + Table2.Rows[0].Height * 2;
            /* Формируем значения */
            width = 25;
            int height1 = height;
            int width1_1 = width;

            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                e.Graphics.DrawString(Table2.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                // width = width + Table2.Columns[0].Width;
                height += Table2.Rows[j].Height;
            }
            height = height1;
            width = width + Table2.Columns[0].Width + 5;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                if (Table2.Rows[j].Cells[1].Value != null)
                {
                    e.Graphics.DrawString(Table2.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[j].Height));

                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                }
                // width = width + Table2.Columns[1].Width;
                height += Table2.Rows[j].Height;
            }
            height = height1;
            width = width + Table2.Columns[1].Width;

            int width1 = width;
            k = 2;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = 0; i < NoCaIzm1 * 2; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table2.Rows[j].Height;
                width = width1;
                k = 2;
            }
            /*Cancel*/
            height = height + 10;
            width = 25;
            k = NoCaIzm1 * 2 + 2;
            for (int i = 0; i < Table2.Columns.Count - NoCaIzm1 * 2 - 2; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[k].Width + 10;
                k++;
                // height = height + Table2.Rows[i].Height;
            }

            /*Формируем вторую часть значений*/

            height = height + Table2.Rows[0].Height * 2;
            width1 = 25;
            width = 25;
            k = NoCaIzm1 * 2 + 2;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = 0; i < Table2.Columns.Count - NoCaIzm1 * 2 - 2; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[k].Width + 10;
                    k++;
                }
                cordY = height;
                height += Table2.Rows[j].Height;
                width = width1;
                k = NoCaIzm1 * 2 + 2;
            }
            /*Cancel*/

        }


        /*Если больше 7*/
        public void Table2PrintViewer3(object sender, PrintPageEventArgs e)
        {
            int height = 550;
            int width = 25;
            Pen p = new Pen(Brushes.Black, 2.5f);

            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            e.Graphics.DrawString(Table2.Columns[0].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[0].Width + 5, Table2.Rows[0].Height * 2));
            width = width + Table2.Columns[0].Width + 5;
            e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[1].Height * 2));
            e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[1].Height * 2));
            e.Graphics.DrawString(Table2.Columns[1].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[0].Height * 2));
            width = width + Table2.Columns[1].Width;

            int k = 2;
            for (int i = 0; i < 10; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[k].Width + 10;
                k++;
            }


            height = height + Table2.Rows[0].Height * 2;
            /* Формируем значения */
            width = 25;
            int height1 = height;
            int width1_1 = width;


            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                e.Graphics.DrawString(Table2.Rows[j].Cells[0].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[0].Width + 5, Table2.Rows[j].Height));
                // width = width + Table2.Columns[0].Width;
                height += Table2.Rows[j].Height;
            }
            height = height1;
            width = width + Table2.Columns[0].Width + 5;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                if (Table2.Rows[j].Cells[1].Value != null)
                {
                    e.Graphics.DrawString(Table2.Rows[j].Cells[1].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[j].Height));

                }
                else
                {
                    e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 12, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[1].Width, Table2.Rows[j].Height));
                }
                // width = width + Table2.Columns[1].Width;
                height += Table2.Rows[j].Height;
            }
            height = height1;
            width = width + Table2.Columns[1].Width;

            int width1 = width;
            k = 2;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = 0; i < 10; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table2.Rows[j].Height;
                width = width1;
                k = 2;
            }
            /*Cancel*/

            height = height + 10;
            width = 25;
            k = 12;
            //k = 11;
            for (int i = 0; i < NoCaIzm1 * 2 - 10; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[k].Width + 10;
                k++;
            }

            for (int i = 0; i < Table2.Columns.Count - NoCaIzm1 * 2 - 2; i++)
            {
                e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                e.Graphics.DrawString(Table2.Columns[k].HeaderText, new System.Drawing.Font("Times New Roman", 12, FontStyle.Bold), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[0].Height * 2));
                width = width + Table2.Columns[k].Width + 10;
                k++;
                // height = height + Table2.Rows[i].Height;
            }

            /*Формируем вторую часть значений*/

            height = height + Table2.Rows[0].Height * 2;
            height1 = height;
            width1 = 25;
            width = 25;
            k = 12;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = 0; i < NoCaIzm1 * 2 - 10; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[k].Width + 10;
                    k++;
                    //width1_1 = width;
                }
                height += Table2.Rows[j].Height;
                width = width1;
                k = 12;
            }
            width = (Table2.Columns[10].Width + 10) * (NoCaIzm1 * 2 - 10) + width;
            width1 = width;
            height = height1;
            k = 2 + NoCaIzm1 * 2;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {
                for (int i = 0; i < Table2.Columns.Count - NoCaIzm1 * 2 - 2; i++)
                {
                    e.Graphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    e.Graphics.DrawRectangle(p, new System.Drawing.Rectangle(width, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    if (Table2.Rows[j].Cells[k].Value != null)
                    {
                        e.Graphics.DrawString(Table2.Rows[j].Cells[k].Value.ToString(), new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    else
                    {
                        e.Graphics.DrawString("", new System.Drawing.Font("Times New Roman", 10, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(width + 10, height, Table2.Columns[k].Width + 10, Table2.Rows[j].Height));
                    }
                    width = width + Table2.Columns[k].Width + 10;
                    k++;
                }
                cordY = height;
                height += Table2.Rows[j].Height;
                width = width1;
                k = 2 + NoCaIzm1 * 2;
            }
            /*Cancel*/

        }

        private void экспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
            {
                SaveExcel();
            }
            else
            {
                if (tabControl2.SelectedIndex != 0 && SposobZadan == "По СО")
                {
                    SaveExcel1();
                }
            }
        }

        private void эксопртВPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
            {
                SaveToPdf();
            }
            else
            {
                if (tabControl2.SelectedIndex == 0 && SposobZadan != "По СО")
                {
                    SaveToPdf1();
                }
                else
                {
                    SaveTpPdf2();
                }
            }
        }


        private void печатьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0 && SposobZadan == "По СО")
            {
                PrintDoc();
            }
            else
            {
                if (tabControl2.SelectedIndex == 0 && SposobZadan != "По СО")
                {
                    PrintDoc1();
                }
                else
                {
                    PrintDoc2();
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            switch (selet_rezim)
            {
                case 2:
                    if (tabControl2.SelectedIndex == 0)
                    {
                        if (Table1.RowCount > 1)
                        {
                            if (textBox10.Text != GWNew.Text)
                            {
                                MessageBox.Show("Длина волны градуировки отличается от длины волны, установленной на приборе!\rИзмените настройки градуировки!");
                            }
                            Graduirovka(sender, e);
                        }
                        else
                        {
                            MessageBox.Show("Создайте градуировку по СО");
                        }
                    }
                    else
                    {
                        if (Table2.RowCount > 1)
                        {
                            if (textBox10.Text != GWNew.Text)
                            {
                                MessageBox.Show("Длина волны градуировки отличается от длины волны, установленной на приборе!\rИзмените настройки градуировки!");
                            }
                            Izmerenie(sender, e);
                        }
                        else
                        {
                            MessageBox.Show("Создайте измерение");
                        }
                    }
                    break;
                case 1:
                    if (IzmerenieFR_Table.RowCount > 1)
                    {
                        IzmerenieFr_izmer();
                    }
                    else
                    {
                        MessageBox.Show("Данная опреция невозможна! Создайте новое измерение!");
                    }
                    break;
                case 6:
                    if (tabControl2.SelectedIndex == 0)
                    {
                        if (Table1.RowCount > 1)
                        {
                            if (textBox10.Text != GWNew.Text)
                            {
                                MessageBox.Show("Длина волны градуировки отличается от длины волны, установленной на приборе!\rИзмените настройки градуировки!");
                            }
                            Graduirovka(sender, e);
                        }
                        else
                        {
                            MessageBox.Show("Создайте градуировку по СО");
                        }
                    }


                    else
                    {
                        if (Table2.RowCount > 1)
                        {
                            if (textBox10.Text != GWNew.Text)
                            {
                                MessageBox.Show("Длина волны градуировки отличается от длины волны, установленной на приборе!\rИзмените настройки градуировки!");
                            }
                            Table2.Rows[0].ReadOnly = true;
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].ReadOnly == true)
                            {
                                Table2.CurrentCell = this.Table2[2, Table2.CurrentCell.RowIndex + 1];
                            }

                            {
                                Application.DoEvents();
                                TableAgro2();
                                Application.DoEvents();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Создайте измерение");
                        }
                    }
                    break;
                case 5:
                    if (ScanTable.Rows.Count > 1)
                    {
                        if (scan_mass != null)
                        {
                            int notNull = 0;
                            for (int i = 0; i < ScanTable.Rows.Count; i++)
                            {
                                for (int k = 0; k < ScanTable.ColumnCount; k++)
                                {
                                    if (ScanTable.Rows[i].Cells[k].Value != null)
                                    {
                                        notNull++;
                                    }
                                }
                            }
                            if (notNull == (ScanTable.Rows.Count - 1) * ScanTable.ColumnCount)
                            {
                                DialogResult result = MessageBox.Show("Внимание! Ваша таблица перезапишется," +
                                    "при этом данные будут сохранены.",
                                    "Предупреждение",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Information,
                                    MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.DefaultDesktopOnly);
                                if (result == DialogResult.Yes)
                                {
                                    TableScan();
                                    TableScan_Save();
                                }
                                this.TopMost = true;

                            }
                            else {
                                this.TopMost = true;
                                Application.DoEvents();
                                TableScan();
                                Application.DoEvents();
                                TableScan_Save();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Вы забыли откалиброваться! Откалибруйтесь!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Создайте новое измерение!");
                    }
                    break;
                case 4:
                    ChartGraf();
                    if (timer2.Enabled == false)
                    { 
                        if (delay > 0)
                        {
                            timer1.Interval = Convert.ToInt32(1000); // 500 миллисекунд
                            timer1.Enabled = true;
                            timer1.Tick += TimerTick1;
                        }
                        else {

                            timeLeft = Convert.ToInt32(start);
                            TableKinetica1.Rows.Clear();
                            TableKinetica(sender, e);
                            button14.Enabled = false;
                            dataGridView3.Rows.Clear();
                            dataGridView4.Rows.Clear();
                            timer2.Start();
                            timer2.Enabled = true;
                            button11.Enabled = true;
                        }
                    }
                    break;
                case 3:
                    if (scan_mass != null)
                    {
                        Application.DoEvents();
                        dataGridView5.Rows.Add(dataGridView5.Rows.Count, "Образец " + dataGridView5.Rows.Count);
                        Array.Resize<double[]>(ref massGEMultiAbs, dataGridView5.Rows.Count - 1);
                        massGEMultiAbs[massGEMultiAbs.Length - 1] = new double[dataGridView5.ColumnCount - 2];
                        Array.Resize<double[]>(ref massGEMultiT, dataGridView5.Rows.Count - 1);
                        massGEMultiT[massGEMultiAbs.Length - 1] = new double[dataGridView5.ColumnCount - 2];
                        StopSpectr = false;
                        button14.Enabled = false;
                        button11.Enabled = true;
                        TableMultiScan();
                        Application.DoEvents();
                    }
                    else
                    {
                        MessageBox.Show("Вы забыли откалиброваться! Откалибруйтесь!");
                    }
                    break;
            }
        }
        public double[][] massGEMultiAbs;
        public double[][] massGEMultiT;
        public void ChartGraf()
        {
          ///  chart3.Series[0].Points.Clear();
         //   chart3.Series[1].Points.Clear();
         //   chart3.ChartAreas[0].AxisX.Interval = interval;
            if (radioButton1.Checked == true)
            {
                TableKinetica1.Columns[1].HeaderText = "Abs";
                TableKinetica1.Columns[2].HeaderText = "%T";
                dataGridView1.Columns[1].HeaderText = "Abs";
                dataGridView1.Columns[2].HeaderText = "%T";
                dataGridView2.Columns[1].HeaderText = "Abs";
                dataGridView2.Columns[2].HeaderText = "%T";
            }
            else
            {
                TableKinetica1.Columns[2].HeaderText = "Abs";
                TableKinetica1.Columns[1].HeaderText = "%T";
                dataGridView1.Columns[1].HeaderText = "%T";
                dataGridView1.Columns[2].HeaderText = "Abs";
                dataGridView2.Columns[1].HeaderText = "%T";
                dataGridView2.Columns[2].HeaderText = "Abs";
            }
            if (TableKinetica1.Columns[1].HeaderText == "Abs")
            {
                //Array.Sort(massGE);
                chart3.ChartAreas[0].AxisY.Minimum = 0;
                chart3.ChartAreas[0].AxisY.Maximum = 3;
                chart3.ChartAreas[0].AxisX.Minimum = 0;
                chart3.ChartAreas[0].AxisX.Maximum = start;
            }
            else
            {
                //Array.Sort(massGE);
                chart3.ChartAreas[0].AxisY.Minimum = 0;
                chart3.ChartAreas[0].AxisY.Maximum = 125;
                chart3.ChartAreas[0].AxisX.Minimum = 0;
                chart3.ChartAreas[0].AxisX.Maximum = start;
            }
          /*  chart3.Series.Add("area" + countButtonClick);
            Random r = new Random();
            int x = r.Next(100, 200), y = r.Next(100, 200);
            chart3.Series["area" + countButtonClick].Color =
                System.Drawing.Color.FromArgb(
                    (byte)(x - 2 * y),
                    (byte)(y + x),
                    (byte)(y - 2 * x));*/
        }
        public void TimerTick1(object sender, EventArgs e)
        {
            label53.Text = Convert.ToString(delay);
            delay--;
            if(delay < 0.0)
            {
                timer1.Stop();
                timer1.Enabled = false;
                timeLeft = Convert.ToInt32(start);
                TableKinetica(sender, e);
                timer2.Start();
                timer2.Enabled = true;
                button11.Enabled = true;
            }
        }
        public void TableScan_Save()
        {
            if (countButtonClick == 1)
            {
                listBox1.Items.Add("Измерение" + countButtonClick);
                countScan[countButtonClick - 1] = new string[ScanTable.Rows.Count - 1, 3];
                ScanTable_Save();
                Application.DoEvents();
                ScanChart.Series.Add("area" + countButtonClick);
                Random r = new Random();
                int x = r.Next(100, 200), y = r.Next(100, 200);
                ScanChart.Series["area" + countButtonClick].Color =
                    System.Drawing.Color.FromArgb(
                        (byte)(x - 2 * y),
                        (byte)(y + x),
                        (byte)(y - 2 * x));
                //   TableScan();
                Application.DoEvents();

            }
            else
            {
                listBox1.Items.Add("Измерение" + countButtonClick);
                Array.Resize<string[,]>(ref countScan, countButtonClick);
                countScan[countButtonClick - 1] = new string[ScanTable.Rows.Count - 1, 3];
                ScanTable_Save();
                Application.DoEvents();
                ScanChart.Series.Add("area" + countButtonClick);
                Random r = new Random();
                int x = r.Next(0, 200), y = r.Next(0, 200);
                ScanChart.Series["area" + countButtonClick].Color =
                    System.Drawing.Color.FromArgb(
                        (byte)(x - 2 * y),
                        (byte)(y + x),
                        (byte)(y - 2 * x));
                //   TableScan();
                Application.DoEvents();

            }
        }
        public void ScanTable_Save()
        {

            for (int i = 0; i < ScanTable.Rows.Count - 1; i++)
            {
                for (int k = 0; k < ScanTable.ColumnCount; k++)
                {
                    countScan[countButtonClick - 1][i, k] = ScanTable.Rows[i].Cells[k].Value.ToString();
                    
                }
            }
            countButtonClick++;
            

        }
        public void TableAgro2()
        {
            double sredPlot = 0;
            int length = 1;
            double[] numbers = new double[length];
         
            while (StopAgro != true)
            {
                Thread.Sleep(1000);
                Application.DoEvents();
                string GE5Izmer = "";
                string GE5_1_1 = "";
                double serValue = 0;
                double SredValue = 0;
                while (GE5Izmer == "")
                {

                    GE5Izmer = "";
                    GE5_1_1 = "";
                    newPort.Write("SA " + countSA + "\r");
                    string indata = newPort.ReadExisting();
                    string indata_0;
                    bool indata_bool = true;
                    while (indata_bool == true)
                    {
                        if (indata.Contains(">"))
                        {
                            indata_bool = false;
                        }

                        else
                        {
                            indata = newPort.ReadExisting();
                        }
                    }

                    newPort.Write("GE 1\r");
                    // Thread.Sleep(500);
                    // GEbyteRecieved4_1 = newPort.ReadBufferSize;
                    //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                    // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
                    // Thread.SpinWait(500);
                    indata_0 = "";
                    for (int i = 0; i <= 5000000; i++)
                    {
                        indata = newPort.ReadExisting();
                        if (indata_0.Contains("\r>"))
                        {
                            break;
                        }
                        indata_0 += indata;
                    }

                    //  indata_0 = "";
                    indata_bool = true;
                    /* while (indata_bool == true)
                     {

                         if (indata.Contains(">"))
                         {
                             indata_0 = indata;
                             indata_bool = false;

                         }
                         else {                   

                                 indata = newPort.ReadExisting();
                                 indata_0 += indata;

                         }
                     }*/
                    Regex regex = new Regex(@"\W");
                    Regex regex1 = new Regex(@"\D");
                    GE5Izmer = regex.Replace(indata_0, "");
                    GE5Izmer = regex1.Replace(GE5Izmer, "");
                }
                //MessageBox.Show("Измерение");
                GEText.Text = GE5Izmer;
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                double OptPlot1 = 0;

                OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
                double OptPlot1_1 = OptPlot1;
                Application.DoEvents();
                //if (OptPlot1_1 > 0.002)
                //{
                numbers[length - 1] = OptPlot1_1;
                length++;
                Array.Resize<double>(ref numbers, length);
                    
                    
                    
                if (numbers.Length > 6)
                {
                    if (numbers[numbers.Length - 2] / numbers[numbers.Length - 3] > 1.4
                        && numbers[numbers.Length - 2] / numbers[numbers.Length - 4] > 1.4
                        && numbers[numbers.Length - 2] / numbers[numbers.Length - 5] > 1.4
                        && numbers[numbers.Length - 2] / numbers[numbers.Length - 6] > 1.4
                        )
                    //  sredPlot = 0;
                    // double[] numbers2 = numbers.ToArray<double>();
                    /* for (int i = 0; i < numbers.Length; i++)
                     {
                         sredPlot += numbers[i];
                     }*/
                    // sredPlot = sredPlot / numbers.Length;
                    //if (sredPlot > 0)
                    // {
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value = string.Format("{0:0.0000}", numbers[numbers.Length-3]);
                        if (aproksim == "Линейная через 0")
                        {

                            if (Table2.Rows[0].Cells[2].Value.ToString() != "" &&
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value != null)
                            {
                                if ((Convert.ToDouble(Table2.Rows[0].Cells[2].Value.ToString()) >
                                    Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString())) && count == 0)
                                {
                                    if (count == 0)
                                    {
                                        count++;
                                        MessageBox.Show("Оптическая плотность контрольногго образца не может быть больше иззмеряемого!");
                                    }

                                }

                                serValue = (Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString()) -
                                    Convert.ToDouble(Table2.Rows[0].Cells[2].Value.ToString())) / Convert.ToDouble(AgroText1.Text);
                            }
                            else
                            {

                                serValue = 0;
                                if (Table2.Rows[0].Cells[2].Value.ToString() == null)
                                {
                                    MessageBox.Show("Измерьте Контрольный образец!");
                                    return;


                                }
                            }
                        }
                        if (aproksim == "Линейная")
                        {
                            if (Table2.Rows[0].Cells[2].Value.ToString() != "" &&
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString() != "")
                            {
                                if ((Convert.ToDouble(Table2.Rows[0].Cells[2].Value.ToString()) >
                                    Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString()))
                                    && count == 0)
                                {
                                    if (count == 0)
                                    {
                                        count++;
                                        MessageBox.Show("Оптическая плотность контрольногго образца не может быть больше иззмеряемого!");
                                    }
                                }
                                serValue =
                                    ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString()) -
                                    Convert.ToDouble(Table2.Rows[0].Cells[2].Value.ToString()) -
                                    Convert.ToDouble(AgroText0.Text))) /
                                    Convert.ToDouble(AgroText1.Text);
                            }
                            else
                            {

                                serValue = 0;
                                if (Table2.Rows[0].Cells[2].Value.ToString() == null)
                                {
                                    MessageBox.Show("Измерьте Контрольный образец!");
                                    return;


                                }
                            }

                        }
                        if (aproksim == "Квадратичная")
                        {
                            if (Table2.Rows[0].Cells[2].Value.ToString() != "" &&
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString() != "")
                            {
                                if ((Convert.ToDouble(Table2.Rows[0].Cells[2].Value.ToString()) >
                                    Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString()))
                                    && count == 0)
                                {
                                    if (count == 0)
                                    {
                                        count++;
                                        MessageBox.Show("Оптическая плотность контрольногго образца не может быть больше иззмеряемого!");
                                    }
                                }
                                serValue =
                                    ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[2].Value.ToString())
                                    -
                                    Convert.ToDouble(Table2.Rows[0].Cells[2].Value.ToString()) -
                                    Convert.ToDouble(AgroText0.Text))) /
                                    (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                            }
                            else
                            {
                                serValue = 0;
                                if (Table2.Rows[0].Cells[2].Value.ToString() == null)
                                {
                                    MessageBox.Show("Измерьте Контрольый образец!");
                                    return;


                                }
                            }


                        }
                        double CValue1 = Convert.ToDouble(F1Text.Text);
                        double CValue2 = Convert.ToDouble(F2Text.Text);

                        if (serValue >= 0)
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells[3].Value =
                                string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                            SredValue +=
                                Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells[3].Value.ToString());
                        }
                        else
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells[3].Value = "";
                        }

                        Application.DoEvents();
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Obrazec"].Value =
                                                      "Образец " + Table2.CurrentCell.RowIndex;
                        Table2.Rows.Add();

                        Table2.Rows[Table2.CurrentCell.RowIndex + 1].ReadOnly = false;
                        Table2.Rows[Table2.CurrentCell.RowIndex + 1].Cells["Column1"].Value =
                        Table2.CurrentCell.RowIndex + 1;
                        Table2.Rows[Table2.CurrentCell.RowIndex + 1].Cells["Column1"].ReadOnly = true;
                        Table2.CurrentCell = this.Table2[2, Table2.CurrentCell.RowIndex + 1];
                        button14.Enabled = true;
                        button11.Enabled = true;
                        Application.DoEvents();
                    }
                    //  length = 1;
                    // numbers = new double[length];

                }
              
            }
        }
        public int countscan = 0;

        public void TableMultiScan()
        {
            countscan = 0;
            while ((countscan != dataGridView5.ColumnCount - 2) && (StopSpectr != true))
            {
                Application.DoEvents();
                string GE5Izmer = "";
                string GE5_1_1 = "";
                while (GE5Izmer == "")
                {
                    SW_MultiScan();
                    GE5Izmer = "";
                    GE5_1_1 = "";
                    newPort.Write("SA " + scan_massSA[countscan] + "\r");
                    string indata = newPort.ReadExisting();
                    string indata_0;
                    bool indata_bool = true;
                    while (indata_bool == true)
                    {
                        if (indata.Contains(">"))
                        {
                            indata_bool = false;
                        }

                        else
                        {
                            indata = newPort.ReadExisting();
                        }
                    }

                    newPort.Write("GE 1\r");
                    // Thread.Sleep(500);
                    // GEbyteRecieved4_1 = newPort.ReadBufferSize;
                    //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                    // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
                    // Thread.SpinWait(500);
                    indata_0 = "";
                    for (int i = 0; i <= 5000000; i++)
                    {
                        indata = newPort.ReadExisting();
                        if (indata_0.Contains("\r>"))
                        {
                            break;
                        }
                        indata_0 += indata;
                    }

                    //  indata_0 = "";
                    indata_bool = true;
                    /* while (indata_bool == true)
                     {

                         if (indata.Contains(">"))
                         {
                             indata_0 = indata;
                             indata_bool = false;

                         }
                         else {                   

                                 indata = newPort.ReadExisting();
                                 indata_0 += indata;

                         }
                     }*/
                    Regex regex = new Regex(@"\W");
                    Regex regex1 = new Regex(@"\D");
                    GE5Izmer = regex.Replace(indata_0, "");
                    GE5Izmer = regex1.Replace(GE5Izmer, "");
                }
                //MessageBox.Show("Измерение");
                GEText.Text = GE5Izmer;
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(scan_mass[countscan]) * 100;
                double OptPlot1 = 0;
               
                OptPlot1 = Math.Log10((Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) / 
                    (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
                double OptPlot1_1 = OptPlot1;
                Application.DoEvents();
                dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells["Abs " + countscan].Value = string.Format("{0:0.0000}", OptPlot1_1);
                massGEMultiAbs[dataGridView5.Rows.Count - 2][countscan] = 
                    Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells["Abs " + countscan].Value);

                massGEMultiT[dataGridView5.Rows.Count - 2][countscan] = (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) * 100;
                countscan++;
                Application.DoEvents();
            }
            button14.Enabled = true;
            button11.Enabled = false;
            Application.DoEvents();
            if(StopSpectr == true)
            {
                MessageBox.Show("Измерение было прервано!");
            }
        }
        public void TableScan()
        {
            for (int i = 0; i < ScanTable.Rows.Count - 1; i++)
            {
                for (int k = 0; k < ScanTable.ColumnCount; k++)
                {
                   if (k > 0)
                    {
                        ScanTable.Rows[i].Cells[k].Value = null;
                    }
                }
            }
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            double[] massWL = new double[ScanTable.Rows.Count - 1];
            double[] massGE = new double[ScanTable.Rows.Count - 1];
            countscan = 0;
          //  ScanChart.Series[0].Points.Clear();
           // ScanChart.Series[1].Points.Clear();
            //  ScanChart.ChartAreas[0].AxisX.MajorGrid.Interval = Convert.ToDouble(string.Format("{0:0.0}", ScanTable.Rows[0].Cells[0].Value)) - Convert.ToDouble(string.Format("{0:0.0}", ScanTable.Rows[1].Cells[0].Value));

            while ((countscan != ScanTable.Rows.Count - 1) && (StopSpectr != true) && label56.Text != "0")
            {
                string GE5Izmer = "";
                string GE5_1_1 = "";
                while (GE5Izmer == "")
                {
                    SW_Scan();
                    GE5Izmer = "";
                    GE5_1_1 = "";
                    newPort.Write("SA " + scan_massSA[countscan] + "\r");
                    string indata = newPort.ReadExisting();
                    string indata_0;
                    bool indata_bool = true;
                    while (indata_bool == true)
                    {
                        if (indata.Contains(">"))
                        {
                            indata_bool = false;
                        }

                        else
                        {
                            indata = newPort.ReadExisting();
                        }
                    }

                    newPort.Write("GE 1\r");
                    // Thread.Sleep(500);
                    // GEbyteRecieved4_1 = newPort.ReadBufferSize;
                    //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
                    // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
                    // Thread.SpinWait(500);
                    indata_0 = "";
                    for (int i = 0; i <= 5000000; i++)
                    {
                        indata = newPort.ReadExisting();
                        if (indata_0.Contains("\r>"))
                        {
                            break;
                        }
                        indata_0 += indata;
                    }

                    indata_bool = true;
                    
                    Regex regex = new Regex(@"\W");
                    Regex regex1 = new Regex(@"\D");
                    GE5Izmer = regex.Replace(indata_0, "");
                    GE5Izmer = regex1.Replace(GE5Izmer, "");
                }
             
                GEText.Text = GE5Izmer;
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(scan_mass[countscan]) * 100;
                double OptPlot1 = 0;
                //MessageBox.Show(RDstring[countSA]);
                OptPlot1 = Math.Log10((Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
                double OptPlot1_1 = OptPlot1;
                Application.DoEvents();
                massWL[countscan] = Convert.ToDouble(ScanTable.Rows[countscan].Cells[0].Value);
                if (ScanTable.Columns[1].HeaderText == "Abs")
                {
                    ScanTable.Rows[countscan].Cells[1].Value = string.Format("{0:0.0000}", OptPlot1_1);
                    massGE[countscan] = Convert.ToDouble(ScanTable.Rows[countscan].Cells[1].Value);
                    ScanTable.Rows[countscan].Cells[2].Value =
                        string.Format("{0:0.0}",
                        ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                        (Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) * 100));
                }
                else
                {
                    ScanTable.Rows[countscan].Cells[2].Value = string.Format("{0:0.0000}", OptPlot1_1);
                    ScanTable.Rows[countscan].Cells[1].Value =
                        string.Format("{0:0.0}",
                        ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                        (Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) * 100));
                    /** listBox1.Items.Add(GE5Izmer);*/
                    massGE[countscan] = Convert.ToDouble(ScanTable.Rows[countscan].Cells[1].Value);

                }
                massWL[countscan] = Convert.ToDouble(ScanTable.Rows[countscan].Cells[0].Value);
                massGE[countscan] = Convert.ToDouble(ScanTable.Rows[countscan].Cells[1].Value);
                Array.Sort(massWL);
                Array.Sort(massGE);
                double x1 = Convert.ToDouble(ScanTable.Rows[countscan].Cells[0].Value);
                double y1 = Convert.ToDouble(ScanTable.Rows[countscan].Cells[1].Value);
                /*ScanChart.Series[0].Points.AddXY(x1, y1);
                ScanChart.Series[0].ChartType = SeriesChartType.Point;
                ScanChart.ChartAreas[0].AxisY.Crossing = 0;
                ScanChart.ChartAreas[0].AxisX.Crossing = 0;*/
                
                ScanChart.Series[countButtonClick].Points.AddXY(x1, y1);
                ScanChart.Series[countButtonClick].ChartType = SeriesChartType.Line;
                if (ScanTable.Rows[countscan].Cells[1].Value != null && ScanTable.Rows[countscan].Cells[2].Value != null)
                {
                    ScanChart.ChartAreas[0].AxisX.Minimum = cancel;
                    ScanChart.ChartAreas[0].AxisX.Maximum = start;

                }
                ScanChart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                ScanChart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                ScanChart.ChartAreas[0].AxisX.Title = ScanTable.Columns["WalveDl"].HeaderText;
                ScanChart.ChartAreas[0].AxisY.Title = ScanTable.Columns["Abs_scan"].HeaderText;
                countscan++;
            }

            double max = 0.0;
            double min = 0.0;
            for (int i = 0; i < ScanTable.Rows.Count; i++)
            {
                if (i == 0)
                {
                    if (Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) > Convert.ToDouble(ScanTable.Rows[i + 1].Cells[1].Value))
                    {
                        max = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        dataGridView1.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                        min = max;
                        double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                        double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        ScanChart.Series[0].Points.AddXY(x1, y1);
                        ScanChart.Series[0].ChartType = SeriesChartType.Point;
                    }
                    else
                    {
                        min = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        dataGridView2.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                        max = min;
                        double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                        double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        ScanChart.Series[0].Points.AddXY(x1, y1);
                        ScanChart.Series[0].ChartType = SeriesChartType.Point;
                    }

                }
                else {
                    if (i + 1 != ScanTable.Rows.Count)
                    {
                        if (Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) > Convert.ToDouble(ScanTable.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) >= Convert.ToDouble(ScanTable.Rows[i + 1].Cells[1].Value)

                            )
                        {
                            max = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            min = max;
                            dataGridView1.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                            double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                            double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            ScanChart.Series[0].Points.AddXY(x1, y1);
                            ScanChart.Series[0].ChartType = SeriesChartType.Point;
                        }
                        if ((Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) < Convert.ToDouble(ScanTable.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) <= Convert.ToDouble(ScanTable.Rows[i + 1].Cells[1].Value))

                            )
                        {
                            min = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            dataGridView2.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                            max = min;
                            double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                            double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            ScanChart.Series[0].Points.AddXY(x1, y1);
                            ScanChart.Series[0].ChartType = SeriesChartType.Point;
                        }
                    }
                }
            }


            /* while (countscan != 0)
             {



                 countscan--;
             }*/
        }
        public void SW_Scan()
        {
            //LogoForm();
            Application.DoEvents();
            string SWText1 = ScanTable.Rows[countscan].Cells[0].Value.ToString();
            double Walve_double = Convert.ToDouble(ScanTable.Rows[countscan].Cells[0].Value.ToString().Replace(".", ","));
            newPort.Write("SW " + Walve_double.ToString(System.Globalization.CultureInfo.GetCultureInfo("en-US")) + "\r");
           // Thread.Sleep(100);
            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = newPort.ReadExisting();
                }
            }
            GWNew.Text = string.Format("{0:0.0}", ScanTable.Rows[countscan].Cells[0].Value.ToString());
            Application.DoEvents();
            // SWF.Application.OpenForms["LogoFrm"].Close();
         //  Thread.Sleep(100);
            // GW();
        }

        public void SW_MultiScan()
        {
            //LogoForm();
            Application.DoEvents();
            string SWText1 = textBoxCO[countscan].Text;
            double Walve_double = Convert.ToDouble(textBoxCO[countscan].Text.Replace(".", ","));
            newPort.Write("SW " + Walve_double.ToString(System.Globalization.CultureInfo.GetCultureInfo("en-US")) + "\r");
            // Thread.Sleep(100);
            string indata = newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = newPort.ReadExisting();
                }
            }
            GWNew.Text = string.Format("{0:0.0}", textBoxCO[countscan].Text);
            Application.DoEvents();
            // SWF.Application.OpenForms["LogoFrm"].Close();
            //  Thread.Sleep(100);
            // GW();
        }

        public void IzmerenieFr_izmer()
        {
            int startIndexCell = 3;
            int endIndexCell = 6;
            int rowIndex = IzmerenieFR_Table.CurrentRow.Index;

            bool doNotWrite = false;
            string SWAnalis = WL_grad1;
            string GE5Izmer = "";
            string GE5_1_1 = "";
            newPort.Write("SA " + countSA + "\r");
            string indata = newPort.ReadExisting();
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = newPort.ReadExisting();

                }
            }

            newPort.Write("GE 1\r");
            // Thread.Sleep(500);
            // GEbyteRecieved4_1 = newPort.ReadBufferSize;
            //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
            // Thread.SpinWait(500);
            indata_0 = "";
            for (int i = 0; i <= 5000000; i++)
            {
                indata = newPort.ReadExisting();
                if (indata_0.Contains("\r>"))
                {
                    break;
                }
                indata_0 += indata;
            }

            //  indata_0 = "";
            indata_bool = true;
            /* while (indata_bool == true)
             {

                 if (indata.Contains(">"))
                 {
                     indata_0 = indata;
                     indata_bool = false;

                 }
                 else {                   

                         indata = newPort.ReadExisting();
                         indata_0 += indata;

                 }
             }*/
            Regex regex = new Regex(@"\W");
            Regex regex1 = new Regex(@"\D");
            GE5Izmer = regex.Replace(indata_0, "");
            GE5Izmer = regex1.Replace(GE5Izmer, "");
            //  MessageBox.Show("Измерение: " + GE5Izmer + "\rКалибровка: " + GE5_1_0 + "\rОтклонение: " + Convert.ToString(Convert.ToDouble(GE5_1_0) / (Convert.ToDouble(GE5Izmer))) +
            //    "\rПоправочный коэффициент: " + RDstring[countSA]);


            GEText.Text = GE5Izmer;

            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;

            OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));



            double OptPlot1_1 = OptPlot1;
            IzmerenieFR_Table.Rows[rowIndex].Cells["Walve"].Value = string.Format("{0:0.0}", GWNew.Text);
            if (IzmerenieFR_Table.CurrentCell.ColumnIndex != 5)
            {
                if ((IzmerenieFR_Table.CurrentCell.ReadOnly != true && rowIndex != IzmerenieFR_Table.Rows.Count - 1) || IzmerenieFR_Table.CurrentCell.ColumnIndex == 3)
                {
                    IzmerenieFR_Table.Rows[rowIndex].Cells["ABS"].Value = string.Format("{0:0.0000}", OptPlot1_1);
                    string k1 = Convert.ToString(IzmerenieFR_Table.Rows[rowIndex].Cells["KOne"].Value);
                    k1 = k1.Replace(".", ",");
                    IzmerenieFR_Table.Rows[rowIndex].Cells["T"].Value = string.Format("{0:0.0}", ((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA]))) * 100);
                    IzmerenieFR_Table.Rows[rowIndex].Cells["Concetracia"].Value = string.Format("{0:0.0000}", (OptPlot1_1 * Convert.ToDouble(k1)));
                    int curentIndex = IzmerenieFR_Table.CurrentCell.ColumnIndex;
                    if (curentIndex != IzmerenieFR_Table.ColumnCount - 1 || rowIndex != IzmerenieFR_Table.Rows.Count - 1)
                    {
                        if (rowIndex != IzmerenieFR_Table.Rows.Count - 2)
                        {
                            IzmerenieFR_Table.CurrentCell = this.IzmerenieFR_Table[curentIndex, rowIndex + 1];
                        }
                        else
                        {
                            MessageBox.Show("Измерения были проведены!");
                        }
                    }

                }
                else
                {
                    MessageBox.Show("Запись запрещена!");
                }
            }
            else
            {
                MessageBox.Show("Производить измерения в данную ячейку запрещено! Только ручное изменение!");
            }
            
        }
        public void Graduirovka(object sender, EventArgs e)
        {
            
            double sum = 0.0;
            int startIndexCell = 2;
            int endIndexCell = startIndexCell + NoCaIzm;
            int rowIndex = Table1.CurrentRow.Index;

            bool doNotWrite = false;
            string SWAnalis = WL_grad1;
            string GE5Izmer = "";
            string GE5_1_1 = "";
            newPort.Write("SA " + countSA + "\r");
            string indata = newPort.ReadExisting();
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = newPort.ReadExisting();

                }
            }

            newPort.Write("GE 1\r");
            // Thread.Sleep(500);
            // GEbyteRecieved4_1 = newPort.ReadBufferSize;
            //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
            // Thread.SpinWait(500);
            indata_0 = "";
            for (int i = 0; i <= 5000000; i++)
            {
                indata = newPort.ReadExisting();
                if (indata_0.Contains("\r>"))
                {
                    break;
                }
                indata_0 += indata;
            }

            //  indata_0 = "";
            indata_bool = true;
            /* while (indata_bool == true)
             {

                 if (indata.Contains(">"))
                 {
                     indata_0 = indata;
                     indata_bool = false;

                 }
                 else {                   

                         indata = newPort.ReadExisting();
                         indata_0 += indata;

                 }
             }*/
            Regex regex = new Regex(@"\W");
            Regex regex1 = new Regex(@"\D");
            GE5Izmer = regex.Replace(indata_0, "");
            GE5Izmer = regex1.Replace(GE5Izmer, "");
            //  MessageBox.Show("Измерение: " + GE5Izmer + "\rКалибровка: " + GE5_1_0 + "\rОтклонение: " + Convert.ToString(Convert.ToDouble(GE5_1_0) / (Convert.ToDouble(GE5Izmer))) +
            //    "\rПоправочный коэффициент: " + RDstring[countSA]);


            GEText.Text = GE5Izmer;
            
            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;
            
                OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));

               
            
            double OptPlot1_1 = OptPlot1;
            if (Table1.CurrentCell.ReadOnly != true)
            {
                Table1.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);

                int curentIndex = Table1.CurrentCell.ColumnIndex;
                if (curentIndex != Table1.ColumnCount - 1 || rowIndex != Table1.Rows.Count - 2)
                {
                    if (rowIndex != Table1.Rows.Count - 2)
                    {
                        Table1.CurrentCell = this.Table1[curentIndex, rowIndex + 1];
                    }
                    else
                    {
                        Table1.CurrentCell = this.Table1[curentIndex + 1, 0];
                    }
                }

            }
            else
            {
                MessageBox.Show("Запись запрещена!");
            }
            GAText.Text = string.Format("{0:0.00}", Aser);
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                {
                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;

                            for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                            {
                                if (Table1.Rows[rowIndex].Cells[l].Value == null)
                                {
                                    cellnull++;
                                    // Table1.Rows[rowIndex].Cells[Table1.CurrentCell.ColumnIndex + 1].Selected = true;

                                }
                            }
                        }


                    }
                }
            }



            if (!doNotWrite)
            {
                if (NoCaSer == 1)
                {
                    radioButton1.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                    radioButton2.Enabled = false;
                }
                if (NoCaSer == 2)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                }
                if (NoCaSer >= 3)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                }
                sum = 0.0;
                /* while (true)
                 {
                     int i = Table1.Columns.Count - 1;//С какого столбца начать
                     if (Table1.Columns.Count == 3 + Convert.ToInt32(CountSeriya2))
                         break;
                     Table1.Columns.RemoveAt(i);
                 }*/

                for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                {
                    if (Table1.Rows[rowIndex].Cells[l].Value == null)
                    {
                        cellnull++;
                    }

                    else
                    {
                        for (int j = 0; j < Table1.Rows.Count - 1; j++)
                        {

                            for (int i1 = startIndexCell + 1; i1 <= endIndexCell; ++i1)
                            {
                                sum += Convert.ToDouble(Table1.Rows[j].Cells[i1].Value);
                                Asred1 = sum / NoCaIzm;
                                // MessageBox.Show(Convert.ToString(Asred1));
                                Table1.Rows[j].Cells["Asred"].Value = string.Format("{0:0.0000}", Asred1);

                            }
                            sum = 0.0;
                        }
                    }
                    Izmerenie1 = true;
                }
                for (int m = 0; m < Table1.Rows.Count - 1; m++)
                {
                    for (int ml = 0; ml < Table1.Rows[m].Cells.Count; ml++)
                    {
                        if (Table1.Rows[m].Cells[ml].Value == null)
                        { doNotWrite = true; }
                    }
                }
                if (!doNotWrite)
                {
                    while (true)
                    {
                        int ml = Table1.Columns.Count - 1;//С какого столбца начать
                        if (Table1.Columns.Count == 3 + NoCaIzm)
                            break;
                        Table1.Columns.RemoveAt(ml);
                    }
                    functionAsred();

                }

            }
        }

        public void Izmerenie(object sender, EventArgs e)
        {
            double CCR = 0.0;
            int rowIndex2 = Table2.CurrentRow.Index;
            bool doNotWrite1 = false;
            double maxEl;
            double minEl;
            double serValue = 0;
            int cellnull = 0;
            El = new double[NoCaIzm1];
            string GE5Izmer = "";
            string GE5_1_1 = "";
            newPort.Write("SA " + countSA + "\r");
            string indata = newPort.ReadExisting();
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {
                    indata_bool = false;
                }
                else {
                    indata = newPort.ReadExisting();
                }
            }
            newPort.Write("GE 1\r");
            // Thread.Sleep(500);
            // GEbyteRecieved4_1 = newPort.ReadBufferSize;
            //  GEbuffer4_1 = new byte[GEbyteRecieved4_1];
            // MessageBox.Show(newPort.Read(GEbuffer4_1, 0, GEbyteRecieved4_1).ToString());
            // Thread.SpinWait(500);
            indata_0 = "";
            for (int i = 0; i <= 5000000; i++)
            {
                indata = newPort.ReadExisting();
                if (indata_0.Contains("\r>"))
                {
                    break;
                }
                indata_0 += indata;
            }
            int indata_zero = 0;
            //  indata_0 = "";
            indata_bool = true;
            /* while (indata_bool == true)
             {

                 if (indata.Contains(">"))
                 {
                     indata_0 = indata;
                     indata_bool = false;

                 }
                 else {                   

                         indata = newPort.ReadExisting();
                         indata_0 += indata;

                 }
             }*/
            GE5Izmer = "";
          Regex  regex = new Regex(@"\W");
           Regex regex1 = new Regex(@"\D");
            GE5Izmer = regex.Replace(indata_0, "");
            GE5Izmer = regex1.Replace(GE5Izmer, "");

            GEText.Text = GE5Izmer;
            // MessageBox.Show(GE5Izmer);
            int curentIndex = Table2.CurrentCell.ColumnIndex;
            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / 
                (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
            double OptPlot1_1 = OptPlot1 - Math.Truncate(OptPlot1);            
            if (selet_rezim == 6)            {
                Table2.Rows[0].ReadOnly = true;
                if (Table2.Rows[Table2.CurrentCell.RowIndex].ReadOnly == true)
                {
                    Table2.CurrentCell = this.Table2[2, Table2.CurrentCell.RowIndex + 1];
                }
            }
            else {
                if (Table2.CurrentCell.ReadOnly != true)

                {

                    Table2.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                }

                else

                {

                    MessageBox.Show("Запись запрещена!");

                }
            }


            GAText.Text = string.Format("{0:0.00}", Aser);


            bool doNotWrite = false;
            double SredValue = 0;

            if (USE_KO == false)
            {
                if (Table2.CurrentCell.Value == null)
                {
                    El = new double[NoCaIzm1];

                    doNotWrite = true;


                    // El = new double[NoCaIzm1 + 1];
                    //for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    //{
                    // El = new double[NoCaIzm1 + 1];
                    SredValue = 0;
                    for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                    {
                        if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (aproksim == "Линейная через 0")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }

                            }
                            if (aproksim == "Линейная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            if (aproksim == "Квадратичная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            double CValue1 = Convert.ToDouble(F1Text.Text);
                            double CValue2 = Convert.ToDouble(F2Text.Text);
                            if (serValue >= 0)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                SredValue += Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                            else
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = "";
                            }
                            CCR = SredValue / NoCaIzm1;
                            if (Convert.ToDouble(textBox7.Text) >= 1)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.00}", ((CCR * Convert.ToDouble(textBox7.Text)) / 100));
                            }
                            else Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                            //Table2.Rows[j].Cells["d%"].Value = El.Max();
                            //  El[i1] = Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());
                        }
                        //El = new double[NoCaIzm1 + 1];
                        if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString() != "")
                            {
                                El[i1 - 1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                        }
                    }

                    Array.Sort(El);
                    maxEl = El[El.Length - 1];
                    minEl = El[0];
                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                    double b = a;
                    // b = b * 10;


                    if (minEl == 0)
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                    }
                    else
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                    }
                }

            }
            else
            {
                for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                {
                    Table2.Rows[0].Cells["C,edconctr;Ser." + i1].ReadOnly = true;
                    if (selet_rezim == 2)
                    {
                        Table2.Rows[0].Cells["Ccr"].ReadOnly = true;
                        Table2.Rows[0].Cells["d%"].ReadOnly = true;
                    }
                }
                if (Table2.CurrentCell.Value == null && Table2.CurrentCell.ReadOnly != true)
                {
                    Table2_UseCo();
                }
            }

            if (!doNotWrite)
            {
                if (USE_KO == true)
                {
                    Table2_UseCo();
                }

                else
                {
                    for (int i = 1; i <= NoCaIzm1; i++)
                    {
                        Table2.Rows[0].Cells["C,edconctr;Ser." + i].ReadOnly = true;
                        Table2.Rows[0].Cells["Ccr"].ReadOnly = true;
                        Table2.Rows[0].Cells["d%"].ReadOnly = true;
                    }
                    El = new double[NoCaIzm1];

                    // for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    // {
                    SredValue = 0;
                    for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                    {
                        if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (aproksim == "Линейная через 0")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }

                            }
                            if (aproksim == "Линейная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            if (aproksim == "Квадратичная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            double CValue1 = Convert.ToDouble(F1Text.Text);
                            double CValue2 = Convert.ToDouble(F2Text.Text);

                            if (serValue >= 0)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                SredValue += Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                            else
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = "";
                            }
                            CCR = SredValue / NoCaIzm1;
                            if (Convert.ToDouble(textBox7.Text) >= 1)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.00}", ((CCR * Convert.ToDouble(textBox7.Text)) / 100));
                            }
                            else Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                            //Table2.Rows[j].Cells["d%"].Value = El.Max();
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString() != "")
                                {
                                    El[i1 - 1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                                }
                            }
                        }

                    }

                    Array.Sort(El);
                    maxEl = El[El.Length - 1];
                    minEl = El[0];
                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                    double b = a;


                    if (minEl == 0)
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                    }
                    else
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                    }
                    // return;

                    /* for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                     {
                         Table2.Rows[0].Cells["C,edconctr;Ser." + i1].Value = "";
                         Table2.Rows[0].Cells["Ccr"].Value = "";
                         Table2.Rows[0].Cells["d%"].Value = "";
                     }*/
                }
            }
            if ((curentIndex != Table2.ColumnCount - 2 || Table2.CurrentCell.RowIndex != Table2.Rows.Count - 2) && Table2.CurrentCell.ReadOnly != true)

            {

                if (Table2.CurrentCell.RowIndex != Table2.Rows.Count - 2)

                {

                    Table2.CurrentCell = this.Table2[curentIndex, Table2.CurrentCell.RowIndex + 1];

                }

                else

                {

                    Table2.CurrentCell = this.Table2[curentIndex + 2, 0];

                }

            }

            else

            {

                Table2.CurrentCell = this.Table2[2, 0];

            }
        

    }
        public void ZapicInTable2()
        {
            bool doNotWrite = false;
            double sum = 0.0;
            int rowIndex = Table2.CurrentRow.Index;
            int curentIndex = Table2.CurrentCell.ColumnIndex;
            double CCR = 0.0;
            double maxEl;
            double minEl;
            double serValue = 0;
            int cellnull = 0;
            if (Table2.CurrentCell.ColumnIndex > 1)
            {
                InputBox _InputBox = new InputBox(this);
                _InputBox.ShowDialog();
                if (Table2.CurrentCell.ReadOnly != true)
                {
                    Table2.CurrentCell.Value = string.Format("{0:0.0000}", CellOpt);
                    CellOpt = 0;
                }
                else
                {
                    MessageBox.Show("Запись запрещена!");
                }
            }
            PodschetTable2();
            if (curentIndex != Table2.ColumnCount - 1 || rowIndex != Table2.Rows.Count - 2)
            {
                if (rowIndex != Table2.Rows.Count - 2)
                {
                    Table2.CurrentCell = this.Table2[curentIndex, rowIndex + 1];
                }
                else
                {
                    Table2.CurrentCell = this.Table2[curentIndex + 2, 0];
                }
            }
        }
    public void PodschetTable2()
        {
            bool doNotWrite = false;
            double sum = 0.0;
            int rowIndex = Table2.CurrentRow.Index;
            int curentIndex = Table2.CurrentCell.ColumnIndex;
            double CCR = 0.0;
            double maxEl;
            double minEl;
            double serValue = 0;
            int cellnull = 0;
            if (USE_KO == false)
            {
                if (Table2.CurrentCell.Value == null)
                {
                    El = new double[NoCaIzm1];

                    doNotWrite = true;


                    // El = new double[NoCaIzm1 + 1];
                    //for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    //{
                    // El = new double[NoCaIzm1 + 1];
                    double SredValue = 0;
                    for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                    {
                        if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (aproksim == "Линейная через 0")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }

                            }
                            if (aproksim == "Линейная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            if (aproksim == "Квадратичная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            double CValue1 = Convert.ToDouble(F1Text.Text);
                            double CValue2 = Convert.ToDouble(F2Text.Text);
                            if (serValue >= 0)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                SredValue += Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                            else
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = "";
                            }
                            CCR = SredValue / NoCaIzm1;
                            if (Convert.ToDouble(textBox7.Text) >= 1)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.00}", ((CCR * Convert.ToDouble(textBox7.Text)) / 100));
                            }
                            else Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                            //Table2.Rows[j].Cells["d%"].Value = El.Max();
                            //  El[i1] = Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());
                        }
                        //El = new double[NoCaIzm1 + 1];
                        if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString() != "")
                            {
                                El[i1 - 1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                        }
                    }

                    Array.Sort(El);
                    maxEl = El[El.Length - 1];
                    minEl = El[0];
                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                    double b = a;
                    // b = b * 10;


                    if (minEl == 0)
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                    }
                    else
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                    }
                }

            }
            else
            {
                for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                {
                    Table2.Rows[0].Cells["C,edconctr;Ser." + i1].ReadOnly = true;
                    Table2.Rows[0].Cells["Ccr"].ReadOnly = true;
                    Table2.Rows[0].Cells["d%"].ReadOnly = true;
                }
                if (Table2.CurrentCell.Value == null && Table2.CurrentCell.ReadOnly != true)
                {
                    El = new double[NoCaIzm1];

                    doNotWrite = true;


                    // El = new double[NoCaIzm1 + 1];
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        // El = new double[NoCaIzm1 + 1];
                        double SredValue = 0;
                        for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                        {
                            if (Table2.Rows[j].Cells["A;Ser" + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                if (aproksim == "Линейная через 0")
                                {
                                    if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() != "" && Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString() != "")
                                    {
                                        serValue = (Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString())) / Convert.ToDouble(AgroText1.Text);
                                    }
                                    else
                                    {

                                        serValue = 0;
                                        if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() == null)
                                        {
                                            MessageBox.Show("Измерьте Контрольый образец!");
                                            return;


                                        }
                                    }
                                }
                                if (aproksim == "Линейная")
                                {
                                    if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() != "" && Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString() != "")
                                    {
                                        serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / Convert.ToDouble(AgroText1.Text);
                                    }
                                    else
                                    {

                                        serValue = 0;
                                        if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() == null)
                                        {
                                            MessageBox.Show("Измерьте Контрольный образец!");
                                            return;


                                        }
                                    }

                                }
                                if (aproksim == "Квадратичная")
                                {
                                    if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() != "" && Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString() != "")
                                    {
                                        serValue = ((Convert.ToDouble(Table2.Rows[j].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                                    }
                                    else
                                    {
                                        serValue = 0;
                                        if (Table2.Rows[0].Cells["A;Ser" + i1].Value.ToString() == null)
                                        {
                                            MessageBox.Show("Измерьте Контрольный образец!");
                                            return;


                                        }
                                    }
                                }
                                double CValue1 = Convert.ToDouble(F1Text.Text);
                                double CValue2 = Convert.ToDouble(F2Text.Text);

                                if (serValue >= 0)
                                {
                                    Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                    SredValue += Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());
                                }
                                else
                                {
                                    Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value = "";
                                }
                                CCR = SredValue / NoCaIzm1;
                                if (Convert.ToDouble(textBox7.Text) >= 1)
                                {
                                    Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.00}", ((CCR * Convert.ToDouble(textBox7.Text)) / 100));
                                }
                                else Table2.Rows[j].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                                //Table2.Rows[j].Cells["d%"].Value = El.Max();
                                //  El[i1] = Convert.ToDouble(Table2.Rows[j].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                            //El = new double[NoCaIzm1 + 1];
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString() != "")
                                {
                                    El[i1 - 1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                                }
                            }
                        }

                        Array.Sort(El);
                        maxEl = El[El.Length - 1];
                        minEl = El[0];
                        double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                        double b = a;
                        // b = b * 10;


                        if (minEl == 0)
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                        }
                        else
                        {
                            Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                        }

                    }
                }
                for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                {
                    Table2.Rows[0].Cells["C,edconctr;Ser." + i1].Value = "";
                    Table2.Rows[0].Cells["Ccr"].Value = "";
                    Table2.Rows[0].Cells["d%"].Value = "";
                }
            }


            if (!doNotWrite)
            {
                if (USE_KO == true)
                {

                    Table2_UseCo();


                }

                else
                {
                    for (int i = 1; i <= NoCaIzm1; i++)
                    {
                        Table2.Rows[0].Cells["C,edconctr;Ser." + i].ReadOnly = true;
                        Table2.Rows[0].Cells["Ccr"].ReadOnly = true;
                        Table2.Rows[0].Cells["d%"].ReadOnly = true;
                    }
                    El = new double[NoCaIzm1];

                    // for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    // {
                    double SredValue = 0;
                    for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                    {
                        if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value == null)
                        {
                            cellnull++;
                        }
                        else
                        {
                            if (aproksim == "Линейная через 0")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {

                                    serValue = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }

                            }
                            if (aproksim == "Линейная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / Convert.ToDouble(AgroText1.Text);
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            if (aproksim == "Квадратичная")
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString() != "")
                                {
                                    serValue = ((Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["A;Ser" + i1].Value.ToString()) - Convert.ToDouble(AgroText0.Text))) / (Convert.ToDouble(AgroText1.Text) + Convert.ToDouble(AgroText2.Text));
                                }
                                else
                                {
                                    serValue = 0;
                                }
                            }
                            double CValue1 = Convert.ToDouble(F1Text.Text);
                            double CValue2 = Convert.ToDouble(F2Text.Text);

                            if (serValue >= 0)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = string.Format("{0:0.0000}", serValue * CValue1 * CValue2);
                                SredValue += Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                            }
                            else
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value = "";
                            }
                            CCR = SredValue / NoCaIzm1;
                            if (Convert.ToDouble(textBox7.Text) >= 1)
                            {
                                Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR) + "±" + string.Format("{0:0.00}", ((CCR * Convert.ToDouble(textBox7.Text)) / 100));
                            }
                            else Table2.Rows[Table2.CurrentCell.RowIndex].Cells["Ccr"].Value = string.Format("{0:0.0000}", CCR);
                            //Table2.Rows[j].Cells["d%"].Value = El.Max();
                            if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value == null)
                            {
                                cellnull++;
                            }
                            else
                            {
                                if (Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString() != "")
                                {
                                    El[i1 - 1] = Convert.ToDouble(Table2.Rows[Table2.CurrentCell.RowIndex].Cells["C,edconctr;Ser." + i1].Value.ToString());
                                }
                            }
                        }

                    }

                    Array.Sort(El);
                    maxEl = El[El.Length - 1];
                    minEl = El[0];
                    double a = ((maxEl - minEl) * 100) / Convert.ToDouble(CCR);
                    double b = a;


                    if (minEl == 0)
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = 0.0000;
                    }
                    else
                    {
                        Table2.Rows[Table2.CurrentCell.RowIndex].Cells["d%"].Value = string.Format("{0:0.00}", b);

                    }
                    // return;

                    /* for (int i1 = 1; i1 <= NoCaIzm1; i1++)
                     {
                         Table2.Rows[0].Cells["C,edconctr;Ser." + i1].Value = "";
                         Table2.Rows[0].Cells["Ccr"].Value = "";
                         Table2.Rows[0].Cells["d%"].Value = "";
                     }*/
                }
            }

      
            doNotWrite = false;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {

                for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                {
                    if (Table2.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                button3.Enabled = true;
                button9.Enabled = true;
            }
        }
        private void измеритьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 0)
            {
                if (Table1.RowCount > 1)
                {
                    Graduirovka(sender, e);
                }
                else
                {
                    MessageBox.Show("Создайте градуировку по СО");
                }
            }
            else
            {
                if (Table2.RowCount > 1)
                {
                    Izmerenie(sender, e);
                }
                else
                {
                    MessageBox.Show("Создайте измерение");
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            Calibrovka();
        }
        public void Calibrovka()
        {
            if (selet_rezim == 5)
            {

                StopSpectr = false;
                countscan = 0;
                scan_massSA = new double[ScanTable.Rows.Count - 1];
                scan_mass = new double[ScanTable.Rows.Count - 1];
                Application.DoEvents();
                LogoForm();
                while ((countscan != ScanTable.Rows.Count - 1) && (StopSpectr != true))                {

                    Application.DoEvents();

                    Application.DoEvents();
                    GWNew.Text = string.Format("{0:0.0}", ScanTable.Rows[countscan].Cells[0].Value.ToString());
                    SW_Scan();
                    Thread.Sleep(50);
                    SAGEScan(ref countSA, ref GE5_1_0);
                    Application.DoEvents();


                    countscan++;
                }
                SWF.Application.OpenForms["LogoFrm"].Close();
                MessageBox.Show("Калибровка закончена!");
                Podskazka.Text = "Можно сканировать";
                label59.Visible = false;
                label28.Visible = true;
            }
            else
            {
                if (selet_rezim == 3)
                {
                    StopSpectr = false;
                    countscan = 0;
                    scan_massSA = new double[dataGridView5.Columns.Count - 2];
                    scan_mass = new double[dataGridView5.Columns.Count - 2];
                    Application.DoEvents();
                    LogoForm();
                    while ((countscan != dataGridView5.Columns.Count - 2) && (StopSpectr != true))
                    {

                        Application.DoEvents();
                        SW_MultiScan();
                        Application.DoEvents();
                        GWNew.Text = string.Format("{0:0.0}", textBoxCO[countscan].Text);
                        button12.Enabled = false;
                        button14.Enabled = false;
                        button11.Enabled = true;
                        SAGEScan(ref countSA, ref GE5_1_0);
                        Application.DoEvents();
                        countscan++;
                    }
                    SWF.Application.OpenForms["LogoFrm"].Close();
                    button12.Enabled = true;
                    button14.Enabled = true;
                    button11.Enabled = false;
                    MessageBox.Show("Калибровка закончена!");
                }
                else
                {
                    SAGE(ref countSA, ref GE5_1_0);
                }
            }

        }

        private void калибровкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Calibrovka();
        }


        private void Table2_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {

        }

               
       

        private void Podskazka_TabIndexChanged(object sender, EventArgs e)
        {
         
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Table2.RowCount == 0)
            {
                k0_1 = Convert.ToDouble(AgroText0.Text);
                k1_1 = Convert.ToDouble(AgroText1.Text);
                k2_1 = Convert.ToDouble(AgroText2.Text);
                USE_KO_1 = USE_KO;
                button11.Enabled = false;
            }
            if (tabControl2.SelectedIndex == 1 && Table2.Rows.Count == 0)
            {
                Podskazka.Text = "Создайте или откройте Измерение!";
                label27.Visible = false;
                label24.Visible = false;
                label25.Visible = true;
                label26.Visible = true;
                label28.Visible = false;
                label33.Visible = false;
                if (Table2.RowCount > 0 && (k0_1 != Convert.ToDouble(AgroText0.Text) || k1_1 != Convert.ToDouble(AgroText1.Text) || k2_1 != Convert.ToDouble(AgroText2.Text) || USE_KO_1 != USE_KO))
                {
                    MessageBox.Show("Внимание: Грдуировка была изменена! Таблица Измерений будет пересчитана по новым коэффициентам!");
                    k0_1 = Convert.ToDouble(string.Format("{0:0.0000}", AgroText0.Text));
                    k1_1 = Convert.ToDouble(string.Format("{0:0.0000}", AgroText1.Text));
                    k2_1 = Convert.ToDouble(string.Format("{0:0.0000}", AgroText2.Text));
                    if (USE_KO_1 != USE_KO && USE_KO == false)
                    {
                        Table2.Rows.RemoveAt(0);
                        USE_KO_1 = USE_KO;
                    }
                    else
                    {
                        if (USE_KO_1 != USE_KO && USE_KO == true)
                        {
                            Table2.Rows.Insert(0, 0, "Контрольный");
                            for (int i = 1; i <= NoCaSer1; i++)
                            {
                                Table2.Rows[0].Cells["A;Ser" + i].Value = string.Format("{0:0.0000}", 0);

                            }
                            //   Table2.Rows.Add();
                            USE_KO_1 = USE_KO;
                            MessageBox.Show("Не забудьте измерить холостую пробу!");
                        }
                    }
                    PodschetTable2();
                }
            }
            else
            {
                if (tabControl2.SelectedIndex == 1 && Table2.Rows.Count > 0)
                {
                    Podskazka.Text = "Измеряйте или введите значения!";
                    label27.Visible = false;
                    label24.Visible = false;
                    label25.Visible = false;
                    label26.Visible = false;
                    label28.Visible = true;
                    label33.Visible = true;
                    if (Table2.RowCount > 0 && (k0_1 != Convert.ToDouble(AgroText0.Text) || k1_1 != Convert.ToDouble(AgroText1.Text) || k2_1 != Convert.ToDouble(AgroText2.Text) || USE_KO_1 != USE_KO))
                    {
                        MessageBox.Show("Внимание: Грдуировка была изменена! Таблица Измерений будет пересчитана по новым коэффициентам!");
                        k0_1 = Convert.ToDouble(string.Format("{0:0.0000}", AgroText0.Text));
                        k1_1 = Convert.ToDouble(string.Format("{0:0.0000}", AgroText1.Text));
                        k2_1 = Convert.ToDouble(string.Format("{0:0.0000}", AgroText2.Text));
                        if (USE_KO_1 != USE_KO && USE_KO == false)
                        {
                            Table2.Rows.RemoveAt(0);
                            USE_KO_1 = USE_KO;
                        }
                        else
                        {
                            if (USE_KO_1 != USE_KO && USE_KO == true)
                            {
                                Table2.Rows.Insert(0, 0, "Контрольный");
                                for (int i = 1; i <= NoCaSer1; i++)
                                {
                                    Table2.Rows[0].Cells["A;Ser" + i].Value = string.Format("{0:0.0000}", 0);

                                }
                                //   Table2.Rows.Add();
                                USE_KO_1 = USE_KO;
                                MessageBox.Show("Не забудьте измерить холостую пробу!");
                            }
                        }
                        PodschetTable2();
                    }
                }
                else
                {
                    Podskazka.Text = "Перейдите в измерения!";
                    label27.Visible = false;
                    label24.Visible = false;
                    label25.Visible = false;
                    label26.Visible = false;
                    label28.Visible = false;
                    label33.Visible = false;
                }
            }
        }


        const int initialValue = -1;   

        int curent = 3;
        int Table1_edit = 0;
        private void Table1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
          
                if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)
                {

                    if (Table1.CurrentCell.Value != "" && Table1.CurrentCell.Value != null)
                    {
                        CellOpt = Convert.ToDouble(Table1.CurrentCell.Value.ToString());
                    }


                Table1.CurrentCell.ReadOnly = true;
               // ZapicInTable1();
                }

            Table1.CurrentCell.ReadOnly = false;

        }

        private void tb_KeyPress(object sender, KeyPressEventArgs e)

        {

            char number = e.KeyChar;

            if (e.KeyChar <= 48 || e.KeyChar >= 58) //цифры, клавиша BackSpace и запятая а ASCII
            {

                e.Handled = true;
                MessageBox.Show("Только цифры!");
                CellOpt = Convert.ToDouble(e.KeyChar);
                return;              


            }

            return;

        }

        int Table2_edit = 0;
        private void Table2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
   
            
                if (Table2.CurrentCell.ColumnIndex >= 2 && Table2.CurrentCell.ReadOnly != true)
                {
                    if (Table2.CurrentCell.Value != "" && Table2.CurrentCell.Value != null)
                    {
                        CellOpt = Convert.ToDouble(Table2.CurrentCell.Value.ToString());
                    }
                Table2.CurrentCell.ReadOnly = true;
                // ZapicInTable1();
            }

            Table2.CurrentCell.ReadOnly = false;



        }



        private void Add_Table2_Click(object sender, EventArgs e)
        {
            if (Table2.RowCount != 0)
            {
                if (USE_KO == true && Table2.Rows.Count <= 21)
                {
                    Table2.Rows.Add();
                    Table2.Rows[Table2.RowCount - 2].ReadOnly = false;
   
                        Table2.Rows[Table2.RowCount - 2].Cells[0].Value = Table2.RowCount - 2;

                }
                else
                {
                    if (USE_KO != true && Table2.Rows.Count <= 20)
                    {
                        Table2.Rows.Add();
                        Table2.Rows[Table2.RowCount - 2].ReadOnly = false;
                        Table2.Rows[Table2.RowCount - 2].Cells[0].Value = Table2.RowCount - 1;
                    }
                    else
                    {
                        MessageBox.Show("Количество образцов от 1 до 20!");
                    }
                }
            }
            else
            {
                MessageBox.Show("Измерение еще не создано! Создайте Измерение.");
            }

        }

        private void Remove_Table2_Click(object sender, EventArgs e)
        {
            if (Table2.RowCount != 0)
            {
                
                if (USE_KO == false)
                {
                    if (Table2.Rows.Count > 2)
                    {
                        Table2.Rows.RemoveAt(Table2.CurrentCell.RowIndex);
                        for (int i = 0; i < Table2.RowCount - 1; i++)
                        {
                            Table2.Rows[i].Cells[0].Value = i + 1;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Количество образцов не может быть меньше 1 !");
                    }
                }
                else
                {
                    if (Table2.Rows.Count > 3)
                    {
                        if (Table2.CurrentCell.RowIndex != 0)
                        {
                            Table2.Rows.RemoveAt(Table2.CurrentCell.RowIndex);
                            for (int i = 0; i < Table2.RowCount - 1; i++)
                            {
                                Table2.Rows[i].Cells[0].Value = i;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Удалять Контрольный опыт запрещено!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Количество образцов не может быть меньше 1 !");
                    }
                }
            }
            else
            {
                MessageBox.Show("Таблица не содержит строк!");
            }
        }

        private void Table1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)
            {

                if (Table1.CurrentCell.Value != "" && Table1.CurrentCell.Value != null)
                {
                    CellOpt = Convert.ToDouble(Table1.CurrentCell.Value.ToString());
                }

                ZapicInTable1();

            }
        }


        private void Table1_CancelRowEdit(object sender, QuestionEventArgs e)
        {
            Table1.EndEdit();
            return;
        }

        private void Table2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Table2.CurrentCell.ColumnIndex >= 2 && Table2.CurrentCell.ReadOnly != true)
            {
                if (Table2.CurrentCell.Value != "" && Table2.CurrentCell.Value != null)
                {
                    CellOpt = Convert.ToDouble(Table2.CurrentCell.Value.ToString());
                }
                
                ZapicInTable2();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (selet_rezim == 2) {
                if (tabControl2.SelectedIndex == 0)
                {
                    if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)
                    {

                        if (Table1.CurrentCell.Value != "" && Table1.CurrentCell.Value != null)
                        {
                            CellOpt = Convert.ToDouble(Table1.CurrentCell.Value.ToString());
                        }

                        ZapicInTable1();

                    }
                }
                else
                {
                    if (Table2.CurrentCell.ColumnIndex >= 2 && Table2.CurrentCell.ReadOnly != true)
                    {
                        if (Table2.CurrentCell.Value != "" && Table2.CurrentCell.Value != null)
                        {
                            CellOpt = Convert.ToDouble(Table2.CurrentCell.Value.ToString());
                        }

                        ZapicInTable2();
                    }
                }

            }
            else {
                if (selet_rezim == 1)
                {
                    if (IzmerenieFR_Table.CurrentCell.ColumnIndex == 5)
                    {
                        IzmerenieFR_Table_Zapic();
                    }
                    else
                    {
                        if (IzmerenieFR_Table.CurrentCell.ReadOnly == true)
                        {
                            MessageBox.Show("Изменения запрещены");
                        }
                    }
                }
                else
                {
                    if (selet_rezim == 6)
                    {
                        if (tabControl2.SelectedIndex == 0)
                        {
                            if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)
                            {

                                if (Table1.CurrentCell.Value != "" && Table1.CurrentCell.Value != null)
                                {
                                    CellOpt = Convert.ToDouble(Table1.CurrentCell.Value.ToString());
                                }

                                ZapicInTable1();

                            }
                        }
                        else
                        {
                            StopAgro = true;
                        }
                    }
                    else
                    {
                        if (selet_rezim == 5)
                        {
                            StopSpectr = true;
                        }
                        else
                        {
                            if(selet_rezim == 4)
                            {
                                if (timer2.Enabled == true)
                                {
                                    timer2.Enabled = true;
                                    timer2.Stop();
                                    MinMax();
                                    button14.Enabled = true;
                                    button11.Enabled = false;
                                    
                                }

                            }
                            else
                            {
                                if(selet_rezim == 3)
                                {
                                   StopSpectr = true;
                                   
                                }
                            }
                        }
                    }
                }
                    
            }
        }
        public void IzmerenieFR_Table_Zapic()
        {
            InputBox _InputBox = new InputBox(this);
            _InputBox.ShowDialog();

            IzmerenieFR_Table.CurrentCell.Value = string.Format("{0:0.0}", CellOpt);
            CellOpt = 0;
            IzmerenieFR_Table.Rows[IzmerenieFR_Table.CurrentRow.Index].Cells[6].Value = string.Format("{0:0.0000}", 
                Convert.ToDouble(IzmerenieFR_Table.Rows[IzmerenieFR_Table.CurrentRow.Index].Cells[3].Value)
                * Convert.ToDouble(IzmerenieFR_Table.Rows[IzmerenieFR_Table.CurrentRow.Index].Cells[5].Value));
        }
        private void button13_Click(object sender, EventArgs e)
        {
            if (ComPodkl == true)
            {
                WalveNew();
            }
            else
            {
                MessageBox.Show("Подключитесь к прибору!");
            }
        }
        public void WalveNew()
        {
            NewWalve _NewWalve = new NewWalve(this);
            _NewWalve.ShowDialog();
        }
        int IzmerFr_count = 0;
        private void button15_Click(object sender, EventArgs e)
        {
            if (IzmerenieFR_Table.RowCount <= 35)
            {
                if (IzmerenieFR_Table.RowCount > 1)
                {
                    IzmerFr_count = IzmerenieFR_Table.RowCount - 1;
                    IzmerenieFR_Table.Rows.Add();
                    IzmerenieFR_Table.Rows[IzmerFr_count].Cells["N"].Value = IzmerFr_count + 1;
                    IzmerenieFR_Table.Rows[IzmerFr_count].Cells["Walve"].Value = GWNew.Text;
                    IzmerenieFR_Table.Rows[IzmerFr_count].Cells["KOne"].Value = "0.0";
                }
                else
                {
                    MessageBox.Show("Создайте новое Измерение");
                }
            }
            else
            {
                MessageBox.Show("Строк не более 35");
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (IzmerenieFR_Table.RowCount > 1)
            {
                if (IzmerenieFR_Table.RowCount > 2)
                {
                    if (IzmerenieFR_Table.CurrentCell.RowIndex != IzmerenieFR_Table.RowCount - 1)
                    {
                        IzmerenieFR_Table.Rows.RemoveAt(IzmerenieFR_Table.CurrentCell.RowIndex);
                        for (int i = 0; i < IzmerenieFR_Table.RowCount - 1; i++)
                        {
                            IzmerenieFR_Table.Rows[i].Cells[0].Value = i + 1;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Удаление запрещено!");
                    }
                }

                else
                {
                    MessageBox.Show("Количество образцов не может быть меньше 1 !");
                }


            }
            else
            {
                MessageBox.Show("Таблица не содержит строк!");
            }
        }

        private void tabPage7_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ScanTable.Rows.Clear();
            int number = listBox1.SelectedIndex;
            int m = countScan[number].GetLength(0);
            int n = countScan[number].GetLength(1);
            ScanTable.ColumnCount = n;
            ScanTable.RowCount = m + 1;
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    ScanTable.Rows[i].Cells[j].Value = countScan[number][i, j];
                }
            }
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            double max = 0.0;
            double min = 0.0;
            for (int i = 0; i < ScanTable.Rows.Count; i++)
            {
                if (i == 0)
                {
                    if (Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) > Convert.ToDouble(ScanTable.Rows[i + 1].Cells[1].Value))
                    {
                        max = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        dataGridView1.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                        min = max;
                        double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                        double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        ScanChart.Series[0].Points.AddXY(x1, y1);
                        ScanChart.Series[0].ChartType = SeriesChartType.Point;
                    }
                    else
                    {
                        min = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        dataGridView2.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                        max = min;
                        double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                        double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                        ScanChart.Series[0].Points.AddXY(x1, y1);
                        ScanChart.Series[0].ChartType = SeriesChartType.Point;
                    }

                }
                else {
                    if (i + 1 != ScanTable.Rows.Count)
                    {
                        if (Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) > Convert.ToDouble(ScanTable.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) >= Convert.ToDouble(ScanTable.Rows[i + 1].Cells[1].Value)

                            )
                        {
                            max = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            min = max;
                            dataGridView1.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                            double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                            double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            ScanChart.Series[0].Points.AddXY(x1, y1);
                            ScanChart.Series[0].ChartType = SeriesChartType.Point;
                        }
                        if ((Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) < Convert.ToDouble(ScanTable.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value) <= Convert.ToDouble(ScanTable.Rows[i + 1].Cells[1].Value))

                            )
                        {
                            min = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            dataGridView2.Rows.Add(ScanTable.Rows[i].Cells[0].Value, ScanTable.Rows[i].Cells[1].Value, ScanTable.Rows[i].Cells[2].Value);
                            max = min;
                            double x1 = Convert.ToDouble(ScanTable.Rows[i].Cells[0].Value);
                            double y1 = Convert.ToDouble(ScanTable.Rows[i].Cells[1].Value);
                            ScanChart.Series[0].Points.AddXY(x1, y1);
                            ScanChart.Series[0].ChartType = SeriesChartType.Point;
                        }
                    }
                }
            }

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void Table1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void RR_Click(object sender, EventArgs e)
        {

        }

        private void SKO_Click(object sender, EventArgs e)
        {

        }

        private void Opt_dlin_cuvet_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Table2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void F1Text_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void F2Text_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void AgroText2_TextChanged(object sender, EventArgs e)
        {

        }

        private void AgroText1_TextChanged(object sender, EventArgs e)
        {

        }

        private void AgroText0_TextChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void printPreviewTable1_Load(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void Podskazki_Enter(object sender, EventArgs e)
        {

        }

        private void Podskazka_Click(object sender, EventArgs e)
        {

        }

        private void OptichPlot_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void GAText_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void GEText_TextChanged(object sender, EventArgs e)
        {

        }

        private void printPreviewTable2_Load(object sender, EventArgs e)
        {

        }

        private void GWNew_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void справкаToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void калибровкаДляОдноволновогоАнализаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void темновойТокToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void настройкаПортаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void приборToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void графикРезультатаОдноволновогоИзмеренияToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void волновойАнализToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void анализToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void label30_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void IzmerenieFR_Table_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void label50_Click(object sender, EventArgs e)
        {

        }

        private void label52_Click(object sender, EventArgs e)
        {

        }

        private void label51_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void ScanChart_Click(object sender, EventArgs e)
        {

        }

        private void ScanTable_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void label42_Click(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void groupBox7_Enter(object sender, EventArgs e)
        {

        }

        private void AgroTable1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }

        private void label41_Click(object sender, EventArgs e)
        {

        }

        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void label39_Click(object sender, EventArgs e)
        {

        }

        private void label40_Click(object sender, EventArgs e)
        {

        }

        private void groupBox9_Enter(object sender, EventArgs e)
        {

        }

        private void label38_Click(object sender, EventArgs e)
        {

        }

        private void groupBox10_Enter(object sender, EventArgs e)
        {

        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void groupBox11_Enter(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void groupBox12_Enter(object sender, EventArgs e)
        {

        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }

        private void label44_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        private void label46_Click(object sender, EventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void label47_Click(object sender, EventArgs e)
        {

        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void label48_Click(object sender, EventArgs e)
        {

        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }

        private void label49_Click(object sender, EventArgs e)
        {

        }

        private void tabPage8_Click(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void IzmerenieFRprintPreviewTable1_Load(object sender, EventArgs e)
        {

        }

        private void button17_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView5.ColumnCount - 2; i++)
            {
                if (dataGridView5.Columns["Abs " + i].HeaderText == "Abs " + textBoxCO[i].Text + " нм")
                {
                    dataGridView5.Columns["Abs " + i].HeaderText = "%T " + textBoxCO[i].Text + " нм";
                }
                else
                {
                    dataGridView5.Columns["Abs " + i].HeaderText = "Abs " + textBoxCO[i].Text + " нм";
                }
            }
            for (int j = 0; j < dataGridView5.Rows.Count-1; j++)
            {
                for (int i = 0; i < dataGridView5.ColumnCount - 2; i++)
                {
                    if (dataGridView5.Columns["Abs " + i].HeaderText == "Abs " + textBoxCO[i].Text + " нм")
                    {
                        if (massGEMultiAbs[j][i].ToString() != null)
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = string.Format("{0:0.0000}", massGEMultiAbs[j][i]);
                        }
                        else
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = null;
                        }
                    }
                    else
                    {
                        if (massGEMultiT[j][i].ToString() != null)
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = string.Format("{0:0.00}", massGEMultiT[j][i]);
                        }
                        else
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = null;
                        }
                    }
                }
            }
        }
        public void SaveTpPdf2()
        {
            bool doNotWrite = false;
            for (int j = 0; j < Table2.Rows.Count - 1; j++)
            {

                for (int i = 2; i < Table2.Rows[j].Cells.Count; i++)
                {
                    if (Table2.Rows[j].Cells[i].Value == null)
                    {
                        doNotWrite = true;
                        break;

                    }
                }
            }
            if (doNotWrite == true)
            {
                MessageBox.Show("Не вся поля таблицы были заполнены!");
            }
            else
            {
                if (Table2.Rows.Count >= 1)
                {
                    ExportToPDF();
                }
                else
                {
                    MessageBox.Show("Создайте таблицу измерений!");
                }
            }
        }
        public void ExportToPDF()
        {
            string head = @"Протокол выполнения измерений";

            //Creating iTextSharp Table from the DataTable data
            PdfPTable pdfTable = new PdfPTable(Table2.ColumnCount);
            PdfPTable pdfTable2 = new PdfPTable(Table2.ColumnCount - 2 - NoCaIzm1);
            if (NoCaIzm1 <= 3)
            {
                //Creating iTextSharp Table from the DataTable data
                pdfTable = new PdfPTable(Table2.ColumnCount);
                pdfTable.DefaultCell.Padding = 5;
                pdfTable.WidthPercentage = 100;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                pdfTable.DefaultCell.BorderWidth = 1;
            }
            else
            {
                if (NoCaIzm1 > 3 && NoCaIzm1 <= 5)
                {
                    pdfTable = new PdfPTable(2 + NoCaIzm1 * 2);
                    pdfTable.DefaultCell.Padding = 5;
                    pdfTable.WidthPercentage = 100;
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable.DefaultCell.BorderWidth = 1;
                    pdfTable2 = new PdfPTable(Table2.ColumnCount - 2 - NoCaIzm1 * 2);
                    pdfTable2.DefaultCell.Padding = 5;
                    pdfTable2.WidthPercentage = 20;
                    pdfTable2.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable2.DefaultCell.BorderWidth = 1;
                }
                else
                {
                    pdfTable = new PdfPTable(12);
                    pdfTable.DefaultCell.Padding = 5;
                    pdfTable.WidthPercentage = 100;
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable.DefaultCell.BorderWidth = 1;
                    pdfTable2 = new PdfPTable(Table2.ColumnCount - 12);
                    pdfTable2.DefaultCell.Padding = 5;
                    pdfTable2.WidthPercentage = 100;
                    pdfTable2.HorizontalAlignment = Element.ALIGN_LEFT;
                    pdfTable2.DefaultCell.BorderWidth = 1;
                }
            }
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\georgia.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 18f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font fontBold1 = new iTextSharp.text.Font(baseFont, 10f, iTextSharp.text.Font.BOLD);
            iTextSharp.text.Font font1 = new iTextSharp.text.Font(baseFont, 5f, iTextSharp.text.Font.BOLD);
            // iTextSharp.text.Font fontLeft = new iTextSharp.text.Font(baseFont, 9f, iTextSharp.text.Font.NORMAL);

            //Adding Header row
            if (NoCaIzm1 <= 3)
            {
                PdfPCell cell;
                for (int i = 0; i < Table2.ColumnCount; i++)
                {
                    cell = new PdfPCell(new Phrase(Table2.Columns[i].HeaderText, fontBold1));
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                    cell.BorderWidth = 1;
                    cell.Padding = 1;
                    cell.PaddingBottom = 5;
                    pdfTable.AddCell(cell);
                }
                for (int j = 0; j < Table2.Rows.Count - 1; j++)
                {
                    for (int i = 0; i < Table2.ColumnCount; i++)
                    {
                        pdfTable.AddCell(new Phrase(Convert.ToString(Table2.Rows[j].Cells[i].Value), font));
                    }
                }
            }
            else
            {
                if (NoCaIzm1 > 3 && NoCaIzm1 <= 5)
                {
                    PdfPCell cell1;
                    PdfPCell cell;
                    int kIzmer1 = 0;
                    //int NoCaIzm1_1 = 5;
                    for (int i = 0; i < 2 + NoCaIzm1 * 2; i++)
                    {
                        cell = new PdfPCell(new Phrase(Table2.Columns[kIzmer1].HeaderText, fontBold1));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                        cell.BorderWidth = 1;
                        cell.Padding = 1;
                        cell.PaddingBottom = 5;
                        pdfTable.AddCell(cell);
                        kIzmer1++;
                    }
                    kIzmer1 = 0;
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        for (int i = 0; i < 2 + NoCaIzm1 * 2; i++)
                        {
                            pdfTable.AddCell(new Phrase(Convert.ToString(Table2.Rows[j].Cells[kIzmer1].Value), font));
                            kIzmer1++;
                        }
                        kIzmer1 = 0;
                    }
                    int kIzmer = 2 + NoCaIzm1 * 2;
                    for (int i = 0; i < Table2.ColumnCount - 2 - NoCaIzm1 * 2; i++)
                    {
                        cell1 = new PdfPCell(new Phrase(Table2.Columns[kIzmer].HeaderText, fontBold1));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                        //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                        cell1.BorderWidth = 1;
                        cell1.Padding = 1;
                        cell1.PaddingBottom = 5;
                        pdfTable2.AddCell(cell1);
                        kIzmer++;
                    }
                    kIzmer = 2 + NoCaIzm1 * 2;
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        for (int i = 0; i < Table2.ColumnCount - 2 - NoCaIzm1 * 2; i++)
                        {
                            pdfTable2.AddCell(new Phrase(Convert.ToString(Table2.Rows[j].Cells[kIzmer].Value), font));
                            kIzmer++;
                        }
                        kIzmer = 2 + NoCaIzm1 * 2;
                    }
                }
                else
                {
                    PdfPCell cell1;
                    PdfPCell cell;
                    int kIzmer1 = 0;
                    for (int i = 0; i < 12; i++)
                    {
                        cell = new PdfPCell(new Phrase(Table2.Columns[kIzmer1].HeaderText, fontBold1));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                        //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                        cell.BorderWidth = 1;
                        cell.Padding = 1;
                        cell.PaddingBottom = 5;
                        pdfTable.AddCell(cell);
                        kIzmer1++;
                    }
                    kIzmer1 = 0;
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        for (int i = 0; i < 12; i++)
                        {
                            pdfTable.AddCell(new Phrase(Convert.ToString(Table2.Rows[j].Cells[kIzmer1].Value), font));
                            kIzmer1++;
                        }
                        kIzmer1 = 0;
                    }
                    int kIzmer = 12;
                    for (int i = 0; i < Table2.ColumnCount - 12; i++)
                    {
                        cell1 = new PdfPCell(new Phrase(Table2.Columns[kIzmer].HeaderText, fontBold1));
                        cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell1.VerticalAlignment = Element.ALIGN_MIDDLE;
                        //cell.BackgroundColor = new iTextSharp.text.Color(161, 235, 157);
                        cell1.BorderWidth = 1;
                        cell1.Padding = 1;
                        cell1.PaddingBottom = 5;
                        pdfTable2.AddCell(cell1);
                        kIzmer++;
                    }
                    kIzmer = 12;
                    for (int j = 0; j < Table2.Rows.Count - 1; j++)
                    {
                        for (int i = 0; i < Table2.ColumnCount - 12; i++)
                        {
                            pdfTable2.AddCell(new Phrase(Convert.ToString(Table2.Rows[j].Cells[kIzmer].Value), font));
                            kIzmer++;
                        }
                        kIzmer = 12;
                    }
                }
            }

            iTextSharp.text.Rectangle orient = PageSize.A4;

            float margintop = 20;
            float marginleft = 25;
            float marginright = 25;
            float marginbottom = 5;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                Document doc = new Document(orient, marginleft, marginright, margintop, marginbottom);
                PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));

                doc.Open();
                //iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("Image.jpeg");

                Paragraph welcomeParagraph = new Paragraph("Протокол выполнения измерений\n", fontBold);
                welcomeParagraph.Alignment = iTextSharp.text.Element.ALIGN_CENTER;

                Paragraph FileName2 = new Paragraph("Имя файла: " + filepath2, font);
                Paragraph Description2 = new Paragraph("Описание: " + textBox8.Text, font);


                Paragraph DateTime2 = new Paragraph("Дата: " + dateTimePicker2.Value.ToString("dd.MM.yyyy"), font);
                Paragraph WaveLength2 = new Paragraph("Длина волны: " + wavelength1, font);
                Paragraph Pogresh2 = new Paragraph("Погрешность методики: " + textBox7.Text, font);
                Paragraph Opt_dlin_cuvet2 = new Paragraph("Оптическая длина кюветы: " + Opt_dlin_cuvet.Text, font);
                Paragraph F1 = new Paragraph("F1 = " + F1Text.Text, font);
                Paragraph F2 = new Paragraph("F2 = " + F2Text.Text, font);

                Paragraph Graduirovka2 = new Paragraph("Градуировка: ", font);
                Paragraph FileName1 = new Paragraph("Имя файла: " + filepath, font);
                Paragraph Description1 = new Paragraph("Описание: " + Description, font);
                Paragraph Date1 = new Paragraph("Дата: " + DateTime, font);
                Paragraph Date2 = new Paragraph("Действительна до: " + dateTimePicker1.Value.AddDays(Days).ToString("dd.MM.yyyy"), font);
                Paragraph Pogresh1 = new Paragraph("Погрешность методики: " + textBox3.Text, font);
                Paragraph GradYrav = new Paragraph("Градуировочное уравнение: " + label14.Text, font);
                Paragraph ND2 = new Paragraph("НД: " + ND, font);

                Paragraph DateIzmer2 = new Paragraph("Данные измерений: ", font);

                Paragraph InformationAboutPribor = new Paragraph("Информация о приборе\n", font);
                var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                const string model = @"pribor/model";
                var model_var = Path.Combine(applicationDirectory, model);
                string SerNomer_Text = @"pribor/SerNomer";
                var SerNomer_Text_var = Path.Combine(applicationDirectory, SerNomer_Text);

                string InventarNomer_Text = @"pribor/InventarNomer";
                var InventarNomer_Text_var = Path.Combine(applicationDirectory, InventarNomer_Text);

                string SrokIstech_Text = @"pribor/SrokIstech";
                var SrokIstech_Text_var = Path.Combine(applicationDirectory, SrokIstech_Text);

                string Poveren_Text = @"pribor/Poveren";
                var Poveren_Text_var = Path.Combine(applicationDirectory, Poveren_Text); ;
                StreamReader fs = new StreamReader(model_var);
                Paragraph Model = new Paragraph("Модель\n" + fs.ReadLine(), font);
                fs.Close();

                StreamReader fs1 = new StreamReader(SerNomer_Text_var);
                Paragraph SerNomer = new Paragraph("Серийный номер\n" + fs1.ReadLine(), font);
                fs1.Close();

                StreamReader fs2 = new StreamReader(InventarNomer_Text_var);
                Paragraph InventarNomer = new Paragraph("Инвентарный номер\n" + fs2.ReadLine(), font);
                fs2.Close();

                StreamReader fs3 = new StreamReader(Poveren_Text_var);
                DateTime data = Convert.ToDateTime(fs3.ReadLine());
                // data.Date.ToString("d.mm.yyyy"); 
                //  MessageBox.Show(Convert.ToString(data));   
                data = data.AddYears(1);
                fs3.Close();
                Paragraph Poveren = new Paragraph("Поверка действительна до\n" + data.Date.ToString("dd.MM.yyyy"), font);


                PdfPTable Information = new PdfPTable(6);
                PdfPCell Informationcell = new PdfPCell(Model);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(SerNomer);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(Poveren);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);

                Informationcell = new PdfPCell(InventarNomer);
                Informationcell.BorderWidth = 0;
                Informationcell.Colspan = 3;
                Information.AddCell(Informationcell);
                /*  Paragraph Spectrofotometr2 = new Paragraph("Спектфотометр: ", font);
                  Paragraph Model2 = new Paragraph("Модель: __________________________", font);
                  Paragraph Date3 = new Paragraph("Поверка действительна до: __________________________", font);
                  Paragraph SerNom2 = new Paragraph("Серийный номер: __________________________", font);
                  Paragraph InventarNo2 = new Paragraph("Инветарный номер: __________________________", font);
                  */
                Paragraph Vipolnil = new Paragraph("Измерения выполнил(а): ____________________________________", font);
                Paragraph welcomeParagraph1 = new Paragraph("\n", fontBold);



                PdfPTable table = new PdfPTable(9);
                PdfPCell cell = new PdfPCell(DateTime2);
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                table.AddCell(cell);

                /*  cell = new PdfPCell();
                  cell.BorderWidth = 0;
                  table.AddCell(cell);*/

                cell = new PdfPCell(WaveLength2);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell(Pogresh2);
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                table.AddCell(cell);

                cell = new PdfPCell(Opt_dlin_cuvet2);
                cell.BorderWidth = 0;
                cell.Colspan = 4;
                table.AddCell(cell);

                cell = new PdfPCell(F1);
                cell.BorderWidth = 0;
                cell.Colspan = 2;
                table.AddCell(cell);

                cell = new PdfPCell(F2);
                cell.BorderWidth = 0;
                cell.Colspan = 3;
                table.AddCell(cell);



                PdfPTable table1 = new PdfPTable(6);
                PdfPCell cell1 = new PdfPCell(FileName1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 6;
                table1.AddCell(cell1);

                /*  cell = new PdfPCell();
                  cell.BorderWidth = 0;
                  table.AddCell(cell);*/

                cell1 = new PdfPCell(Description1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 6;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Date1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Date2);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(Pogresh1);
                cell1.BorderWidth = 0;
                cell1.Colspan = 2;
                table1.AddCell(cell1);

                cell1 = new PdfPCell(GradYrav);
                cell1.BorderWidth = 0;
                cell1.Colspan = 6;
                table1.AddCell(cell1);

                PdfPTable table2 = new PdfPTable(1);
                PdfPCell cell2 = new PdfPCell(table1);
                cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                cell2.BorderWidth = 0;
                //cell2.Colspan = 1;
                table2.AddCell(cell2);

                PdfPTable table3 = new PdfPTable(1);
                PdfPCell cell3 = new PdfPCell(table);
                cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                cell3.BorderWidth = 0;
                table3.AddCell(cell3);

                doc.Add(welcomeParagraph);
                doc.Add(welcomeParagraph1);
                doc.Add(FileName2);
                doc.Add(welcomeParagraph1);
                doc.Add(Description2);
                doc.Add(welcomeParagraph1);
                doc.Add(table3);
                doc.Add(welcomeParagraph1);
                doc.Add(InformationAboutPribor);
                doc.Add(welcomeParagraph1);
                doc.Add(Information);
                doc.Add(welcomeParagraph1);
                doc.Add(Graduirovka2);
                doc.Add(welcomeParagraph1);
                doc.Add(table2);
                doc.Add(welcomeParagraph1);
                doc.Add(ND2);
                // doc.Add(pdfTable);
                doc.Add(welcomeParagraph1);
                doc.Add(DateIzmer2);
                doc.Add(welcomeParagraph1);
                if (NoCaIzm1 <= 3)
                {
                    doc.Add(pdfTable);
                }
                else
                {
                    if (NoCaIzm1 > 3 && NoCaIzm1 <= 5)
                    {
                        doc.Add(pdfTable);
                        doc.Add(welcomeParagraph1);
                        doc.Add(pdfTable2);
                    }
                    else
                    {
                        doc.Add(pdfTable);
                        doc.Add(welcomeParagraph1);
                        doc.Add(pdfTable2);
                    }
                }
                doc.Add(welcomeParagraph1);
                doc.Add(Vipolnil);
                // doc.Add(Chart_Image);


                doc.Close();
                // sfd.Visible = true;
            }

        }
    }
}
