using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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

using System.Globalization;
using Microsoft.Win32;
using System.Xml.Linq;


namespace Ecoview_V2._0
{
    public partial class ExcelResim : Form
    {
        EcoviewProfessional1 _Analis;
        public ExcelResim(EcoviewProfessional1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }
        bool form_close = false;
        private void button1_Click(object sender, EventArgs e)
        {
            _Analis.openFileDialog1.InitialDirectory = "C";
            _Analis.openFileDialog1.Title = "Open File";
            _Analis.openFileDialog1.FileName = "";
            _Analis.openFileDialog1.Filter = "Excel файл|*.xls; *.xlsx";
            if (_Analis.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _Analis.filepath = _Analis.openFileDialog1.FileName;
                    textBox1.Text = _Analis.filepath;

                    
                }
                catch (Exception t) { MessageBox.Show("exeption" + t.Message); }
            }
        }

        private void ExcelResim_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (_Analis.filepath != null && Walve.Text != "")
            {

                string Dlina = Walve.Text;

               
                SW();
                _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);

                _Analis.button11.Enabled = true;

                //  _Analis.button10.Enabled = true;


                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                _Analis.workBook = excelApp.Workbooks.Open(_Analis.filepath);
                _Analis.workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_Analis.workBook.Worksheets.get_Item(1);
                // Открываем созданный excel-файл
                excelApp.Visible = true;
                excelApp.UserControl = true;

                _Analis.label60.Visible = true;
                _Analis.label60.Text = "Длина волны для измерения: " + Dlina;

                _Analis.label61.Visible = true;
                _Analis.label61.Text = "Файл измерений: " + _Analis.filepath + "\n\nРежим измерений: Abs" +
                    "\n\nДля измерения выберите нужную ячейку " +
                    "в открывшейся таблице Excel и нажмите кнопку ИЗМЕРИТЬ на панели инструментов." +
                    "\n\nВыполняйте процедуру необходимое количество раз.\n\nНе забывайте сохранять таблицу.";
                _Analis.Podskazka.Text = "Откалибруйтесь!";
                _Analis.button11.Enabled = false;
                _Analis.button14.Enabled = true;
                _Analis.button12.Enabled = true;
                form_close = true;
                _Analis.label25.Visible = false;
                _Analis.label26.Visible = false;
                _Analis.label59.Visible = true;
                Close();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл для записи или не задали длину волны");

            }
        }

        private void Walve_Leave(object sender, EventArgs e)
        {
            if (_Analis.ComPort == true && Walve.Text != "")
            {
                if (_Analis.versionPribor.Contains("V"))
                {
                    if (Convert.ToDouble(Walve.Text.Replace(".", ",")) < 315)
                    {
                        Walve.Text = Convert.ToString(315);
                    }
                    if (Convert.ToDouble(Walve.Text.Replace(".", ",")) > 1050)
                    {
                        Walve.Text = Convert.ToString(1050);
                    }
                }
                else
                {
                    if (_Analis.versionPribor.Contains("U") && _Analis.versionPribor.Contains("2"))
                    {
                        if (Convert.ToDouble(Walve.Text.Replace(".", ",")) < 190)
                        {
                            Walve.Text = Convert.ToString(190);
                        }
                        if (Convert.ToDouble(Walve.Text.Replace(".", ",")) > 1050)
                        {
                            Walve.Text = Convert.ToString(1050);
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(Walve.Text.Replace(".", ",")) < 200)
                        {
                            Walve.Text = Convert.ToString(200);
                        }
                        if (Convert.ToDouble(Walve.Text.Replace(".", ",")) > 1050)
                        {
                            Walve.Text = Convert.ToString(1050);
                        }
                    }
                }
            }
        }

        private void ExcelResim_FormClosing(object sender, FormClosingEventArgs e)
        {
         /*   if (form_close == false)
            {
                Application.Exit();
            }*/
        }
        public void SW()
        {
            _Analis.LogoForm();
            string SWText1 = Walve.Text;
            double Walve_double = Convert.ToDouble(Walve.Text.Replace(".", ","));
            _Analis.newPort.Write("SW " + Walve_double.ToString(System.Globalization.CultureInfo.GetCultureInfo("en-US")) + "\r");
            string indata = _Analis.newPort.ReadExisting();

            bool indata_bool = true;
            while (indata_bool == true)
            {
                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = _Analis.newPort.ReadExisting();
                }
            }
            _Analis.GWNew.Text = string.Format("{0:0.0}", Walve.Text);
            _Analis.GWNew.Text = _Analis.GWNew.Text.Replace(",", ".");
            SWF.Application.OpenForms["LogoFrm"].Close();
            // _Analis.GW();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            form_close = false;
            _Analis.label25.Visible = true;
            _Analis.label26.Visible = false;
            _Analis.label59.Visible = false;
            Close();
        }

        private void Walve_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && Walve.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && Walve.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && Walve.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar <= 45 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки ','");
            }
        }

        private void Walve_TextChanged(object sender, EventArgs e)
        {

        }


    }
}
