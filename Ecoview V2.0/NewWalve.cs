using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SWF = System.Windows.Forms;

namespace Ecoview_V2._0
{
    public partial class NewWalve : Form
    {
        EcoviewStandart1 _Analis;
        public NewWalve(EcoviewStandart1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }
        bool form_close = false;
        private void NewWalve_Load(object sender, EventArgs e)
        {
            if (_Analis.ComPodkl == true)
            {
                Walve.Text = _Analis.GWNew.Text;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            form_close = false;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SW();
            _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);
            form_close = true;
            Close();
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
            _Analis.GWNew.Text = string.Format("{0:0.0}", Convert.ToDouble(Walve.Text));
            _Analis.GWNew.Text = _Analis.GWNew.Text.Replace(",", ".");
            SWF.Application.OpenForms["LogoFrm"].Close();
            // _Analis.GW();
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
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
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

        private void NewWalve_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (form_close == false)
            {
                Application.Exit();
            }
        }
    }
}
