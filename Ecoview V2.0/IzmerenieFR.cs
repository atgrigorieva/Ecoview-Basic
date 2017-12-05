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
    public partial class IzmerenieFR : Form
    {
        EcoviewProfessional1 _Analis;
        public IzmerenieFR(EcoviewProfessional1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
            if (_Analis.ComPodkl == true)
            {
                Walve.Text = _Analis.GWNew.Text;
            }

        }
        bool form_close = false;
        private void button2_Click(object sender, EventArgs e)
        {
            // form_close = false;
            SWF.Application.OpenForms["IzmerenieFR"].Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            k1_linear0.Text = k1_linear0.Text.Replace(",", ".");
            _Analis.IzmerenieFR_RowsRemove2();
            string Dlina = Walve.Text;
            for (int i = 0; i < Convert.ToInt32(countIzmer.Text); i++)
            {
                _Analis.IzmerenieFR_Table.Rows.Add();
                _Analis.IzmerenieFR_Table.Rows[i].Cells["N"].Value = i + 1;
                _Analis.IzmerenieFR_Table.Rows[i].Cells["Walve"].Value = Walve.Text;
                _Analis.IzmerenieFR_Table.Rows[i].Cells["KOne"].Value = string.Format("{0:0.0}", k1_linear0.Text);
            }
            _Analis.IzmerenieFR_Table.CurrentCell = _Analis.IzmerenieFR_Table[3, 0];
            SW();
            _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);
            form_close = true;
            _Analis.button11.Enabled = true;
            _Analis.DateTime = dateTimePicker1.Value.Date.ToString("dd.MM.yyyy");
            _Analis.Ispolnitel = administrant.Text;
            _Analis.Description = annotation.Text;
            _Analis.direction = direction_text.Text;
            _Analis.code = code_text.Text;

            _Analis.label26.Visible = false;
            _Analis.label25.Visible = false;
            _Analis.label59.Visible = true;
            _Analis.Podskazka.Text = "Измерьте 0 Asb/100%T";
            //  _Analis.button10.Enabled = true;
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
            _Analis.GWNew.Text = string.Format("{0:0.0}", Walve.Text);
            SWF.Application.OpenForms["LogoFrm"].Close();
            // _Analis.GW();
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && k1_linear0.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && k1_linear0.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && k1_linear0.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && k1_linear0.Text.IndexOf('-') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
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

        private void IzmerenieFR_FormClosing(object sender, FormClosingEventArgs e)
        {
            /* if(form_close == false)
             {
                 Application.Exit();
             }*/
        }

        private void IzmerenieFR_Load(object sender, EventArgs e)
        {

        }
    }
}
