using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ecoview_V2._0
{
    public partial class FotometrScan : Form
    {
        EcoviewProfessional1 _Analis;
        public FotometrScan(EcoviewProfessional1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
            comboBox1.Text = comboBox1.Items[3].ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(Convert.ToDouble(textBox2.Text) < Convert.ToDouble(textBox3.Text))
            {
                MessageBox.Show("Начальная длина волны не может быть меньше последней!");
                return;
            }
            else
            {
                _Analis.ScanTable.Rows.Clear();
                _Analis.start = Convert.ToDouble(textBox2.Text);
                _Analis.cancel = Convert.ToDouble(textBox3.Text);
                int rows = 0;
                for(double i = _Analis.start; i >= _Analis.cancel; i -= Convert.ToDouble(comboBox1.SelectedItem.ToString()))
                {
                   _Analis.ScanTable.Rows.Add(string.Format("{0:0.0}", i));
                    rows++;
                }
                
            }
            if(radioButton1.Checked == true)
            {
                _Analis.ScanTable.Columns["Abs_scan"].HeaderText = "Abs";
            }
            else
            {
                _Analis.ScanTable.Columns["Abs_scan"].HeaderText = "%T";
            }
            _Analis.Description = textBox1.Text;
            Close();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && textBox3.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && textBox3.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && textBox3.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar >= 58 || e.KeyChar <= 47) && number != 8 && number != 44 && number != 46) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && textBox3.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && textBox3.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && textBox3.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar >= 58 || e.KeyChar <= 47) && number != 8 && number != 44 && number != 46) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
            }
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (_Analis.ComPort == true && textBox3.Text != "")
            {
                if (_Analis.versionPribor.Contains("V"))
                {
                    if (Convert.ToDouble(textBox3.Text.Replace(".", ",")) < 315)
                    {
                        textBox3.Text = Convert.ToString(315);
                    }
                    if (Convert.ToDouble(textBox3.Text.Replace(".", ",")) > 1050)
                    {
                        textBox3.Text = Convert.ToString(1050);
                    }
                }
                else
                {
                    if (_Analis.versionPribor.Contains("U") && _Analis.versionPribor.Contains("2"))
                    {
                        if (Convert.ToDouble(textBox3.Text.Replace(".", ",")) < 190)
                        {
                            textBox3.Text = Convert.ToString(190);
                        }
                        if (Convert.ToDouble(textBox3.Text.Replace(".", ",")) > 1050)
                        {
                            textBox3.Text = Convert.ToString(1050);
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(textBox3.Text.Replace(".", ",")) < 200)
                        {
                            textBox3.Text = Convert.ToString(200);
                        }
                        if (Convert.ToDouble(textBox3.Text.Replace(".", ",")) > 1050)
                        {
                            textBox3.Text = Convert.ToString(1050);
                        }
                    }
                }
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (_Analis.ComPort == true && textBox2.Text != "")
            {
                if (_Analis.versionPribor.Contains("V"))
                {
                    if (Convert.ToDouble(textBox2.Text.Replace(".", ",")) < 315)
                    {
                        textBox2.Text = Convert.ToString(315);
                    }
                    if (Convert.ToDouble(textBox2.Text.Replace(".", ",")) > 1050)
                    {
                        textBox2.Text = Convert.ToString(1050);
                    }
                }
                else
                {
                    if (_Analis.versionPribor.Contains("U") && _Analis.versionPribor.Contains("2"))
                    {
                        if (Convert.ToDouble(textBox2.Text.Replace(".", ",")) < 190)
                        {
                            textBox2.Text = Convert.ToString(190);
                        }
                        if (Convert.ToDouble(textBox2.Text.Replace(".", ",")) > 1050)
                        {
                            textBox2.Text = Convert.ToString(1050);
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(textBox2.Text.Replace(".", ",")) < 200)
                        {
                            textBox2.Text = Convert.ToString(200);
                        }
                        if (Convert.ToDouble(textBox2.Text.Replace(".", ",")) > 1050)
                        {
                            textBox2.Text = Convert.ToString(1050);
                        }
                    }
                }
            }
        }
    }
}
