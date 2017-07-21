using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using SWF = System.Windows.Forms;
namespace Ecoview_V2._0
{
    public partial class Kinetica : Form
    {
        EcoviewStandart1 _Analis;
        public Kinetica(EcoviewStandart1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
            comboBox1.Text = comboBox1.Items[3].ToString();
            _Analis.timer2.Tick -= _Analis.TableKinetica;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((Convert.ToDouble(textBox4.Text) % Convert.ToDouble(comboBox1.SelectedItem.ToString())) != 0)
            {
                MessageBox.Show("Общее время должно быть кратно интервалу!");
                return;
            }
            else
            {
                _Analis.TableKinetica1.Rows.Clear();
                //_Analis.countButtonClick = 1;
                _Analis.start = Convert.ToDouble(textBox4.Text);
                _Analis.interval = Convert.ToDouble(comboBox1.SelectedItem.ToString());
                _Analis.delay = Convert.ToDouble(textBox3.Text);
                SW();
               // _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);
                _Analis.massWL = new double[0];
                _Analis.massGE = new double[0];
                _Analis.countscan = 0;
                _Analis.dataGridView3.Rows.Clear();
                _Analis.dataGridView4.Rows.Clear();
                _Analis.chart3.Series[0].Points.Clear();
                _Analis.chart3.Series[1].Points.Clear();
                if (radioButton1.Checked == true)
                {
                    _Analis.TableKinetica1.Columns[1].HeaderText = "Abs";
                    _Analis.TableKinetica1.Columns[2].HeaderText = "%T";
                    _Analis.dataGridView3.Columns[1].HeaderText = "Abs";
                    _Analis.dataGridView3.Columns[2].HeaderText = "%T";
                    _Analis.dataGridView4.Columns[1].HeaderText = "Abs";
                    _Analis.dataGridView4.Columns[2].HeaderText = "%T";
                }
                else
                {
                    _Analis.TableKinetica1.Columns[2].HeaderText = "Abs";
                    _Analis.TableKinetica1.Columns[1].HeaderText = "%T";
                    _Analis.dataGridView3.Columns[1].HeaderText = "%T";
                    _Analis.dataGridView3.Columns[2].HeaderText = "Abs";
                    _Analis.dataGridView4.Columns[1].HeaderText = "%T";
                    _Analis.dataGridView4.Columns[2].HeaderText = "Abs";
                }
                if (_Analis.TableKinetica1.Columns[1].HeaderText == "Abs")
                {
                    //Array.Sort(massGE);
                    _Analis.chart3.ChartAreas[0].AxisY.Minimum = 0;
                    _Analis.chart3.ChartAreas[0].AxisY.Maximum = 3;
                    _Analis.chart3.ChartAreas[0].AxisX.Minimum = 0;
                    _Analis.chart3.ChartAreas[0].AxisX.Maximum = _Analis.start;
                }
                else
                {
                    //Array.Sort(massGE);
                    _Analis.chart3.ChartAreas[0].AxisY.Minimum = 0;
                    _Analis.chart3.ChartAreas[0].AxisY.Maximum = 125;
                    _Analis.chart3.ChartAreas[0].AxisX.Minimum = 0;
                    _Analis.chart3.ChartAreas[0].AxisX.Maximum = _Analis.start;
                }
                _Analis.Description = textBox1.Text;
                _Analis.label53.Text = Convert.ToString(_Analis.delay);
                /*  TimerCallback tm = new TimerCallback(_Analis.TableKinetica);
                  System.Threading.Timer timer = new System.Threading.Timer(tm, _Analis.delay,
                      Convert.ToInt32(_Analis.start), Convert.ToInt32(_Analis.interval));*/
                _Analis.timer2.Interval = Convert.ToInt32(_Analis.interval*1000); // 500 миллисекунд
                _Analis.timer2.Enabled = false;

                _Analis.code = textBox7.Text;
                _Analis.direction = textBox6.Text;
                _Analis.DateTime = dateTimePicker1.Value.AddDays(_Analis.Days).ToString("dd.MM.yyyy");
                _Analis.Ispolnitel = textBox5.Text;
                // button1.Click += button1_Click;
                _Analis.timer2.Tick += _Analis.TableKinetica;
               // _Analis.timer2.Start();
                Close();

            }
        }
        public void SW()
        {
            _Analis.LogoForm();
            string SWText1 = textBox2.Text;
            double Walve_double = Convert.ToDouble(textBox2.Text.Replace(".", ","));
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
            _Analis.GWNew.Text = string.Format("{0:0.0}", Convert.ToDouble(textBox2.Text));
            _Analis.GWNew.Text = _Analis.GWNew.Text.Replace(",", ".");
            SWF.Application.OpenForms["LogoFrm"].Close();
            // _Analis.GW();
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

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if(Convert.ToDouble(textBox3.Text) > 3600)
            {
                textBox3.Text = "3600,0";
            }
            if(Convert.ToDouble(textBox3.Text) < 0)
            {
                textBox3.Text = "0,0";
            }
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            if (Convert.ToDouble(textBox3.Text) > 360000)
            {
                textBox3.Text = "360000,0";
            }
            if (Convert.ToDouble(textBox3.Text) < 0)
            {
                textBox3.Text = "0,0";
            }
        }
    }
}
