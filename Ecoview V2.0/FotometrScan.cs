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
        EcoviewStandart1 _Analis;
        public FotometrScan(EcoviewStandart1 parent)
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
                _Analis.dataGridView1.Rows.Clear();
                _Analis.dataGridView2.Rows.Clear();
                _Analis.listBox1.Items.Clear();
                _Analis.start = Convert.ToDouble(textBox2.Text);
                _Analis.cancel = Convert.ToDouble(textBox3.Text);
                int rows = 0;
                for(double i = _Analis.start; i >= _Analis.cancel; i -= Convert.ToDouble(comboBox1.SelectedItem.ToString()))
                {
                   _Analis.ScanTable.Rows.Add(string.Format("{0:0.0}", i));
                    rows++;
                }
                
            }
            _Analis.ScanChart.Series[0].Points.Clear();
            _Analis.ScanChart.Series[1].Points.Clear();

            _Analis.ScanChart.Series.Clear();
            _Analis.ScanChart.Series.Add("Series1");
            _Analis.ScanChart.Series.Add("Series2");
            if (radioButton1.Checked == true)
            {
                _Analis.ScanTable.Columns[1].HeaderText = "Abs";
                _Analis.ScanTable.Columns[2].HeaderText = "%T";
                _Analis.dataGridView1.Columns[1].HeaderText = "Abs";
                _Analis.dataGridView1.Columns[2].HeaderText = "%T";
                _Analis.dataGridView2.Columns[1].HeaderText = "Abs";
                _Analis.dataGridView2.Columns[2].HeaderText = "%T";
            }
            else
            {
                _Analis.ScanTable.Columns[2].HeaderText = "Abs";
                _Analis.ScanTable.Columns[1].HeaderText = "%T";
                _Analis.dataGridView1.Columns[1].HeaderText = "%T";
                _Analis.dataGridView1.Columns[2].HeaderText = "Abs";
                _Analis.dataGridView2.Columns[1].HeaderText = "%T";
                _Analis.dataGridView2.Columns[2].HeaderText = "Abs";
            }
            if (_Analis.ScanTable.Columns[1].HeaderText == "Abs")
            {
                //Array.Sort(massGE);
                _Analis.ScanChart.ChartAreas[0].AxisY.Minimum = 0;
                _Analis.ScanChart.ChartAreas[0].AxisY.Maximum = 3;
                _Analis.ScanChart.ChartAreas[0].AxisX.Minimum = _Analis.cancel;
                _Analis.ScanChart.ChartAreas[0].AxisX.Maximum = _Analis.start;
            }
            else
            {
                //Array.Sort(massGE);
                _Analis.ScanChart.ChartAreas[0].AxisY.Minimum = 0;
                _Analis.ScanChart.ChartAreas[0].AxisY.Maximum = 125;
                _Analis.ScanChart.ChartAreas[0].AxisX.Minimum = _Analis.cancel;
                _Analis.ScanChart.ChartAreas[0].AxisX.Maximum = _Analis.start;
            }
            _Analis.Description = textBox1.Text;
            _Analis.button11.Enabled = true;
            _Analis.StopSpectr = false;
            _Analis.scan_mass = null;
            int countStr = _Analis.ScanTable.Rows.Count;
            _Analis.countScan = new string[1][,];
            _Analis.button3.Enabled = true;
           // _Analis.button10.Enabled = true;
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

        private void FotometrScan_Load(object sender, EventArgs e)
        {

        }
    }
}
