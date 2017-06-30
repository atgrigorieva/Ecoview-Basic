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
    public partial class AgroGrad_new : Form
    {
        EcoviewStandart1 _Analis;
        public AgroGrad_new(EcoviewStandart1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }

        private void AgroGrad_new_Load(object sender, EventArgs e)
        {
            if (_Analis.ComPodkl == true)
            {
                WL_grad.Text = _Analis.GWNew.Text;
            }
            var height = 22;
            var labelx = 6;
            for (int i = 0; i <= 9; i++)
            {
                var label = new Label();
                label.Name = "CO" + i++.ToString();
                label.Text = "CO " + i-- + " =";
                label.Width = 40;
                label.Location = new Point(labelx, height);
                height += label.Height;
                groupBox6.Controls.Add(label);
            }
            var height1 = 19;
            var textBoxx = 52;


            for (int i = 0; i <= 9; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "COtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = Convert.ToString("0,00");
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx, height1);
                height1 += _Analis.textBoxCO[i].Height + 3;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox6.Controls.Add(_Analis.textBoxCO[i]);
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
            }
            var height2 = 22;
            var labelx1 = 198;
            for (int i = 10; i <= 19; i++)
            {
                var label = new Label();
                label.Name = "CO" + i++.ToString();
                label.Text = "CO " + i-- + " =";
                label.Width = 40;
                label.Location = new Point(labelx1, height2);
                height2 += label.Height;
                this.Controls.Add(label);
                groupBox6.Controls.Add(label);
            }
            var height3 = 19;
            var textBoxx3 = 244;
            for (int i = 10; i <= 19; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "COtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = Convert.ToString("0,00");
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx3, height3);
                height3 += _Analis.textBoxCO[i].Height + 3;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox6.Controls.Add(_Analis.textBoxCO[i]);
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
            }
            for (int i = Convert.ToInt32(numericUpDown4.Value) - 1; i >= 0; i--)
            {
                this._Analis.textBoxCO[i].Enabled = true;

            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            groupBox6.Enabled = false;
            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            groupBox6.Enabled = true;
        }
        void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }
        int oldValue;
        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            _Analis.NoCoIzmer = Convert.ToInt32(numericUpDown4.Value);
            if (numericUpDown4.Value == 1)
            {
                radioButton2.Enabled = false;
                radioButton3.Enabled = false;
            }
            if (numericUpDown4.Value == 2)
            {
                radioButton2.Enabled = true;
                radioButton3.Enabled = false;
            }
            if (numericUpDown4.Value >= 3)
            {
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
            }



            if (this.oldValue > numericUpDown4.Value)
            {

                for (int i1 = 0; i1 <= 19; i1++)
                {
                    _Analis.textBoxCO[i1].Enabled = false;
                }

                for (int i = _Analis.NoCoIzmer - 1; i >= 0; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            else
            {
                for (int i = _Analis.NoCoIzmer - 1; i >= 1; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            oldValue = _Analis.NoCoIzmer;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && textBox4.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && textBox4.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && textBox4.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && textBox4.Text.IndexOf('-') != -1) || (number == 43 && textBox4.Text.IndexOf('+') != -1))
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

        private void Down_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && Down.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && Down.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && Down.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && Down.Text.IndexOf('-') != -1) || (number == 43 && Down.Text.IndexOf('+') != -1))
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

        private void Up_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && Up.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && Up.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && Up.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && Up.Text.IndexOf('-') != -1) || (number == 43 && Up.Text.IndexOf('+') != -1))
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

        private void k0Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && Agrok0Text.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && Agrok0Text.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && Agrok0Text.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && Agrok0Text.Text.IndexOf('-') != -1) || (number == 43 && Agrok0Text.Text.IndexOf('+') != -1))
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

        private void k1Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && Agrok0Text.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && Agrok0Text.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && Agrok0Text.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && Agrok0Text.Text.IndexOf('-') != -1) || (number == 43 && Agrok0Text.Text.IndexOf('+') != -1))
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

        private void k2Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && Agrok0Text.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && Agrok0Text.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && Agrok0Text.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && Agrok0Text.Text.IndexOf('-') != -1) || (number == 43 && Agrok0Text.Text.IndexOf('+') != -1))
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

        private void WL_grad_Leave(object sender, EventArgs e)
        {
            if (_Analis.ComPort == true && WL_grad.Text != "")
            {
                if (_Analis.versionPribor.Contains("V"))
                {
                    if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) < 315)
                    {
                        WL_grad.Text = Convert.ToString(315);
                    }
                    if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) > 1050)
                    {
                        WL_grad.Text = Convert.ToString(1050);
                    }
                }
                else
                {
                    if (_Analis.versionPribor.Contains("U") && _Analis.versionPribor.Contains("2"))
                    {
                        if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) < 190)
                        {
                            WL_grad.Text = Convert.ToString(190);
                        }
                        if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) > 1050)
                        {
                            WL_grad.Text = Convert.ToString(1050);
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) < 200)
                        {
                            WL_grad.Text = Convert.ToString(200);
                        }
                        if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) > 1050)
                        {
                            WL_grad.Text = Convert.ToString(1050);
                        }
                    }
                }
            }
        }

        private void USE_KO_CheckedChanged(object sender, EventArgs e)
        {
            if (USE_KO.Checked == true)
            {
                _Analis.USE_KO = true;
            }
            else
            {
                _Analis.USE_KO = false;
            }
        }

        private void WL_grad_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && WL_grad.Text.IndexOf(',') != -1)
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

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool Save = false;
            double f = 0.0;
            for (int i = 0; i < Convert.ToInt32(numericUpDown4.Value); ++i)
            {
                if (Convert.ToDouble(_Analis.textBoxCO[i].Text) <= f && radioButton6.Checked == true && radioButton7.Checked != true)
                {
                    MessageBox.Show("Значение CO не может быть МЕНЬШЕ или РАВНО Нулю!");
                    Save = false;
                    break;
                }
                else
                {
                    Save = true;

                }
                if (WL_grad.Text == "")
                {
                    MessageBox.Show("Заполните поле Длина волны");
                    Save = false;
                    break;
                }
                else
                {
                    Save = true;

                }
            }
            if (Save != false)
            {
                DialogResult result = MessageBox.Show(
              "Все текущие параметры и данные градуировки будут потеряны. Продолжить?",
              "Подтверждение",
              MessageBoxButtons.YesNo,
              MessageBoxIcon.Information,
              MessageBoxDefaultButton.Button1,
              MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                {
                    _Analis.chart1.Series[0].Points.Clear();
                    _Analis.chart1.Series[1].Points.Clear();
                    while (true)
                    {
                        int i = _Analis.AgroTable1.Columns.Count - 1;//С какого столбца начать
                        if (_Analis.AgroTable1.Columns.Count == 3 + _Analis.NoCaIzm)
                            break;
                        _Analis.AgroTable1.Columns.RemoveAt(i);
                    }
                    WL();
                    _Analis.textBox1.Text = Description.Text;
                    _Analis.textBox2.Text = Opt_dlin_cuvet.Text;
                    _Analis.Description = Description.Text;
                    _Analis.textBox3.Text = textBox4.Text;
                    _Analis.Veshestvo1 = Veshestvo.Text;
                    _Analis.wavelength1 = WL_grad.Text;
                    _Analis.WidthCuvette = Opt_dlin_cuvet.Text;
                    _Analis.Ispolnitel = Ispolnitel.Text;
                    _Analis.BottomLine = Down.Text;
                    _Analis.TopLine = Up.Text;
                    _Analis.ND = ND.Text;
                    _Analis.DateTime = dateTimePicker1.Text;
                    _Analis.Days = Convert.ToInt32(numericUpDown1.Value);
                    _Analis.CountSeriya = Convert.ToString(numericUpDown3.Value);
                    _Analis.CountInSeriya = Convert.ToString(numericUpDown4.Value);
                    _Analis.textBox10.Text = WL_grad.Text;
                    _Analis.textBox9.Text = WL_grad.Text;
                    _Analis.textBox11.Text = Veshestvo.Text;
                    _Analis.textBox12.Text = Veshestvo.Text;
                    if (radioButton4.Checked == true)
                    {
                        if (radioButton7.Checked == true)
                        {
                            _Analis.groupBox2.Visible = true;
                            if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                            {
                                Agrok0Text.Enabled = false;
                                Agrok1Text.Enabled = true;
                                Agrok2Text.Enabled = false;
                                Agrok0Text.Text = string.Format("{0:0.0000}", 0);
                                Agrok2Text.Text = string.Format("{0:0.0000}", 0);
                                _Analis.AgroText0.Text = string.Format("{0:0.0000}", Agrok0Text.Text);
                                _Analis.AgroText1.Text = string.Format("{0:0.0000}", Agrok1Text.Text);
                                _Analis.AgroText2.Text = string.Format("{0:0.0000}", Agrok2Text.Text);
                                double k0 = Convert.ToDouble(Agrok0Text.Text);
                                double k1 = Convert.ToDouble(Agrok1Text.Text);
                                double k2 = Convert.ToDouble(Agrok2Text.Text);
                                _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C";

                                for (double i = 0; i <= 3; i++)
                                {
                                    double x2 = i;
                                    double y2 = i * k1;
                                    _Analis.chart1.Series[1].Points.AddXY(x2, y2);
                                    //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                    _Analis.chart1.Series[0].Enabled = false;
                                    _Analis.chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + Ed.Text;
                                    _Analis.chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                                    _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                    //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(AgroTable1.Rows[AgroTable1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                    _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                                }
                                // _Analis.chart1
                            }
                            else
                            {
                                if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                                {

                                    Agrok0Text.Enabled = true;
                                    Agrok1Text.Enabled = true;
                                    Agrok2Text.Enabled = false;

                                    Agrok2Text.Text = string.Format("{0:0.0000}", 0);
                                    _Analis.AgroText0.Text = string.Format("{0:0.0000}", Agrok0Text.Text);
                                    _Analis.AgroText1.Text = string.Format("{0:0.0000}", Agrok1Text.Text);
                                    _Analis.AgroText2.Text = string.Format("{0:0.0000}", Agrok2Text.Text);
                                    double k0 = Convert.ToDouble(Agrok0Text.Text);
                                    double k1 = Convert.ToDouble(Agrok1Text.Text);
                                    double k2 = Convert.ToDouble(Agrok2Text.Text);
                                    _Analis.label14.Text = "A(C) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*C" + k0.ToString(" + 0.0000 ;- 0.0000 ");
                                    for (double i = 0; i <= 3; i++)
                                    {
                                        double x2 = i;
                                        double y2 = i * k1 + k0;
                                        _Analis.chart1.Series[1].Points.AddXY(x2, y2);
                                        _Analis.chart1.Series[0].Enabled = false;
                                        //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                        _Analis.chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + Ed.Text;
                                        _Analis.chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                                        _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(AgroTable1.Rows[AgroTable1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                        _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                                    }
                                }
                                else
                                {
                                    Agrok0Text.Enabled = true;
                                    Agrok1Text.Enabled = true;
                                    Agrok2Text.Enabled = true;

                                    _Analis.AgroText0.Text = string.Format("{0:0.0000}", Agrok0Text.Text);
                                    _Analis.AgroText1.Text = string.Format("{0:0.0000}", Agrok1Text.Text);
                                    _Analis.AgroText2.Text = string.Format("{0:0.0000}", Agrok2Text.Text);
                                    double k0 = Convert.ToDouble(Agrok0Text.Text);
                                    double k1 = Convert.ToDouble(Agrok1Text.Text);
                                    double k2 = Convert.ToDouble(Agrok2Text.Text);
                                    _Analis.label14.Text = "A(C) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString(" + 0.0000 ;- 0.0000 ") + "*C" + k2.ToString(" + 0.0000 ;- 0.0000 ") + "*C^2";
                                    for (double i = 0; i <= 3; i++)
                                    {
                                        double x2 = i;
                                        double y2 = i * k1 + k0 + i * k2 * k2;
                                        _Analis.chart1.Series[0].Enabled = false;
                                        _Analis.chart1.Series[1].Points.AddXY(x2, y2);
                                        //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                        _Analis.chart1.ChartAreas[0].AxisX.Title = "Концетрация, " + Ed.Text;
                                        _Analis.chart1.ChartAreas[0].AxisY.Title = "Оптическая плотность, А";
                                        _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(AgroTable1.Rows[AgroTable1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                        _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                                    }
                                }
                            }
                            if (radioButton6.Checked == true)
                            {
                                _Analis.SposobZadan = "По СО";
                            }
                            else
                            {
                                _Analis.SposobZadan = "Ввод коэффициентов";
                                _Analis.button11.Enabled = false;
                            }
                        }
                        else
                        {
                            _Analis.label14.Text = "";
                            if (radioButton4.Checked == true)
                            {
                                _Analis.radioButton4.Checked = true;
                            }
                            else
                            {
                                _Analis.radioButton5.Checked = true;
                            }
                            if (radioButton4.Checked == true)
                            {
                                _Analis.Zavisimoct = "A(C)";
                            }
                            else
                            {
                                _Analis.Zavisimoct = "C(A)";
                            }
                            if (radioButton1.Checked == true)
                            {
                                _Analis.aproksim = "Линейная через 0";
                            }
                            else
                            {
                                if (radioButton2.Checked == true)
                                {
                                    _Analis.aproksim = "Линейная";
                                }
                                else
                                {
                                    _Analis.aproksim = "Квадратичная";
                                }
                            }
                        }
                    }
                    else
                    {
                        if (radioButton7.Checked == true)
                        {
                            _Analis.groupBox2.Visible = true;
                            if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                            {
                                Agrok0Text.Enabled = false;
                                Agrok1Text.Enabled = true;
                                Agrok2Text.Enabled = false;
                                Agrok0Text.Text = string.Format("{0:0.0000}", 0);
                                Agrok2Text.Text = string.Format("{0:0.0000}", 0);
                                _Analis.AgroText0.Text = string.Format("{0:0.0000}", Agrok0Text.Text);
                                _Analis.AgroText1.Text = string.Format("{0:0.0000}", Agrok1Text.Text);
                                _Analis.AgroText2.Text = string.Format("{0:0.0000}", Agrok2Text.Text);
                                double k0 = Convert.ToDouble(Agrok0Text.Text);
                                double k1 = Convert.ToDouble(Agrok1Text.Text);
                                double k2 = Convert.ToDouble(Agrok2Text.Text);
                                _Analis.label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A";
                                for (double i = 0; i <= 3; i++)
                                {
                                    double x2 = i;
                                    double y2 = i * k1;
                                    _Analis.chart1.Series[1].Points.AddXY(y2, x2);
                                    //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                    _Analis.chart1.Series[0].Enabled = false;
                                    _Analis.chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + Ed.Text;
                                    _Analis.chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                                    _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                    //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(AgroTable1.Rows[AgroTable1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                    _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                                }
                            }
                            else
                            {
                                if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                                {

                                    Agrok0Text.Enabled = true;
                                    Agrok1Text.Enabled = true;
                                    Agrok2Text.Enabled = false;

                                    Agrok2Text.Text = string.Format("{0:0.0000}", 0);
                                    _Analis.AgroText0.Text = string.Format("{0:0.0000}", Agrok0Text.Text);
                                    _Analis.AgroText1.Text = string.Format("{0:0.0000}", Agrok1Text.Text);
                                    _Analis.AgroText2.Text = string.Format("{0:0.0000}", Agrok2Text.Text);
                                    double k0 = Convert.ToDouble(Agrok0Text.Text);
                                    double k1 = Convert.ToDouble(Agrok1Text.Text);
                                    double k2 = Convert.ToDouble(Agrok2Text.Text);
                                    _Analis.label14.Text = "C(A) = " + k1.ToString("0.0000 ;- 0.0000 ") + "*A" + k0.ToString(" + 0.0000 ;- 0.0000 ");
                                    for (double i = 0; i <= 3; i++)
                                    {
                                        double x2 = i;
                                        double y2 = i * k1 + k0;
                                        _Analis.chart1.Series[1].Points.AddXY(y2, x2);
                                        _Analis.chart1.Series[0].Enabled = false;
                                        //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                        _Analis.chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + Ed.Text;
                                        _Analis.chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                                        _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(AgroTable1.Rows[AgroTable1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                        _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                                    }
                                }
                                else
                                {
                                    Agrok0Text.Enabled = true;
                                    Agrok1Text.Enabled = true;
                                    Agrok2Text.Enabled = true;

                                    _Analis.AgroText0.Text = string.Format("{0:0.0000}", Agrok0Text.Text);
                                    _Analis.AgroText1.Text = string.Format("{0:0.0000}", Agrok1Text.Text);
                                    _Analis.AgroText2.Text = string.Format("{0:0.0000}", Agrok2Text.Text);
                                    double k0 = Convert.ToDouble(Agrok0Text.Text);
                                    double k1 = Convert.ToDouble(Agrok1Text.Text);
                                    double k2 = Convert.ToDouble(Agrok2Text.Text);
                                    _Analis.label14.Text = "C(A) = " + k0.ToString(" 0.0000 ;- 0.0000 ") + k1.ToString(" + 0.0000 ;- 0.0000 ") + "*A" + k2.ToString(" + 0.0000 ;- 0.0000 ") + "*A^2";
                                    for (double i = 0; i <= 3; i++)
                                    {
                                        double x2 = i;
                                        double y2 = i * k1 + k0 + i * k2 * k2;
                                        _Analis.chart1.Series[0].Enabled = false;
                                        _Analis.chart1.Series[1].Points.AddXY(y2, x2);
                                        //  _Analis.chart1.Series[1].ChartType = SeriesChartType.Line;
                                        _Analis.chart1.ChartAreas[0].AxisY.Title = "Концетрация, " + Ed.Text;
                                        _Analis.chart1.ChartAreas[0].AxisX.Title = "Оптическая плотность, А";
                                        _Analis.chart1.ChartAreas[0].AxisX.Minimum = 0;
                                        //  chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(AgroTable1.Rows[AgroTable1.Rows.Count - 2].Cells["Concetr"].Value) + y2;
                                        _Analis.chart1.ChartAreas[0].AxisY.Minimum = 0;
                                    }
                                }
                            }
                            if (radioButton6.Checked == true)
                            {
                                _Analis.SposobZadan = "По СО";
                                _Analis.button14.Enabled = true;
                            }
                            else
                            {
                                _Analis.SposobZadan = "Ввод коэффициентов";
                                _Analis.button11.Enabled = false;
                                _Analis.button14.Enabled = false;
                            }
                        }
                        else
                        {
                            if (radioButton4.Checked == true)
                            {
                                _Analis.radioButton4.Checked = true;
                            }
                            else
                            {
                                _Analis.radioButton5.Checked = true;
                            }
                            if (radioButton4.Checked == true)
                            {
                                _Analis.Zavisimoct = "A(C)";
                            }
                            else
                            {
                                _Analis.Zavisimoct = "C(A)";
                            }
                            if (radioButton1.Checked == true)
                            {
                                _Analis.aproksim = "Линейная через 0";
                            }
                            else
                            {
                                if (radioButton2.Checked == true)
                                {
                                    _Analis.aproksim = "Линейная";
                                }
                                else
                                {
                                    _Analis.aproksim = "Квадратичная";
                                }
                            }
                        }
                    }
                    _Analis.edconctr = Ed.Text;
                    if (_Analis.ComPodkl == true)
                    {
                        _Analis.SW();
                        _Analis.SAGE(ref _Analis.countSA, ref _Analis.GE5_1_0);
                        _Analis.IzmerCreate = true;
                    }
                    else
                    {
                        _Analis.IzmerCreate = false;
                    }
                    this.TopMost = true;
                    if (_Analis.IzmerCreate == true)
                    {
                        _Analis.button14.Enabled = true;
                    }
                    else
                    {
                        _Analis.button14.Enabled = false;
                    }
                    Close();
                }
            }
        }
        public void WL()
        {
            _Analis.WLREMOVEAgro1();
            _Analis.WLREMOVESTRAgro1();
            _Analis.NoCaIzm = Convert.ToInt32(numericUpDown3.Value);
            _Analis.NoCaSer = Convert.ToInt32(numericUpDown4.Value);
            _Analis.WL_grad1 = WL_grad.Text;
            _Analis.WLADDAgro1();
            _Analis.WLADDSTRAgro1();


        }
    }
}
