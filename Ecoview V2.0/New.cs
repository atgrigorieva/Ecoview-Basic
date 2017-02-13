using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ecoview_V2._0
{
    public partial class New : Form
    {
        Ecoview _Analis;
        public New(Ecoview parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }

        public void New_Load(object sender, EventArgs e)
        {
            // k0Text1.Text = "k0=";
            DLWave.Text = _Analis.wavelength1;
            int index = Opt_dlin_cuvet.FindString(_Analis.WidthCuvette);
            numericUpDown3.Value = 1;
            numericUpDown4.Value = 1;
            //  MessageBox.Show(index.ToString());
            Opt_dlin_cuvet.SelectedIndex = index;


            Description.Text = _Analis.Description;
            Sozdana.Text = _Analis.DateTime;
            Zavisimost.Text = _Analis.Zavisimoct;
            Aproksimaciya.Text = _Analis.aproksim;
            label11.Text = Convert.ToString(_Analis.CountSeriya);
            label10.Text = Convert.ToString(_Analis.CountInSeriya);
            label9.Text = string.Format("{0:0.0000}", _Analis.textBox4.Text);
            label8.Text = string.Format("{0:0.0000}", _Analis.textBox5.Text);
            label7.Text = string.Format("{0:0.0000}", _Analis.textBox6.Text);
            label12.Text = _Analis.SposobZadan;
            Ed_Izmer.Text = _Analis.edconctr;
            dateTimePicker1.Text = _Analis.DateTime;
            Deistvie.Text = dateTimePicker1.Value.AddDays(_Analis.Days).ToString("dd.MM.yyyy");

            _Analis.Opt_dlin_cuvet.SelectedIndex = index;
            if (_Analis.USE_KO == true)
            {
                USE_KO.Checked = true;
            }
            else
            {
                USE_KO.Checked = false;
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
            "Все текущие параметры и данные измерений будут потеряны. Продолжить?",
            "Подтверждение",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button1,
            MessageBoxOptions.DefaultDesktopOnly);
            if (result == DialogResult.Yes)
            {
                _Analis.NoCaIzm1 = Convert.ToInt32(numericUpDown3.Text);
                _Analis.NoCaSer1 = Convert.ToInt32(numericUpDown4.Text);
                _Analis.textBox8.Text = textBox1.Text;
                _Analis.F1Text.Text = textBox2.Text;
                _Analis.F2Text.Text = textBox3.Text;
                _Analis.textBox7.Text = textBox4.Text;
                _Analis.dateTimePicker2.Text = dateTimePicker1.Text;
                _Analis.WLREMOVE2();
                _Analis.WLREMOVESTR2();
                _Analis.WLADD2();
                _Analis.WLADDSTR2();
                _Analis.IzmerenieOpen = true;
                _Analis.параметрыToolStripMenuItem.Enabled = true;
                _Analis.button10.Enabled = true;
                if (_Analis.ComPodkl == true)
                {
                    _Analis.IzmerCreate1 = true;

                }
                else
                {
                    _Analis.IzmerCreate1 = false;
                }
                if (_Analis.IzmerCreate == true)
                {
                    _Analis.button14.Enabled = true;
                }
                else
                {
                    _Analis.button14.Enabled = false;
                }
                _Analis.Podskazka.Text = "Измеряйте образцы!";
                _Analis.label27.Visible = false;
                _Analis.label24.Visible = false;
                _Analis.label25.Visible = false;
                _Analis.label26.Visible = false;
                _Analis.label28.Visible = true;
            }
            this.TopMost = true;
            Close();
        }

        private void Opt_dlin_cuvet_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = Opt_dlin_cuvet.SelectedIndex;
            _Analis.Opt_dlin_cuvet.SelectedIndex = index;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            _Analis.NoCaIzm1 = Convert.ToInt32(numericUpDown3.Value);

        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            _Analis.NoCaSer1 = Convert.ToInt32(numericUpDown4.Value);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (number == 44 && textBox2.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && textBox2.Text.IndexOf('-') != -1) || (number == 43 && textBox2.Text.IndexOf('+') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 46 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (number == 44 && textBox3.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && textBox3.Text.IndexOf('-') != -1) || (number == 43 && textBox3.Text.IndexOf('+') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 46 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
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
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 46 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void USE_KO_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void USE_KO_Click(object sender, EventArgs e)
        {
            if (sender is CheckBox)
                ((CheckBox)sender).Checked = !((CheckBox)sender).Checked;
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
