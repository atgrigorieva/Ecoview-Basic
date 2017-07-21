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
    public partial class MultiWave : Form
    {
        EcoviewStandart1 _Analis;
        public MultiWave(EcoviewStandart1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }

        private void MultiWave_Load(object sender, EventArgs e)
        {
            comBoxAdd();
            LabelAdd();
        }
        public int oldValue = 3;
        private EventHandler comboBox1_SelectedIndexChange;

        public void comBoxAdd()
        {
            for(int i = 1; i <= 20; i++)
            {
                comboBox1.Items.Add(i);
            }
            comboBox1.Text = comboBox1.Items[2].ToString();
        }
        public void LabelAdd()
        {
            var height = 80;
            var labelx = 26;
            for (int i = 0; i <= 9; i++)
            {
                var label = new Label();
                label.Name = "WLlabel" + i++.ToString();
                label.Text = "ДВ " + i-- + " =";
                label.Width = 85;
                label.Location = new Point(labelx, height);
                height += label.Height+10;
                groupBox1.Controls.Add(label);
            }
            var height1 = 77;
            var textBoxx = 112;


            for (int i = 0; i <= 9; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "WLtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = string.Format("{0:0.0}", 400.00 + i * 20);
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx, height1);
                height1 += _Analis.textBoxCO[i].Height + 13;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox1.Controls.Add(_Analis.textBoxCO[i]);
                _Analis.textBoxCO[i].Enter += new EventHandler(txt_Enter);
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
            }
            var height2 = 80;
            var labelx1 = 218;
            for (int i = 10; i <= 19; i++)
            {
                var label = new Label();
                label.Name = "WLlabel" + i++.ToString();
                label.Text = "ДВ " + i-- + " =";
                label.Width = 85;
                label.Location = new Point(labelx1, height2);
                height2 += label.Height+10;
                this.Controls.Add(label);
                groupBox1.Controls.Add(label);
            }
            var height3 = 77;
            var textBoxx3 = 304;
            for (int i = 10; i <= 19; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "WLtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = string.Format("{0:0.0}", 400.00 + i * 20);

                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx3, height3);
                height3 += _Analis.textBoxCO[i].Height + 13;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox1.Controls.Add(_Analis.textBoxCO[i]);
                _Analis.textBoxCO[i].Enter += new EventHandler(txt_Enter);
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);               
               
            }
            this.comboBox1.SelectedIndexChanged += new EventHandler(comboBox1_SelectedIndexChanged);
            for (int i = oldValue - 1; i >= 0; i--)
            {
                _Analis.textBoxCO[i].Enabled = true;

            }
        }
        void txt_KeyPress(object sender, KeyPressEventArgs e)
        {

                char number = e.KeyChar;
                if (e.KeyChar == 46 && _Analis.textBoxCO[active].Text.IndexOf(',') == -1)
                {
                    e.KeyChar = ',';

                }
                else
                {

                    if (e.KeyChar == 46 && _Analis.textBoxCO[active].Text.IndexOf(',') != -1)
                    {
                        e.Handled = true;
                        return;
                    }

                }
                if (number == 44 && _Analis.textBoxCO[active].Text.IndexOf(',') != -1)
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
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            _Analis.NoCoIzmer = Convert.ToInt32(comboBox1.SelectedItem.ToString());

            if (this.oldValue > Convert.ToInt32(comboBox1.SelectedItem.ToString()))
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
                for (int i = _Analis.NoCoIzmer - 1; i >= 0; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            oldValue = _Analis.NoCoIzmer;
        }
        int active = 0;
        public void txt_Enter(object sender, EventArgs e)
        {
            active = 0;
            for (int i = 0; i < 19; i++)
            {
                if (_Analis.textBoxCO[i].Focused)
                {
                     active = i;
                    _Analis.textBoxCO[i].Leave += new EventHandler(txt_Leave);
                }
            }
        }
        private void txt_Leave(object sender, EventArgs e)
        {

            if (_Analis.ComPort == true && _Analis.textBoxCO[active].Text != "")
            {
                if (_Analis.versionPribor.Contains("V"))
                {
                    if (Convert.ToDouble(_Analis.textBoxCO[active].Text.Replace(".", ",")) < 315)
                    {
                        _Analis.textBoxCO[active].Text = Convert.ToString(315);
                    }
                    if (Convert.ToDouble(_Analis.textBoxCO[active].Text.Replace(".", ",")) > 1050)
                    {
                        _Analis.textBoxCO[active].Text = Convert.ToString(1050);
                    }
                }
                else
                {
                    if (_Analis.versionPribor.Contains("U") && _Analis.versionPribor.Contains("2"))
                    {
                        if (Convert.ToDouble(_Analis.textBoxCO[active].Text.Replace(".", ",")) < 190)
                        {
                            _Analis.textBoxCO[active].Text = Convert.ToString(190);
                        }
                        if (Convert.ToDouble(_Analis.textBoxCO[active].Text.Replace(".", ",")) > 1050)
                        {
                            _Analis.textBoxCO[active].Text = Convert.ToString(1050);
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(_Analis.textBoxCO[active].Text.Replace(".", ",")) < 200)
                        {
                            _Analis.textBoxCO[active].Text = Convert.ToString(200);
                        }
                        if (Convert.ToDouble(_Analis.textBoxCO[active].Text.Replace(".", ",")) > 1050)
                        {
                            _Analis.textBoxCO[active].Text = Convert.ToString(1050);
                        }
                    }
                }
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _Analis.dataGridView5.Rows.Clear();
            while (true)
            {
                int i = _Analis.dataGridView5.Columns.Count - 1;//С какого столбца начать
                if (_Analis.dataGridView5.Columns[i].Name == "Obrazec1")
                    break;
                _Analis.dataGridView5.Columns.RemoveAt(i);
            }
            for (int i = 0; i < Convert.ToInt32(comboBox1.SelectedItem.ToString()); i++)
            {
                DataGridViewTextBoxColumn firstColumn1 =
                new DataGridViewTextBoxColumn();
                firstColumn1.HeaderText = "Abs " + _Analis.textBoxCO[i].Text + " нм";
                firstColumn1.Name = "Abs " + i;
                firstColumn1.ValueType = Type.GetType("System.Double");
                firstColumn1.ReadOnly = true;
                _Analis.dataGridView5.Columns.Add(firstColumn1);
            }
            _Analis.code = textBox1.Text;
            _Analis.direction = textBox2.Text;
            _Analis.DateTime = dateTimePicker1.Value.AddDays(_Analis.Days).ToString("dd.MM.yyyy");
            _Analis.Ispolnitel = textBox3.Text;
            _Analis.massGEMultiAbs = new double[1][];
            _Analis.massGEMultiT = new double[1][];
            Close();
        }
    }
}
