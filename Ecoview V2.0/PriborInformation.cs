using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace Ecoview_V2._0
{
    public partial class PriborInformation : Form
    {
        Ecoview _Analis;
        public PriborInformation(Ecoview parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }
        

        private void PriborInformation_Load(object sender, EventArgs e)
        {
            textBox3.Enabled = false;
            Pribor();
        }
        public void Pribor()
        {
            string model = @"pribor/model";

            string SerNomer_Text = @"pribor/SerNomer";
            string InventarNomer_Text = @"pribor/InventarNomer";
            string SrokIstech_Text = @"pribor/SrokIstech";
            string Poveren_Text = @"pribor/Poveren";
            StreamReader fs = new StreamReader(model);
            string model1;
            model1 = fs.ReadLine();
            int index = Model1.FindString(model1);
            if (index != -1)
            {
                Model1.SelectedIndex = index;

            }
            else
            {
                Model1.SelectedIndex = 0;

            }
            fs.Close();

            StreamReader fs1 = new StreamReader(SerNomer_Text);
            textBox1.Text = fs1.ReadLine();
            fs1.Close();

            StreamReader fs2 = new StreamReader(InventarNomer_Text);
            textBox2.Text = fs2.ReadLine();
            fs2.Close();

            StreamReader fs3 = new StreamReader(SrokIstech_Text);
            textBox3.Text = fs3.ReadLine();
            fs3.Close();

            if (textBox3.Text != "")
            {
                checkBox1.Checked = true;
            }
            else
            {
                textBox3.Enabled = false;
            }

            StreamReader fs4 = new StreamReader(Poveren_Text);
            dateTimePicker1.Text = fs4.ReadLine();
            fs4.Close();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s1 = "";
            string model = @"pribor/model";
            string s = Model1.SelectedItem.ToString();

            File.WriteAllText(model, string.Empty);
            File.AppendAllText(model, s, Encoding.UTF8);

            string SerNomer = textBox1.Text;
            string InventarNomer = textBox2.Text;
            string SrokIstech = textBox3.Text;
            string SerNomer_Text = @"pribor/SerNomer";
            string InventarNomer_Text = @"pribor/InventarNomer";
            string SrokIstech_Text = @"pribor/SrokIstech";
            string Poveren_Text = @"pribor/Poveren";

            File.WriteAllText(SerNomer_Text, string.Empty);
            File.AppendAllText(SerNomer_Text, textBox1.Text, Encoding.UTF8);
            File.WriteAllText(InventarNomer_Text, string.Empty);
            File.AppendAllText(InventarNomer_Text, textBox2.Text, Encoding.UTF8);
            File.WriteAllText(SrokIstech_Text, string.Empty);
            File.AppendAllText(SrokIstech_Text, textBox3.Text, Encoding.UTF8);
            File.WriteAllText(Poveren_Text, string.Empty);
            File.AppendAllText(Poveren_Text, dateTimePicker1.Value.ToString("dd.MM.yyyy"), Encoding.UTF8);
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox3.Enabled = true;
            }
            else
            {
                textBox3.Enabled = false;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
