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
        EcoviewStandart1 _Analis;
        public PriborInformation(EcoviewStandart1 parent)
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
            var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            const string model = @"pribor/model";
            var model_var = Path.Combine(applicationDirectory, model);

            string SerNomer_Text = @"pribor/SerNomer";
            var SerNomer_Text_var = Path.Combine(applicationDirectory, SerNomer_Text);

            string InventarNomer_Text = @"pribor/InventarNomer";
            var InventarNomer_Text_var = Path.Combine(applicationDirectory, InventarNomer_Text);

            string SrokIstech_Text = @"pribor/SrokIstech";
            var SrokIstech_Text_var = Path.Combine(applicationDirectory, SrokIstech_Text);

            string Poveren_Text = @"pribor/Poveren";
            var Poveren_Text_var = Path.Combine(applicationDirectory, Poveren_Text);

            string address_lab_Text = @"pribor/address_lab";
            var address_lab_var = Path.Combine(applicationDirectory, address_lab_Text);

            string name_lab_Text = @"pribor/name_lab";
            var name_lab_var = Path.Combine(applicationDirectory, name_lab_Text);

            StreamReader fs = new StreamReader(model_var);
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

            StreamReader fs1 = new StreamReader(SerNomer_Text_var);
            textBox1.Text = fs1.ReadLine();
            fs1.Close();

            StreamReader fs2 = new StreamReader(InventarNomer_Text_var);
            textBox2.Text = fs2.ReadLine();
            fs2.Close();

            StreamReader fs3 = new StreamReader(SrokIstech_Text_var);
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

            StreamReader fs4 = new StreamReader(Poveren_Text_var);
            dateTimePicker1.Text = fs4.ReadLine();
            fs4.Close();

            StreamReader fs5 = new StreamReader(address_lab_var);
            textBox5.Text = fs5.ReadLine();
            fs5.Close();

            StreamReader fs6 = new StreamReader(name_lab_var);
            textBox4.Text = fs6.ReadLine();
            fs6.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s1 = "";
            var applicationDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            const string model = @"pribor/model";
            var model_var = Path.Combine(applicationDirectory, model);

            string s = Model1.SelectedItem.ToString();

            File.WriteAllText(model, string.Empty);
            File.AppendAllText(model, s, Encoding.UTF8);

            string SerNomer = textBox1.Text;
            string InventarNomer = textBox2.Text;
            string SrokIstech = textBox3.Text;


            string SerNomer_Text = @"pribor/SerNomer";
            var SerNomer_Text_var = Path.Combine(applicationDirectory, SerNomer_Text);

            string InventarNomer_Text = @"pribor/InventarNomer";
            var InventarNomer_Text_var = Path.Combine(applicationDirectory, InventarNomer_Text);

            string SrokIstech_Text = @"pribor/SrokIstech";
            var SrokIstech_Text_var = Path.Combine(applicationDirectory, SrokIstech_Text);

            string Poveren_Text = @"pribor/Poveren";
            var Poveren_Text_var = Path.Combine(applicationDirectory, Poveren_Text);


            string address_lab_Text = @"pribor/address_lab";
            var address_lab_var = Path.Combine(applicationDirectory, address_lab_Text);

            string name_lab_Text = @"pribor/name_lab";
            var name_lab_var = Path.Combine(applicationDirectory, name_lab_Text);

            File.WriteAllText(SerNomer_Text_var, string.Empty);
            File.AppendAllText(SerNomer_Text_var, textBox1.Text, Encoding.UTF8);
            File.WriteAllText(InventarNomer_Text_var, string.Empty);
            File.AppendAllText(InventarNomer_Text_var, textBox2.Text, Encoding.UTF8);
            File.WriteAllText(SrokIstech_Text_var, string.Empty);
            File.AppendAllText(SrokIstech_Text_var, textBox3.Text, Encoding.UTF8);
            File.WriteAllText(Poveren_Text_var, string.Empty);
            File.AppendAllText(Poveren_Text_var, dateTimePicker1.Value.ToString("dd.MM.yyyy"), Encoding.UTF8);
            File.WriteAllText(address_lab_var, string.Empty);
            File.AppendAllText(address_lab_var, textBox5.Text, Encoding.UTF8);
            File.WriteAllText(name_lab_var, string.Empty);
            File.AppendAllText(name_lab_var, textBox4.Text, Encoding.UTF8);

            _Analis.address_lab = textBox5.Text;
            _Analis.name_lab = textBox4.Text;
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
