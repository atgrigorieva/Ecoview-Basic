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
    public partial class ProgrammVersion : Form
    {
        public ProgrammVersion()
        {
            InitializeComponent();
            richTextBox1.Focus();
        }

        private void ProgrammVersion_Load(object sender, EventArgs e)
        {
            
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Net.WebClient wc = new System.Net.WebClient();
                string versionURL = "http://pe-lab.ru/ecoview-version/version-normal";

                if (label1.Text.Substring(6) == wc.DownloadString(versionURL))
                {
                    
                    richTextBox1.Text = "Вы используете актуальную версию программы!";
                }
                else
                {
                    richTextBox1.Text = "Доступна новая версия " + wc.DownloadString(versionURL) + "\nОбратитесь к поставщику прибора!";
                }
                richTextBox1.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Bold);
                richTextBox1.Location = new Point(204, 115);

            }

            catch
            {
                richTextBox1.Text = "Отсутствует подключение к интернету или сервер временно не доступен!\n";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
