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
    public partial class Select : Form
    {
        EcoviewProfessional1 _Analis;
        public Select(EcoviewProfessional1 parent)
        {
            InitializeComponent();
            this._Analis = parent;
        }
        bool click = false;
        private void button1_Click(object sender, EventArgs e)
        {
            click = false;
            Application.Exit();
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            click = true;
            if (radioButton1.Checked == true)
            {
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage3);
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage4);
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage2);
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage5);
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage6);
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                _Analis.Ecoview_Header = "Eciview Professional v1.0 Фотометрический режим";
                _Analis.tabControl2.SelectedIndex = 2;
                _Analis.tabControl2.SelectTab(_Analis.tabPage1);
                _Analis.selet_rezim = 1;
            }
            else
            {
                if(radioButton2.Checked == true)
                {
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage1);
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage2);
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage5);
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage6);
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                    _Analis.Ecoview_Header = "Eciview Professional v1.0 Количественный режим";
                    _Analis.tabControl2.SelectedIndex = 0;
                    _Analis.tabControl2.SelectTab(_Analis.tabPage3);
                    _Analis.selet_rezim = 2;
                }
                else
                {
                    if(radioButton3.Checked == true)
                    {
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage3);
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage4);
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage1);
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage5);
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage6);
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                        _Analis.Ecoview_Header = "Eciview Professional v1.0 Многоволновой режим";
                        _Analis.tabControl2.SelectedIndex = 3;
                        _Analis.tabControl2.SelectTab(_Analis.tabPage2);
                        _Analis.selet_rezim = 3;
                    }
                    else
                    {
                        if(radioButton4.Checked == true)
                        {
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage3);
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage4);
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage2);
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage1);
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage6);
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                            _Analis.Ecoview_Header = "Eciview Professional v1.0 Кинетический режим";
                            _Analis.tabControl2.SelectedIndex = 4;
                            _Analis.tabControl2.SelectTab(_Analis.tabPage5);
                            _Analis.selet_rezim = 4;
                        }
                        else
                        {
                            if(radioButton5.Checked == true)
                            {
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage3);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage4);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage2);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage5);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage1);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                                _Analis.Ecoview_Header = "Eciview Professional v1.0 Сканирование спектра";
                                _Analis.tabControl2.SelectedIndex = 5;
                                _Analis.tabControl2.SelectTab(_Analis.tabPage6);
                                _Analis.selet_rezim = 5;
                            }
                            else
                            {
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage2);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage5);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage6);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage1);
                                _Analis.Ecoview_Header = "Eciview Professional v1.0 Агро режим";
                                _Analis.tabControl2.SelectedIndex = 6;
                                _Analis.tabControl2.SelectTab(_Analis.tabPage3);
                                _Analis.selet_rezim = 6;
                                _Analis.tabControl2.TabPages[0].Text = "Градуировка Агро";
                                
                            }
                        }
                    }
                }
            }
            Close();
        }

        private void Select_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(click != true)
            {
                Application.Exit();
            }
        }

        private void Select_Load(object sender, EventArgs e)
        {

        }
    }
}
