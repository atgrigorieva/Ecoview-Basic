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
        EcoviewStandart1 _Analis;
        public Select(EcoviewStandart1 parent)
        {
            InitializeComponent();
            this._Analis = parent;

            // _Analis = this.Owner as EcoviewProfessional1;


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
                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage9);
                _Analis.Ecoview_Header = "Eciview Normal v1.0 Фотометрический режим";
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
                    _Analis.tabControl2.TabPages.Remove(_Analis.tabPage9);
                    _Analis.Ecoview_Header = "Eciview Normal v1.0 Количественный режим";
                    _Analis.tabControl2.SelectedIndex = 0;
                    _Analis.tabControl2.SelectTab(_Analis.tabPage3);
                    _Analis.button13.Enabled = false;
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
                        _Analis.tabControl2.TabPages.Remove(_Analis.tabPage9);
                        _Analis.Ecoview_Header = "Eciview Normal v1.0 Многоволновой режим";
                        _Analis.tabControl2.SelectedIndex = 3;
                        _Analis.button13.Enabled = false;
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
                            _Analis.tabControl2.TabPages.Remove(_Analis.tabPage9);
                            _Analis.Ecoview_Header = "Eciview Normal v1.0 Кинетический режим";
                            _Analis.tabControl2.SelectedIndex = 4;
                            _Analis.button13.Enabled = false;
                            _Analis.tabControl2.SelectTab(_Analis.tabPage5);
                            _Analis.selet_rezim = 4;
                        }
                        else
                        {
                            if(radioButton5.Checked == true)
                            {
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage3);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage4);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage5);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage2);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage1);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage6);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage7);
                                _Analis.tabControl2.TabPages.Remove(_Analis.tabPage8);
                                _Analis.Ecoview_Header = "Eciview Normal v1.0 Работа в Excel";
                                _Analis.tabControl2.SelectedIndex = 9;
                                
                                _Analis.tabControl2.SelectTab(_Analis.tabPage9);
                                _Analis.selet_rezim = 9;
                            }
                        }
                        
                    }
                }
            }
            _Analis.Podskazka.Text = "Подключите прибор!";
            _Analis.label27.Visible = false;            
            _Analis.label24.Visible = true;
            _Analis.label25.Visible = false;
            _Analis.label26.Visible = false;
            _Analis.label28.Visible = false;
            Close();
        }

        private void Select_FormClosing(object sender, FormClosingEventArgs e)
        {
           
            if (click != true)
            {
                ///   Select.Dispose();
                /// Dispose();
               System.Windows.Forms.Application.ExitThread( );  
               // Application.Exit();
                //Application.Current.Shutdown();
///
            }
        }

        private void Select_Load(object sender, EventArgs e)
        {
            
        }

        private void Select_FormClosed(object sender, FormClosedEventArgs e)
        {
            /*if (click != true)
            {
                this.Close();
            }*/
        }


    }
}
