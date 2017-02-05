using System;
using System.Windows.Forms;
using System.IO.Ports;
using System.Linq;
using System.IO;
using System.Text;

namespace Ecoview_V2._0
{
    public partial class SettingPort : Form
    {
        Ecoview _Analis;
        public SettingPort(Ecoview parent)
        {
            InitializeComponent();
            this._Analis = parent;



            //CO();
            // SW();
            // InitializeTimer();
            string[] ports = SerialPort.GetPortNames();


            for (int i = 0; i < ports.Length; i++)
            {
                SerialPort newPort = new SerialPort();

                // настройки порта (Communication interface)
                newPort.PortName = ports[i];
                newPort.BaudRate = 19200;
                newPort.DataBits = 8;
                newPort.Parity = System.IO.Ports.Parity.None;
                newPort.StopBits = System.IO.Ports.StopBits.One;
                // Установка таймаутов чтения/записи (read/write timeouts)
                newPort.ReadTimeout = 100;
                newPort.WriteTimeout = 100;
                //    newPort.DataReceived += new SerialDataReceivedEventHandler(newPort_DataReceived);
                newPort.RtsEnable = false;
                newPort.DtrEnable = true;
                newPort.Open();// MessageBox.Show("ПОРТ ОТКРЫТ " + newPort.PortName);
                newPort.Write("^*^\r");
                int byteRecieved = newPort.ReadBufferSize;
                System.Threading.Thread.Sleep(50);
                byte[] buffer = new byte[byteRecieved];
                try
                {
                    newPort.Read(buffer, 0, byteRecieved);
                    newPort.DiscardInBuffer();
                    newPort.DiscardOutBuffer();
                    newPort.Close();

                } // Читаем ответ(если ничего не пришло отваливаемся по ReadTimeout = 500
                catch (TimeoutException)
                { /* Девайса нет */

                    newPort.DiscardInBuffer();
                    newPort.DiscardOutBuffer();
                    newPort.Close();
                    ports[i] = null;
                    ports = ports.Where(x => x != null).ToArray();
                    i--;

                }

            }
            string s1 = "";
            StreamReader fs = new StreamReader(@"openport.port");
            string s = "";


            s = fs.ReadLine();
            s1 = s;
            fs.Close();

            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(ports);
            if (ports.Length != 0 && s1 == Convert.ToString(0))
            {
                comboBox1.SelectedIndex = 0;
                _Analis.nonPort = true;
            }
            else
            {
                if (ports.Length != 0 && s != Convert.ToString(0))
                {
                    int index = comboBox1.FindString(s1);
                    if (index != -1)
                    {
                        comboBox1.SelectedIndex = index;
                        _Analis.nonPort = true;
                    }
                    else
                    {
                        comboBox1.SelectedIndex = 0;
                        _Analis.nonPort = true;
                    }
                }
                else
                {
                    MessageBox.Show("Подсоедините спектрофотометр и попробуйте подключиться снова!");
                    _Analis.nonPort = false;
                    Close();
                    // Dispose();
                }
            }
        }

        private void SettingPort_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            _Analis.portsName = comboBox1.SelectedItem.ToString();

            Close();
        }

        private void SettingPort_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_Analis.nonPort == false)
            {
                _Analis.nonPort = false;
                MessageBox.Show("Порт не выбран!");
                Close();
            }
            else
            {
                _Analis.nonPort = true;
                Close();
            }
        }
    }
}
