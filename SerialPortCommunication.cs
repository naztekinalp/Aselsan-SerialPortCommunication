using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Management;
using System.Linq;
namespace SerialPortTutorial
{
    public partial class COMPortForm : Form
    {
        private SerialPort serialPort;

        public COMPortForm()
        {
            InitializeComponent();
            refreshComboBox();
        }

        private void refreshComboBox()
        {
            List<string> portNames = GetAllPorts();
            comboBoxSerialPort.Items.Clear();
            if (portNames != null)
            {
                comboBoxSerialPort.Items.AddRange(portNames.ToArray());
                comboBoxSerialPort.SelectedIndex = 0;
            }

            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("Select * from Win32_SerialPort"))
            {
                string[] portnames = SerialPort.GetPortNames();

                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
                var tList = (from n in portnames
                             join p in ports on n equals p["DeviceID"].ToString()
                             where p["PNPDeviceID"].ToString().Contains("VID_1FD7&PID_3701&MI_02")
                             select n).ToList();

                if (tList != null && tList.Count > 0)
                {
                    this.comboBoxSerialPort.Items.Clear();
                    for (int i = 0; i < tList.Count; i++)
                    {
                        comboBoxSerialPort.Items.Add(tList[i]);
                    }
                    comboBoxSerialPort.SelectedIndex = 0;
                }
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e) // Refresh button
        {
            refreshComboBox();
        }

        private void buttonConnect_Click(object sender, EventArgs e) // Connect button
        {
            if (comboBoxSerialPort.Items.Count == 0)
            {
                MessageBox.Show("Serial Port number could not selected.");
                return;
            }

            try
            {
                if (serialPort != null)
                {
                    if (serialPort.IsOpen)
                        serialPort.Close();
                }
                serialPort = new SerialPort(comboBoxSerialPort.SelectedItem.ToString());
                serialPort.DataReceived += serialPort_DataReceived;
                serialPort.BaudRate = 115200;  // Beginning of serial port properties
                serialPort.ReadBufferSize = 4096;
                serialPort.DataBits = 8;
                serialPort.Parity = Parity.None;
                serialPort.StopBits = StopBits.One;  // End of serial port properties
                serialPort.DtrEnable = true; // To enable COM LOG - COM9
                serialPort.Open();
                buttonConnect.Enabled = false;
                buttonDisconnect.Enabled = true;
                buttonSend.Enabled = true;
                buttonPTT.Enabled = true;
                buttonRelease.Enabled = true;
                boxSend.Enabled = true;
                boxReceive.Enabled = true;

                MessageBox.Show("Port connected.");
            }

            catch (Exception ex)
            {
                MessageBox.Show("Port: " + comboBoxSerialPort.SelectedItem.ToString() + " could not open.\n" + ex.Message);
            }
        }

        void serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e) // For data receive
        {
            while (serialPort.BytesToRead > 0)
            {
                int count = serialPort.BytesToRead;
                byte[] asd = new byte[count];
                serialPort.Read(asd, 0, count);
                string txt = Encoding.UTF8.GetString(asd);
                txt = txt.Replace("\n", "\r\n");
                boxReceive.BeginInvoke(new MethodInvoker(() =>
                    {
                        boxReceive.Text += txt + "\n";
                        boxReceive.SelectionStart = boxReceive.TextLength;
                        boxReceive.ScrollToCaret();
                    }));
            }           
        }

        private void buttonDisconnect_Click(object sender, EventArgs e)  // Disconnect button
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                serialPort.Close();
                buttonConnect.Enabled = true;
                buttonDisconnect.Enabled = false;
                buttonSend.Enabled = false; 
                buttonPTT.Enabled = false;
                buttonRelease.Enabled = false;
                boxSend.Enabled = false;
                boxReceive.Enabled = false;
                MessageBox.Show("Port disconnected.");
            }
        }

        private void buttonSend_Click(object sender, EventArgs e) // Send button
        {
            if (serialPort == null)
                return;

            if (!serialPort.IsOpen)
                return;

            serialPort.Write(boxSend.Text);
            boxSend.Clear();
        }

        private void COMPortForm_FormClosing(object sender, FormClosingEventArgs e) // Form close
        {
            if (serialPort == null)
                return;

            if (serialPort.IsOpen)
                serialPort.Close();
        }

        private void comboBoxSerialPort_MouseEnter(object sender, EventArgs e)
        {
            refreshComboBox();
        }

        private void comboBoxSerialPort_SelectedIndexChanged(object sender, EventArgs e) // Form open 
        {
            buttonPTT.Enabled = false;
            buttonRelease.Enabled = false;
        }

        private void buttonPTT_Click(object sender, EventArgs e) // PTT Button
        {
            serialPort.Write("ccc");
            serialPort.Write("bb tx ana 380750\n");
            serialPort.Write("\x04");
            buttonPTT.Enabled = false;
            buttonRelease.Enabled = true;
        }

        private void buttonRelease_Click(object sender, EventArgs e) // Release button
        {
            serialPort.Write("ccc");
            serialPort.Write("bb tx end\n");
            serialPort.Write("\x04");
            buttonPTT.Enabled = true;
            buttonRelease.Enabled = false;
        }
    }
}
