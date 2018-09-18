using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SmartHomeServer
{
    public partial class Form1 : Form
    {
        private int curr_x;
        private int curr_y;
        private bool isWndMove;
        private int PageFlag;
        public static ArrayList arduino_port = new ArrayList();
        public static bool arduino_port_flag = false;
        int UpdateCheck = 0;
        SerialPort ComPort = new SerialPort();
        Process process;




        public Form1()
        {
            InitializeComponent();
            AutodetectArduinoPort();
            AddComboBox();
        }

        private void btn_CTA_Click_X(object sender, EventArgs e)
        {
            btn_CTA.BackColor = Color.White;
            btn_CTA.ForeColor = Color.Black;
            btn_CP.BackColor = Color.Black;
            btn_CP.ForeColor = Color.White;
            btn_SC.BackColor = Color.Black;
            btn_SC.ForeColor = Color.White;
            pl_CTA.Visible = true;
            pl_CP.Visible = false;
            pl_SC.Visible = false;
        }

        private void btn_CP_Click_X(object sender, EventArgs e)
        {
            btn_CTA.BackColor = Color.Black;
            btn_CTA.ForeColor = Color.White;
            btn_CP.BackColor = Color.White;
            btn_CP.ForeColor = Color.Black;
            btn_SC.BackColor = Color.Black;
            btn_SC.ForeColor = Color.White;
            pl_CTA.Visible = false;
            pl_CP.Visible = true;
            pl_SC.Visible = false;
        }

        private void btn_SC_Click_X(object sender, EventArgs e)
        {
            btn_CTA.BackColor = Color.Black;
            btn_CTA.ForeColor = Color.White;
            btn_CP.BackColor = Color.Black;
            btn_CP.ForeColor = Color.White;
            btn_SC.BackColor = Color.White;
            btn_SC.ForeColor = Color.Black;
            pl_CTA.Visible = false;
            pl_CP.Visible = false;
            pl_SC.Visible = true;
        }

        private void btn_close_Click_X(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_MouseDown_X(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.curr_x = e.X;
                this.curr_y = e.Y;
                this.isWndMove = true;
            }
        }

        private void panel1_MouseMove_X(object sender, MouseEventArgs e)
        {
            if (this.isWndMove)
                this.Location = new Point(this.Left + e.X - this.curr_x, this.Top + e.Y - this.curr_y);
        }

        private void panel1_MouseUp_X(object sender, MouseEventArgs e)
        {
            this.isWndMove = false;
        }

        public static void AutodetectArduinoPort()
        {
            //Search Arduino Port
            arduino_port.Clear();
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher(connectionScope, serialQuery);
            foreach (ManagementObject item in managementObjectSearcher.Get())
            {
                string desc = item["Description"].ToString();
                string deviceId = item["DeviceID"].ToString();

                if (desc.Contains("Arduino"))
                {
                    //arduino_port[array_count++] = deviceId;
                    arduino_port.Add(deviceId);
                    arduino_port_flag = true;
                }
            }

            //If didn't find Arduino port , list every COM port
            if (arduino_port.Count <= 0)
            {
                foreach (string port in SerialPort.GetPortNames())
                {
                    arduino_port.Add(port);
                    arduino_port_flag = false;
                    //Console.WriteLine("get{0}",port);
                }

            }
        }

        public void AddComboBox()
        {
            if (UpdateCheck != arduino_port.Count)
            {
                if (arduino_port.Count > 0)
                {
                    cb_comport.Items.Clear();
                    for (int i = 0; i < arduino_port.Count; i++)
                    {
                        if (arduino_port_flag == true)
                        {
                            cb_comport.Items.Add("Auto Find Arduino:" + arduino_port[i].ToString());
                        }
                        else
                        {
                            cb_comport.Items.Add(arduino_port[i].ToString());
                        }
                    }
                    cb_comport.SelectedItem = cb_comport.Items[0];
                    UpdateCheck = arduino_port.Count;
                }
                if (cb_comport.SelectedItem != null)
                    btn_comopen.Enabled = true;
                else btn_comopen.Enabled = false;
            }

        }

        private void btn_research_Click_X(object sender, EventArgs e)
        {
            AutodetectArduinoPort();
            AddComboBox();
        }

        private void btn_comopen_Click_X(object sender, EventArgs e)
        {
            if (btn_comopen.Text == "Connect")
            {
                btn_comopen.Text = "Connected";//change text
                //bt_clean.PerformClick();//執行一次'CLEAR'按鈕
                //tb_Srx.AppendText("Detecting Motor........... \r\n");
                btn_comopen.ForeColor = Color.Coral;//change text color

                cb_comport.Enabled = false;//選擇COM的選擇框停用
                if (arduino_port_flag == true)
                {
                    int I = cb_comport.SelectedItem.ToString().IndexOf(":");
                    String SS = cb_comport.SelectedItem.ToString().Substring(I + 1, cb_comport.SelectedItem.ToString().Length - (I + 1));
                    ComPort = new SerialPort(SS);// SerialPort選擇為 選擇COM的選擇框 的COM
                }
                else
                {
                    ComPort = new SerialPort(cb_comport.SelectedItem.ToString());// SerialPort選擇為 選擇COM的選擇框 的COM
                }


                ComPort.BaudRate = Convert.ToInt32(9600); //57600;//設定SerialPort的鲍率為 選擇鲍率的選擇框 的值
                ComPort.Parity = Parity.None;//?
                ComPort.StopBits = StopBits.One;//設定serial port的停止bit為1 bits
                ComPort.DataBits = 8;//設定每一位元組長度為8
                ComPort.Handshake = Handshake.None;//不使用Handshake列舉的值(Handshake為serial port的一個協定)
                ComPort.RtsEnable = true;//在序列通訊期間啟用 Request to Send (RTS) 信號-----在這種模式下，發送端在傳送資料前先送出RTS (Request to Send)要求封包，接收端在收到此一訊息時，會送出CTS (Clear to Send) 封包，告訴發送端可以送出資料並且告訴其他的無線裝置在這段時間內不能傳送任何資料，以避免碰撞。
                ComPort.DtrEnable = true;//在序列通訊期間啟用 Data Terminal Ready(DTR) 信號。-----當Modem已經準備好接收來自PC的資料，它置高DTR線，表示和電話線的連接已經建立。讀取DSR線置高，PC機開始發送資料。一個簡單的規則是DTR/DSR用於表示系統通信就緒，而RTS/CTS用於單個資料包的傳輸。

                ComPort.ReadTimeout = 1000;//設定讀取作業未完成時延遲1秒
                ComPort.WriteTimeout = 1000;//設定寫入作業未完成時延遲1秒


                //ComPort.DataReceived += new SerialDataReceivedEventHandler(ComPortDataReceived);

                //ComPort.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler001); //Arduino

                try//執行可能會出錯的地方
                {
                    if (ComPort.IsOpen == false) ComPort.Open();


                }
                catch (Exception ex)//出錯時執行
                {
                    string msbC = ex.Message + "\n Plz check HW connection again!!";
                    // DialogResult dResult = CatMsg(0, MSGBOX_EXPT, msbC, MessageBoxButtons.OK);

                    btn_comopen.PerformClick();
                    btn_comopen.Enabled = false;

                    process = System.Diagnostics.Process.Start(Application.ExecutablePath); // to start new instance of application
                    if (process.WaitForInputIdle(15000))
                    {
                        //Win32Api.SetWindowPos(process.MainWindowHandle, Win32Api.HWND_TOP, 0, 0, this.Width, this.Height, 0);
                        this.Close(); //to turn off current app
                    }
                }
                Thread.Sleep(1000);
                pl_CP.Enabled = true;
                pl_SC.Enabled = true;
            }
            else if (btn_comopen.Text == "Connected")
            {
                btn_comopen.Text = "Connect";
                btn_comopen.ForeColor = Color.White;
                ComPort.DataReceived += null;
                ComPort.Close();
                cb_comport.Enabled = true;
                pl_CP.Enabled = false;
                pl_SC.Enabled = false;
            }
        }
        public void Port_write(String data)
        {
            if (ComPort.IsOpen == true)
            {
                String DataSendHandler = data;
                if (DataSendHandler.Length > 0)
                {
                    ComPort.Write(DataSendHandler.ToString());
                }
            }
            else btn_comopen.PerformClick();
        }











        private void btn_close_Click(object sender, EventArgs e)
        {
            btn_close_Click_X( sender,  e);
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            panel1_MouseDown_X(sender, e);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            panel1_MouseMove_X( sender,  e);
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            panel1_MouseUp_X( sender,  e);
        }

        private void btn_research_Click(object sender, EventArgs e)
        {
            btn_research_Click_X( sender,  e);
        }

        private void btn_comopen_Click(object sender, EventArgs e)
        {
            btn_comopen_Click_X( sender, e);
        }

        private void btn_CTA_Click(object sender, EventArgs e)
        {
            btn_CTA_Click_X(sender, e);
        }

        private void btn_CP_Click(object sender, EventArgs e)
        {
            btn_CP_Click_X(sender, e);
        }

        private void btn_SC_Click(object sender, EventArgs e)
        {
            btn_SC_Click_X(sender, e);
        }

        private void btn_ASwitch_Click(object sender, EventArgs e)
        {
            Port_write("A0");
        }

        private void btn_AON_Click(object sender, EventArgs e)
        {
            Port_write("A1");
        }

        private void btn_AOFF_Click(object sender, EventArgs e)
        {
            Port_write("A2");
        }

        private void btn_BSwitch_Click(object sender, EventArgs e)
        {
            Port_write("B0");
        }

        private void btn_BON_Click(object sender, EventArgs e)
        {
            Port_write("B1");
        }

        private void btn_BOFF_Click(object sender, EventArgs e)
        {
            Port_write("B2");
        }

        private void btn_CSwitch_Click(object sender, EventArgs e)
        {
            Port_write("C0");
        }

        private void btn_CON_Click(object sender, EventArgs e)
        {
            Port_write("C1");
        }

        private void btn_COFF_Click(object sender, EventArgs e)
        {
            Port_write("C2");
        }

        private void btn_DSwitch_Click(object sender, EventArgs e)
        {
            Port_write("D0");
        }

        private void btn_DON_Click(object sender, EventArgs e)
        {
            Port_write("D1");
        }

        private void btn_DOFF_Click(object sender, EventArgs e)
        {
            Port_write("D2");
        }

        private void btn_ESwitch_Click(object sender, EventArgs e)
        {
            Port_write("E0");
        }

        private void btn_ELow_Click(object sender, EventArgs e)
        {
            Port_write("E1");
        }

        private void btn_FSwitch_Click(object sender, EventArgs e)
        {
            Port_write("F0");
        }

        private void btn_FHigh_Click(object sender, EventArgs e)
        {
            Port_write("F1");
        }

        private void btn_FLow_Click(object sender, EventArgs e)
        {
            Port_write("F2");
        }

        private void btn_FMode_Click(object sender, EventArgs e)
        {
            Port_write("F3");
        }

        private void btn_Sence1_Click(object sender, EventArgs e)
        {
            Port_write("A1B1C1D1");
        }

        private void btn_Sence2_Click(object sender, EventArgs e)
        {
            Port_write("A0B0C0D0");
        }

        private void btn_Sence3_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence4_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence5_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence6_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence7_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence8_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence9_Click(object sender, EventArgs e)
        {

        }

        private void btn_Sence10_Click(object sender, EventArgs e)
        {
            Port_write("A1B0C0D1E1");
        }

        private void btn_Sence11_Click(object sender, EventArgs e)
        {
            Port_write("A1B0C0D0E1");
        }

        private void btn_Sence12_Click(object sender, EventArgs e)
        {
            Port_write("A0B0C0D0E1");
        }
    }
}
