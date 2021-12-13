using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace TestGUIDriver
{
    public partial class Form1 : Form
    {
        private DateTime lasTime;

        public Form1()
        {
            InitializeComponent();
            Visible = false;
            this.Hide();
            axMSComm1.CommPort = 3;
            axMSComm1.PortOpen = true;
            lasTime = DateTime.Now;
            axMSComm1.OnComm += AxMSComm1_OnComm;
        }

        private void AxMSComm1_OnComm(object sender, EventArgs e)
        {
            Console.WriteLine((DateTime.Now - lasTime).Milliseconds);
            if ((DateTime.Now - lasTime).Milliseconds > 500)
            {
                if (axMSComm1.CommEvent == 4)
                {
                    lasTime = DateTime.Now;
                    keybd_event((byte)Keys.F4, 0, 0, 0);
                }
            }
        }

        [DllImport("user32.dll", EntryPoint = "keybd_event")]
        public static extern void keybd_event(
            byte bVk,    //虚拟键值
            byte bScan,// 一般为0
            int dwFlags,  //这里是整数类型  0 为按下，2为释放
            int dwExtraInfo  //这里是整数类型 一般情况下设成为 0
        );
    }
}
