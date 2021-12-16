using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Configuration;
using System.Threading;
using System.Windows.Forms;

namespace TestGUIDriver
{
    public partial class Form1 : Form
    {
        private DateTime lasTime;
        private short port, comMap;
        private int delay;
        private Keys keymap;
        private bool isDebugMode;
        private IKeyDown keyDownHandler;

        public Form1()
        {
            InitializeComponent();
            lasTime = DateTime.Now;
            try
            {
                initSetting();
                hideWindow();
                initSerial();
            }
            catch (FormatException e)
            {
                MessageBox.Show("配置文件格式错误，请重新编辑配置文件。");
                System.Environment.Exit(0);
            }
            catch (NullReferenceException e)
            {
                MessageBox.Show("配置数量不正确，请重新编辑配置文件。");
                System.Environment.Exit(0);
            }
            catch (COMException e)
            {
                MessageBox.Show("COM端口配置错误，请修改配置文件");
                System.Environment.Exit(0);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, e.Source);
                System.Environment.Exit(0);
            }
        }

        private void initSetting()
        {
            string file = Application.ExecutablePath;
            Configuration config = ConfigurationManager.OpenExeConfiguration(file);
            port = Convert.ToInt16(config.AppSettings.Settings["port"].Value);
            delay = Convert.ToInt32(config.AppSettings.Settings["delay"].Value);
            comMap = Convert.ToInt16(config.AppSettings.Settings["com_map"].Value);
            keymap = (Keys)Enum.Parse(typeof(Keys), config.AppSettings.Settings["keymap"].Value);
            isDebugMode = config.AppSettings.Settings["debug_mode"].Value == "true";
            if (config.AppSettings.Settings["enable_win32_api"].Value == "true")
            {
                keyDownHandler = new Win32ApiKeyDown();
            }
            else
            {
                keyDownHandler = new DotNetApiKeyDown();
            }
            keyDownHandler.KeyDown(keymap);
        }

        private void hideWindow()
        {
            Visible = false;
            this.Hide();
        }

        private void initSerial()
        {
            axMSComm1.CommPort = port;
            axMSComm1.PortOpen = true;
            axMSComm1.OnComm += AxMSComm1_OnComm;
        }

        private void AxMSComm1_OnComm(object sender, EventArgs e)
        {
            if (isDebugMode) MessageBox.Show(DateTime.Now+": "+"com key="+axMSComm1.CommEvent);
            if ((DateTime.Now - lasTime).Milliseconds > delay)
            {
                if (axMSComm1.CommEvent == comMap)
                {
                    lasTime = DateTime.Now;
                    keyDownHandler.KeyDown(keymap);
                }
            }
        }
    }

    public interface IKeyDown
    {
        void KeyDown(Keys key);
    }

    public class Win32ApiKeyDown:IKeyDown
    {
        public void KeyDown(Keys key)
        {
            keybd_event((byte)key, 0, 0, 0);
            Thread.Sleep(30);
            keybd_event((byte)key, 0, 2, 0);
        }


        [DllImport("user32.dll", EntryPoint = "keybd_event")]
        public static extern void keybd_event(
            byte bVk,    //虚拟键值
            byte bScan,// 一般为0
            int dwFlags,  //这里是整数类型  0 为按下，2为释放
            int dwExtraInfo  //这里是整数类型 一般情况下设成为 0
        );
    }

    public class DotNetApiKeyDown : IKeyDown
    {
        public void KeyDown(Keys key)
        {
            var keyStr = key.ToString().Length > 1 ? "{" + key.ToString() + "}" : key.ToString();
            SendKeys.Send(keyStr);
        }
    }
}
