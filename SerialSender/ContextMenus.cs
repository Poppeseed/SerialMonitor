using System;
using System.IO.Ports;
using System.Diagnostics;
using System.Windows.Forms;
using SerialSender.Properties;
using System.Drawing;
using OpenHardwareMonitor;
using OpenHardwareMonitor.Hardware;
using System.Timers;
using System.Net.NetworkInformation;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
//using Windows.Media.Control;

namespace SerialSender
{

    class ContextMenus
    {

        SerialPort SelectedSerialPort;
        ContextMenuStrip menu;
        Computer thisComputer;
        DispatcherTimer TxTimer;

        public ContextMenus()
        {
            TxTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
            TxTimer.Tick += (s, ev) => SendHwData();

            string[] ports = SerialPort.GetPortNames();
            string port = Properties.Settings.Default.Port;
            if (ports.Contains(port))
            {
                Selected_Serial(port);
            }
        }

        public ContextMenuStrip Create()
        {
           
            thisComputer = new OpenHardwareMonitor.Hardware.Computer() { };
            thisComputer.CPUEnabled = true;
            thisComputer.GPUEnabled = true;
            thisComputer.HDDEnabled = true;
            thisComputer.MainboardEnabled = true;
            thisComputer.RAMEnabled = true;
            thisComputer.Open();
            

            menu = new ContextMenuStrip();
            CreateMenuItems();
            return menu;
        }

        void CreateMenuItems_Ports()
        {
            string[] ports = SerialPort.GetPortNames();

            foreach (string port in ports)
            {
                var item = new ToolStripMenuItem(port);
                item.Text = port;
                item.Click += new EventHandler((sender, e) => Selected_Serial(port));
                item.Image = Properties.Resources.Serial;
                if (SelectedSerialPort.PortName.Equals(port) && SelectedSerialPort.IsOpen)
                {
                    item.Checked = true;
                }
                menu.Items.Add(item);
            }
        }

        void CreateMenuItems()
        {
            CreateMenuItems_Ports();
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem("Refresh Ports", Resources.Refresh1, new EventHandler((sender, e) => InvalidateMenu())));
            menu.Items.Add(new ToolStripMenuItem("Close Port", Resources.Close2, new EventHandler((sender, e) => CloseSerial())));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem("Exit", Resources.Exit, new EventHandler((sender, e) => Exit_Click())));
        }

        void InvalidateMenu()
        {
            menu.Items.Clear();
            CreateMenuItems();
        }

        void Selected_Serial(string selected_port)
        {
            if(TxTimer.IsEnabled)
            {
                TxTimer.Stop();
            }

            if(SelectedSerialPort != null)
            {
                CloseSerial();
            }

            Console.WriteLine("Selected port");
            Console.WriteLine(selected_port);
            Console.ReadLine();
            SelectedSerialPort = new SerialPort(selected_port);
            if ( !SelectedSerialPort.IsOpen)
            {
                SelectedSerialPort.Open();
            }

            if(!Properties.Settings.Default.Port.Equals(selected_port))
            {
                Properties.Settings.Default.Port = selected_port;
                Properties.Settings.Default.Save();
            }

            if (menu != null && menu.Items.Count > 0)
            {
                foreach (var x in menu.Items)
                {
                    if (x is ToolStripMenuItem)
                    {

                        if (((ToolStripMenuItem)x).Name.Equals(SelectedSerialPort.PortName) || ((ToolStripMenuItem)x).Text.Equals(SelectedSerialPort.PortName))
                        {
                            ((ToolStripMenuItem)x).Checked = true;
                        }

                        if (x is ToolStripMenuItem)
                        {
                            if (((ToolStripMenuItem)x).Name.Equals("Close Port") || ((ToolStripMenuItem)x).Text.Equals("Close Port"))
                            {
                                ((ToolStripMenuItem)x).Enabled = true;
                            }
                        }
                    }
                }
            }

            TxTimer.Start();
        }

        void CloseSerial()
        {
            if (TxTimer.IsEnabled)
            {
                TxTimer.Stop();
            }

            if (SelectedSerialPort.IsOpen)
            {
                SelectedSerialPort.Close();
            }

            foreach (var x in menu.Items)
            {
                if (x is ToolStripMenuItem)
                {
                    //if (((ToolStripMenuItem)x).Name.Equals(SelectedSerialPort.PortName))
                    {
                        ((ToolStripMenuItem)x).Checked = false;
                    }


                    if (x is ToolStripMenuItem)
                    {
                        if (((ToolStripMenuItem)x).Name.Equals("Close Port") || ((ToolStripMenuItem)x).Text.Equals("Close Port"))
                        {
                            ((ToolStripMenuItem)x).Enabled = false;
                        }
                    }
                }
            }

            SelectedSerialPort = null;
        }
        
        class HardwareInfo
        {
            public string cpuTemp = "";
            public string gpuTemp = "";
            public string gpuLoad = "";
            public string cpuLoad = "";
            public string ramUsed = "";
            public string ramLoad = "";
            public string ramAvailable = "";
        }

        private HardwareInfo getHardwareInfo()
        {
            var hwI = new HardwareInfo();

            foreach (IHardware hw in thisComputer.Hardware)
            {
                Console.WriteLine("Checking: " + hw.HardwareType);
                Console.ReadLine();

                hw.Update();
                // searching for all sensors and adding data to listbox
                foreach (OpenHardwareMonitor.Hardware.ISensor s in hw.Sensors)
                {
                    Console.WriteLine("Sensor: " + s.Name + " Type: " + s.SensorType + " Value: " + s.Value);
                    Console.ReadLine();

                    if (s.SensorType == OpenHardwareMonitor.Hardware.SensorType.Temperature)
                    {
                        if (s.Value != null)
                        {
                            int curTemp = (int)s.Value;
                            switch (s.Name)
                            {
                                case "CPU Package":
                                    hwI.cpuTemp = curTemp.ToString();
                                    break;
                                case "GPU Core":
                                    hwI.gpuTemp = curTemp.ToString();
                                    break;

                            }
                        }
                    }
                    if (s.SensorType == OpenHardwareMonitor.Hardware.SensorType.Load)
                    {
                        if (s.Value != null)
                        {
                            int curLoad = (int)s.Value;
                            switch (s.Name)
                            {
                                case "CPU Total":
                                    hwI.cpuLoad = curLoad.ToString();
                                    break;
                                case "GPU Core":
                                    hwI.gpuLoad = curLoad.ToString();
                                    break;
                                case "Memory":
                                    hwI.ramLoad = curLoad.ToString();
                                    break;
                            }
                        }
                    }
                    if (s.SensorType == OpenHardwareMonitor.Hardware.SensorType.Data)
                    {
                        if (s.Value != null)
                        {
                            switch (s.Name)
                            {
                                case "Used Memory":
                                    decimal ramUsedValue = Math.Round((decimal)s.Value, 1);
                                    hwI.ramUsed = ramUsedValue.ToString();
                                    break;
                                case "Available Memory":
                                    decimal ramAvailableValue = Math.Round((decimal)s.Value, 1);
                                    hwI.ramAvailable = ramAvailableValue.ToString();
                                    break;
                            }
                        }
                    }
                }
            }
            return hwI;
        }

        private string getCurrentSong()
        {
            string curSong = "";
            Process[] processlist = Process.GetProcesses();


            //

            //var gsmtcsm = GlobalSystemMediaTransportControlsSessionManager.RequestAsync().GetAwaiter().GetResult().GetCurrentSession();
            //var mediaProperties = gsmtcsm.TryGetMediaPropertiesAsync().GetAwaiter().GetResult();
            //Console.WriteLine("{0} - {1}", mediaProperties.Artist, mediaProperties.Title);

            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {

                    if (process.ProcessName == "AIMP3")
                    {
                        curSong = process.MainWindowTitle;
                    }
                    else if (process.ProcessName == "foobar2000" && (process.MainWindowTitle.IndexOf("[") > 0))
                    {
                        curSong = process.MainWindowTitle.Substring(0, process.MainWindowTitle.IndexOf("[") - 1);
                    }
                }
            }

            return curSong;
        }


        private void SendHwData()
        {
            // enumerating all the hardware
            var hwI = getHardwareInfo();
            string curSong = getCurrentSong();

            // string ramLoad= ramUsed / ramAvailable)
            string arduinoData = "C" + hwI.cpuTemp + "c " + hwI.cpuLoad + "%|G" + hwI.gpuTemp +"c " + hwI.gpuLoad + "%|R"+ hwI.ramUsed +"G " + hwI.ramLoad  + "%|S" + curSong + "|";

            try
            {
                if (SelectedSerialPort.IsOpen)
                {
                    SelectedSerialPort.WriteLine(arduinoData);
                }
            }
            catch (InvalidOperationException e)
            {
                CloseSerial();
            }
            catch (Exception e)
            {

            }
        }


        void Exit_Click()
        {
            Application.Exit();
        }
    }
}