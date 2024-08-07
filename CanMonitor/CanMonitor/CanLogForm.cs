using libCanopenSimple;
using Microsoft.CSharp;
using N_SettingsMgr;
using PDOInterface;
using PFMMeasurementService.Models.Devices.Buses;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Peak.Can.Basic;
using TPCANHandle = System.Byte;
using can_hw;

namespace CanMonitor
{
    public partial class CanLogForm : DockContent, ICanDocument
    {
        #region PEAK-specific members
        private TPCANHandle[] m_HandlesArray;
        private TPCANHandle m_PcanHandle;
        private TPCANBaudrate m_PcanBaudRate;
        #endregion


        public DockPanel dockpanel;

        List<ListViewItem> listitems = new List<ListViewItem>();






        Dictionary<UInt32, List<byte>> sdotransferdata = new Dictionary<uint, List<byte>>();

        System.Windows.Forms.Timer updatetimer = new System.Windows.Forms.Timer();

        StreamWriter sw;


        Timer savetimer;

        public CanLogForm()
        {
            try
            {
                SettingsMgr.readXML(Path.Combine(Program.appdatafolder, "settings.xml"));
            }
            catch (Exception e)
            {
            }

            InitializeComponent();



            this.comboBox_vendor.Items.Clear();
            var vendorNames = Enum.GetNames(typeof(SupportedVendor));
            for (int i = 0; i < vendorNames.Length; i++)
            {
                this.comboBox_vendor.Items.Add(vendorNames[i]);
            }
            this.comboBox_vendor.SelectedIndex = 1;

            this.m_HandlesArray = new TPCANHandle[] {
                PCANBasic.PCAN_USBBUS1,
                PCANBasic.PCAN_USBBUS2,
                PCANBasic.PCAN_USBBUS3,
                PCANBasic.PCAN_USBBUS4,
                PCANBasic.PCAN_USBBUS5,
                PCANBasic.PCAN_USBBUS6,
                PCANBasic.PCAN_USBBUS7,
                PCANBasic.PCAN_USBBUS8,
            };

            textBox_info.AppendText("Searching for drivers...\r\n\r\n");
            string[] founddrivers = Directory.GetFiles("drivers\\", "*.dll");

            foreach (string driver in founddrivers)
            {
                textBox_info.AppendText(string.Format("Found driver {0}\r\n", driver));
                drivers.Add(driver.Substring(0, driver.Length - 4));
            }

            //enumerateports();
            enumerateports((SupportedVendor)this.comboBox_vendor.SelectedIndex);

            if (comboBox_port.Items.Count == 0)
            {
                MessageBox.Show("No COM ports detected, if the CanUSB is connected please ensure\n that it is set to Load VCP in properties page in device manager\nOr please insert a device now and press 'R' to refresh port list");
            }

            lco.dbglevel = debuglevel.DEBUG_NONE;

            lco.nmtecevent += log_NMTEC;
            lco.nmtevent += log_NMT;
            lco.sdoevent += log_SDO;
            lco.pdoevent += log_PDO;
            lco.emcyevent += log_EMCY;

            listView1.DoubleBuffering(true);



            listView1.ListViewItemSorter = null;

            updatetimer.Interval = 1000;
            updatetimer.Tick += updatetimer_Tick;

            updatetimer_Tick(null, new EventArgs());

            var autoloadPath = Path.Combine(assemblyfolder, "autoload.txt");
            if (File.Exists(autoloadPath))
            {
                string[] autoload = System.IO.File.ReadAllLines(autoloadPath);

                foreach (string plugin in autoload)
                {
                    loadplugin(plugin, false);
                }
            }

            if (appdatafolder != assemblyfolder)
            {
                autoloadPath = Path.Combine(appdatafolder, "autoload.txt");
                if (File.Exists(autoloadPath))
                {
                    string[] autoload = System.IO.File.ReadAllLines(autoloadPath);

                    foreach (string plugin in autoload)
                    {
                        loadplugin(plugin, false);
                    }
                }
            }


            Properties.Settings.Default.Reload();

#if false
            //select last used port
            foreach (driverport dp in comboBox_port.Items)
            {

                if (dp.port == Properties.Settings.Default.lastport)
                {
                    comboBox_port.SelectedItem = dp;
                    break;
                }
            }
#endif

            this.Shown += Form1_Shown;

            updatetimer.Enabled = true;

            savetimer = new Timer();
            savetimer.Interval = 10;
            savetimer.Tick += Savetimer_Tick;
            savetimer.Enabled = true;

            last = DateTime.Now;
        }


        DateTime last;

        public void clearlist()
        {
            lock (listitems)
            {
                listitems.Clear();
                listView1.Items.Clear();
            }
        }

        private void Savetimer_Tick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;

            if (Properties.Settings.Default.AutoFileLog == true && now.Hour != last.Hour)
            {
                //Force a save

                string dtstring = now.ToString("dd-MMM-yyyy-HH-mm-ss");

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                desktopPath += Path.DirectorySeparatorChar + Properties.Settings.Default.FileLogFolder;
                dtstring = desktopPath + Path.DirectorySeparatorChar + dtstring;

                if (!Directory.Exists(desktopPath))
                {
                    Directory.CreateDirectory(desktopPath);
                }


                dtstring += ".xml";
                dosave(dtstring);

                listView1.Items.Clear();

            }

            last = DateTime.Now;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {

            this.Text = "CanOpen Log";

            //FIXME
            //if (Properties.Settings.Default.autoconnect == true)
            // {
            //     button_open_Click(this, new EventArgs());
            // }


        }



        void appendfile(string[] ss)
        {

            if (sw == null)
                return;

            StringBuilder sb = new StringBuilder();
            foreach (string s in ss)
            {
                sb.AppendFormat("{0}\t", s);
            }

            sw.WriteLine(sb.ToString());

        }

        void updatetimer_Tick(object sender, EventArgs e)
        {

            bool limit = Properties.Settings.Default.limitlines;
            int linelimit = Properties.Settings.Default.linelimit;

            if (listitems.Count != 0)
                lock (listitems)
                {
                    listView1.BeginUpdate();

                    listView1.Items.AddRange(listitems.ToArray());

                    listitems.Clear();

                    //    Properties.Settings.Default.Reload();
                    bool s = Properties.Settings.Default.autoscroll;

                    if (Properties.Settings.Default.autoscroll && listView1.Items.Count > 2)
                        listView1.EnsureVisible(listView1.Items.Count - 1);

                    if (limit)
                    {
                        while (listView1.Items.Count > linelimit)
                            listView1.Items.RemoveAt(0);

                    }

                    listView1.EndUpdate();

                }




        }

        private void log_NMT(canpacket payload, DateTime dt)
        {



            string[] items = new string[6];
            items[0] = dt.ToString("MM/dd/yyyy HH:mm:ss.fff");
            items[1] = "NMT";
            items[2] = string.Format("{0:x3}", payload.cob);
            items[3] = "";
            items[4] = BitConverter.ToString(payload.data).Replace("-", string.Empty);

            string msg = "";

            if (payload.data.Length != 2)
                return;

            switch (payload.data[0])
            {
                case 0x01:
                    msg = "Enter operational";
                    break;
                case 0x02:
                    msg = "Enter stop";
                    break;
                case 0x80:
                    msg = "Enter pre-operational";
                    break;
                case 0x81:
                    msg = "Reset node";
                    break;
                case 0x82:
                    msg = "Reset communications";
                    break;

            }

            if (payload.data[1] == 0)
            {
                msg += " - All nodes";
            }
            else
            {
                msg += string.Format(" - Node 0x{0:x2}", payload.data[1]);
            }

            items[5] = msg;

            if (Properties.Settings.Default.showNMTEC)
            {

                ListViewItem i = new ListViewItem(items);

                i.ForeColor = Color.Red;

                lock (listitems)
                    listitems.Add(i);
            }

            appendfile(items);

        }

        private void log_NMTEC(canpacket payload, DateTime dt)
        {



            string[] items = new string[6];
            items[0] = dt.ToString("MM/dd/yyyy HH:mm:ss.fff");
            items[1] = "NMTEC";
            items[2] = string.Format("{0:x3}", payload.cob);
            items[3] = string.Format("{0:x3}", payload.cob & 0x0FF);
            items[4] = BitConverter.ToString(payload.data).Replace("-", string.Empty);

            string msg = "";
            switch (payload.data[0])
            {
                case 0:
                    msg = "BOOT";
                    break;
                case 4:
                    msg = "STOPPED";
                    break;
                case 5:
                    msg = "Heart Beat";
                    break;
                case 0x7f:
                    msg = "Heart Beat (Pre op)";
                    break;

            }

            items[5] = msg;

            ListViewItem i = new ListViewItem(items);

            i.ForeColor = Color.DarkGreen;

            appendfile(items);

            if (Properties.Settings.Default.showNMTEC && (Properties.Settings.Default.showHB == true || payload.data[0] == 0))
            {
                lock (listitems)
                {
                    listitems.Add(i);
                }
            }





        }

        private void log_SDO(canpacket payload, DateTime dt)
        {


            string[] items = new string[6];
            items[0] = dt.ToString("MM/dd/yyyy HH:mm:ss.fff");
            items[1] = "SDO";
            items[2] = string.Format("{0:x3}", payload.cob);

            if (payload.cob >= 0x580 && payload.cob < 0x600)
            {
                items[3] = string.Format("{0:x3}", ((payload.cob + 0x80) & 0x0FF));
            }
            else
            {
                items[3] = string.Format("{0:x3}", payload.cob & 0x0FF);
            }

            items[4] = BitConverter.ToString(payload.data).Replace("-", string.Empty);

            string msg = "";


            int SCS = payload.data[0] >> 5; //7-5

            int n = (0x03 & (payload.data[0] >> 2)); //3-2 data size for normal packets
            int e = (0x01 & (payload.data[0] >> 1)); // expidited flag
            int s = (payload.data[0] & 0x01); // data size set flag //also in block
            int c = s;

            int sn = (0x07 & (payload.data[0] >> 1)); //3-1 data size for segment packets
            int t = (0x01 & (payload.data[0] >> 4));  //toggle flag

            int cc = (0x01 & (payload.data[0] >> 2));



            UInt16 index = (UInt16)(payload.data[1] + (payload.data[2] << 8));
            byte sub = payload.data[3];


            int valid = 7;
            int validsn = 7;


            if (n != 0)
                valid = 8 - (7 - n);

            if (sn != 0)
                validsn = 8 - (7 - sn);


            if (payload.cob >= 0x580 && payload.cob <= 0x600)
            {
                string mode = "";
                string sdoproto = "";

                string setsize = "";

                switch (SCS)
                {
                    case 0:
                        mode = "upload segment response";
                        sdoproto = string.Format("{0} {1} Valid bytes = {2} {3}", mode, t == 1 ? "TOG ON" : "TOG OFF", validsn, c == 0 ? "MORE" : "END");

                        if (c == 1)
                        {
                            //ipdo.endsdo(payload.cob, index, sub, null);
                            //END
                        }

                        if (sdotransferdata.ContainsKey(payload.cob))
                        {

                            for (int x = 1; x <= validsn; x++)
                            {
                                sdotransferdata[payload.cob].Add(payload.data[x]);
                            }

                            if (c == 1)
                            {

                                StringBuilder hex = new StringBuilder(sdotransferdata[payload.cob].Count * 2);
                                StringBuilder ascii = new StringBuilder(sdotransferdata[payload.cob].Count * 2);
                                foreach (byte b in sdotransferdata[payload.cob])
                                {
                                    hex.AppendFormat("{0:x2} ", b);
                                    ascii.AppendFormat("{0}", (char)Convert.ToChar(b));
                                }

                                //  textBox_info.Invoke(new MethodInvoker(delegate
                                //  {
                                //      textBox_info.AppendText(String.Format("SDO UPLOAD COMPLETE for cob 0x{0:x3}\r\n", payload.cob))
                                //
                                //      textBox_info.AppendText(hex.ToString() + "\r\n");
                                //     textBox_info.AppendText(ascii.ToString() + "\r\n\r\n");
                                //
                                //                                }));
                            }

                        }

                        break;
                    case 1:
                        mode = "download segment response";
                        sdoproto = string.Format("{0} {1}", mode, t == 1 ? "TOG ON" : "TOG OFF");
                        break;
                    case 2:
                        mode = "initate upload response";
                        string nbytes = "";

                        if (e == 1 && s == 1)
                        {
                            //n is valid
                            nbytes = string.Format("Valid bytes = {0}", 4 - n);
                        }

                        if (e == 0 && s == 1)
                        {
                            byte[] size = new byte[4];
                            Array.Copy(payload.data, 4, size, 0, 4);
                            UInt32 isize = (UInt32)BitConverter.ToUInt32(size, 0);
                            nbytes = string.Format("Bytes = {0}", isize);

                            if (sdotransferdata.ContainsKey(payload.cob))
                                sdotransferdata.Remove(payload.cob);

                            sdotransferdata.Add(payload.cob, new List<byte>());
                        }

                        sdoproto = string.Format("{0} {1} {2} 0x{3:x4}/{4:x2}", mode, nbytes, e == 1 ? "Normal" : "Expedite", index, sub);
                        break;
                    case 3:
                        mode = "initate download response";
                        sdoproto = string.Format("{0} 0x{1:x4}/{2:x2}", mode, index, sub);



                        break;

                    case 5:
                        mode = "Block download response";

                        byte segperblock = payload.data[4];
                        sdoproto = string.Format("{0} 0x{1:x4}/{2:x2} Blksize = {3}", mode, cc == 0 ? "NO SERVER CRC" : "SERVER CRC", index, sub, segperblock);

                        break;


                    default:
                        mode = string.Format("SCS {0}", SCS);
                        break;

                }



                msg = sdoproto;


            }
            else
            {
                //Client to server

                string mode = "";
                string sdoproto = "";

                switch (SCS)
                {
                    case 0:
                        mode = "download segment request";
                        sdoproto = string.Format("{0} {1} Valid bytes = {2} {3}", mode, t == 1 ? "TOG ON" : "TOG OFF", validsn, c == 0 ? "MORE" : "END");


                        if (sdotransferdata.ContainsKey(payload.cob))
                        {

                            for (int x = 1; x <= validsn; x++)
                            {
                                sdotransferdata[payload.cob].Add(payload.data[x]);
                            }

                            if (c == 1)
                            {

                                StringBuilder hex = new StringBuilder(sdotransferdata[payload.cob].Count * 2);
                                StringBuilder ascii = new StringBuilder(sdotransferdata[payload.cob].Count * 2);
                                foreach (byte b in sdotransferdata[payload.cob])
                                {
                                    hex.AppendFormat("{0:x2} ", b);
                                    ascii.AppendFormat("{0}", (char)Convert.ToChar(b));
                                }

                                //sdoproto += "\nDATA = " + hex.ToString() + "(" + ascii + ")";

                                /*  textBox_info.Invoke(new MethodInvoker(delegate
                                  {
                                      textBox_info.AppendText(String.Format("SDO DOWNLOAD COMPLETE for cob 0x{0:x3}\n", payload.cob));

                                      textBox_info.AppendText(hex.ToString() + "\n");
                                      textBox_info.AppendText(ascii.ToString() + "\n");
                                  }));*/


                                //Console.WriteLine(hex.ToString());
                                //Console.WriteLine(ascii.ToString());

                                sdotransferdata.Remove(payload.cob);
                            }
                        }


                        break;
                    case 1:
                        string nbytes = "";

                        if (e == 1 && s == 1)
                        {
                            //n is valid
                            nbytes = string.Format("Valid bytes = {0}", 4 - n);
                        }

                        if (e == 0 && s == 1)
                        {
                            byte[] size2 = new byte[4];
                            Array.Copy(payload.data, 4, size2, 0, 4);
                            UInt32 isize2 = (UInt32)BitConverter.ToUInt32(size2, 0);
                            nbytes = string.Format("Bytes = {0}", isize2);
                        }

                        mode = "initate download request";
                        sdoproto = string.Format("{0} {1} {2} 0x{3:x4}/{4:x2}", mode, nbytes, e == 1 ? "Normal" : "Expedite", index, sub);
                        if (sdotransferdata.ContainsKey(payload.cob))
                            sdotransferdata.Remove(payload.cob);

                        sdotransferdata.Add(payload.cob, new List<byte>());

                        break;
                    case 2:
                        mode = "initate upload request";
                        sdoproto = string.Format("{0} 0x{1:x4}/{2:x2}", mode, index, sub);
                        break;
                    case 3:
                        mode = "upload segment request";
                        sdoproto = string.Format("{0} {1}", mode, t == 1 ? "TOG ON" : "TOG OFF");
                        break;

                    case 5:
                        mode = "Block download";
                        sdoproto = string.Format("{0}", mode);
                        break;

                    case 6:
                        mode = "Initate Block download request";

                        byte[] size = new byte[4];
                        Array.Copy(payload.data, 4, size, 0, 4);
                        UInt32 isize = (UInt32)BitConverter.ToUInt32(size, 0);

                        sdoproto = string.Format("{0} 0x{1:x4}/{2:x2} Size = {3}", mode, cc == 0 ? "NO CLIENT CRC" : "CLIENT CRC", index, sub, isize);
                        break;


                    default:
                        mode = string.Format("CSC {0}", SCS);
                        break;

                }


                msg = sdoproto;

            }


            if ((payload.data[0] & 0x80) != 0)
            {
                byte[] errorcode = new byte[4];
                errorcode[0] = payload.data[4];
                errorcode[1] = payload.data[5];
                errorcode[2] = payload.data[6];
                errorcode[3] = payload.data[7];

                UInt32 err = BitConverter.ToUInt32(errorcode, 0);

                if (ErrorCodes.sdoerrormessages.ContainsKey(err))
                {

                    msg += " " + ErrorCodes.sdoerrormessages[err];

                }

            }
            else
            {
                if (Program.pluginManager.ipdo != null)
                    msg += " " + Program.pluginManager.ipdo.decodesdo(payload.cob, index, sub, payload.data);
            }


            items[5] = msg;
            appendfile(items);


            if (Properties.Settings.Default.showsdo)
            {
                ListViewItem i = new ListViewItem(items);

                if ((payload.data[0] & 0x80) != 0)
                {
                    i.BackColor = Color.Orange;
                }

                i.ForeColor = Color.DarkBlue;

                lock (listitems)
                    listitems.Add(i);
            }

        }

        private void log_PDO(canpacket[] payloads, DateTime dt)
        {


            foreach (canpacket payload in payloads)
            {

                string[] items = new string[6];
                items[0] = dt.ToString("MM/dd/yyyy HH:mm:ss.fff");
                items[1] = "PDO";
                items[2] = string.Format("{0:x3}", payload.cob);
                items[3] = "";
                items[4] = BitConverter.ToString(payload.data).Replace("-", string.Empty);

                if (Program.pluginManager.pdoprocessors.ContainsKey(payload.cob))
                {
                    string msg = null;
                    try
                    {
                        msg = Program.pluginManager.pdoprocessors[payload.cob](payload.data);
                    }
                    catch (Exception)
                    {
                        msg += "!! DECODE EXCEPTION !!";
                    }

                    if (msg == null)
                    {
                        continue;

                    }
                    else
                    {
                        items[5] = msg;
                    }
                }
                else
                {
                    items[5] = string.Format("Len = {0}", payload.len);
                }

                if (Properties.Settings.Default.showpdo)
                {
                    ListViewItem i = new ListViewItem(items);

                    lock (listitems)
                        listitems.Add(i);
                }

                appendfile(items);
            }

        }

        private void log_EMCY(canpacket payload, DateTime dt)
        {
            string[] items = new string[6];
            string[] items2 = new string[5];

            items[0] = dt.ToString("MM/dd/yyyy HH:mm:ss.fff");
            items[1] = "EMCY";
            items[2] = string.Format("{0:x3}", payload.cob);
            items[3] = string.Format("{0:x3}", payload.cob - 0x080);
            items[4] = BitConverter.ToString(payload.data).Replace("-", string.Empty);
            //items[4] = "EMCY";

            items2[0] = dt.ToString("MM/dd/yyyy HH:mm:ss.fff");
            items2[1] = items[2];
            items2[2] = items[3];

            UInt16 code = (UInt16)(payload.data[0] + (payload.data[1] << 8));
            byte bits = (byte)(payload.data[3]);
            UInt32 info = (UInt32)(payload.data[4] + (payload.data[5] << 8) + (payload.data[6] << 16) + (payload.data[7] << 24));

            if (ErrorCodes.errcode.ContainsKey(code))
            {

                string bitinfo;

                if (ErrorCodes.errbit.ContainsKey(bits))
                {
                    bitinfo = ErrorCodes.errbit[bits];
                }
                else
                {
                    bitinfo = string.Format("bits 0x{0:x2}", bits);
                }

                items[5] = string.Format("Error: {0} - {1} info 0x{2:x8}", ErrorCodes.errcode[code], bitinfo, info);
            }
            else
            {
                items[5] = string.Format("Error code 0x{0:x4} bits 0x{1:x2} info 0x{2:x8}", code, bits, info);
            }

            items2[3] = items[5];

            ListViewItem i = new ListViewItem(items);
            ListViewItem i2 = new ListViewItem(items2);

            i.ForeColor = Color.White;
            i2.ForeColor = Color.White;

            if (code == 0)
            {
                i.BackColor = Color.Green;
                i2.BackColor = Color.Green;

            }
            else
            {
                i.BackColor = Color.Red;
                i2.BackColor = Color.Red;

            }

            if (Properties.Settings.Default.showsdo)
            {
                lock (listitems)
                    listitems.Add(i);

            }



            appendfile(items);

        }


        Dictionary<UInt16, string> errcode = new Dictionary<ushort, string>();
        Dictionary<UInt16, string> errbit = new Dictionary<ushort, string>();

        private void interror()
        {

            errcode.Add(0x0000, "error Reset or No Error");
            errcode.Add(0x1000, "Generic Error");
            errcode.Add(0x2000, "Current");
            errcode.Add(0x2100, "device input side");
            errcode.Add(0x2200, "Current inside the device");
            errcode.Add(0x2300, "device output side");
            errcode.Add(0x3000, "Voltage");
            errcode.Add(0x3100, "Mains Voltage");
            errcode.Add(0x3200, "Voltage inside the device");
            errcode.Add(0x3300, "Output Voltage");
            errcode.Add(0x4000, "Temperature");
            errcode.Add(0x4100, "Ambient Temperature");
            errcode.Add(0x4200, "Device Temperature");
            errcode.Add(0x5000, "Device Hardware");
            errcode.Add(0x6000, "Device Software");
            errcode.Add(0x6100, "Internal Software");
            errcode.Add(0x6200, "User Software");
            errcode.Add(0x6300, "Data Set");
            errcode.Add(0x7000, "Additional Modules");
            errcode.Add(0x8000, "Monitoring");
            errcode.Add(0x8100, "Communication");
            errcode.Add(0x8110, "CAN Overrun (Objects lost)");
            errcode.Add(0x8120, "CAN in Error Passive Mode");
            errcode.Add(0x8130, "Life Guard Error or Heartbeat Error");
            errcode.Add(0x8140, "recovered from bus off");
            errcode.Add(0x8150, "CAN-ID collision");
            errcode.Add(0x8200, "Protocol Error");
            errcode.Add(0x8210, "PDO not processed due to length error");
            errcode.Add(0x8220, "PDO length exceeded");
            errcode.Add(0x8230, "destination object not available");
            errcode.Add(0x8240, "Unexpected SYNC data length");
            errcode.Add(0x8250, "RPDO timeout");
            errcode.Add(0x9000, "External Error");
            errcode.Add(0xF000, "Additional Functions");
            errcode.Add(0xFF00, "Device specific");

            errcode.Add(0x2310, "Current at outputs too high (overload)");
            errcode.Add(0x2320, "Short circuit at outputs");
            errcode.Add(0x2330, "Load dump at outputs");
            errcode.Add(0x3110, "Input voltage too high");
            errcode.Add(0x3120, "Input voltage too low");
            errcode.Add(0x3210, "Internal voltage too high");
            errcode.Add(0x3220, "Internal voltage too low");
            errcode.Add(0x3310, "Output voltage too high");
            errcode.Add(0x3320, "Output voltage too low");

            errbit.Add(0x00, "Error Reset or No Error");
            errbit.Add(0x01, "CAN bus warning limit reached");
            errbit.Add(0x02, "Wrong data length of the received CAN message");
            errbit.Add(0x03, "Previous received CAN message wasn't processed yet");
            errbit.Add(0x04, "Wrong data length of received PDO");
            errbit.Add(0x05, "Previous received PDO wasn't processed yet");
            errbit.Add(0x06, "CAN receive bus is passive");
            errbit.Add(0x07, "CAN transmit bus is passive");
            errbit.Add(0x08, "Wrong NMT command received");
            errbit.Add(0x09, "(unused)");
            errbit.Add(0x0A, "(unused)");
            errbit.Add(0x0B, "(unused)");
            errbit.Add(0x0C, "(unused)");
            errbit.Add(0x0D, "(unused)");
            errbit.Add(0x0E, "(unused)");
            errbit.Add(0x0F, "(unused)");

            errbit.Add(0x10, "(unused)");
            errbit.Add(0x11, "(unused)");
            errbit.Add(0x12, "CAN transmit bus is off");
            errbit.Add(0x13, "CAN module receive buffer has overflowed");
            errbit.Add(0x14, "CAN transmit buffer has overflowed");
            errbit.Add(0x15, "TPDO is outside SYNC window");
            errbit.Add(0x16, "(unused)");
            errbit.Add(0x17, "(unused)");
            errbit.Add(0x18, "SYNC message timeout");
            errbit.Add(0x19, "Unexpected SYNC data length");
            errbit.Add(0x1A, "Error with PDO mapping");
            errbit.Add(0x1B, "Heartbeat consumer timeout");
            errbit.Add(0x1C, "Heartbeat consumer detected remote node reset");
            errbit.Add(0x1D, "(unused)");
            errbit.Add(0x1E, "(unused)");
            errbit.Add(0x1F, "(unused)");

            errbit.Add(0x20, "Emergency message wasn't sent");
            errbit.Add(0x21, "(unused)");
            errbit.Add(0x22, "Microcontroller has just started");
            errbit.Add(0x23, "(unused)");
            errbit.Add(0x24, "(unused)");
            errbit.Add(0x25, "(unused)");
            errbit.Add(0x26, "(unused)");
            errbit.Add(0x27, "(unused)");

            errbit.Add(0x28, "Wrong parameters to CO_errorReport() function");
            errbit.Add(0x29, "Timer task has overflowed");
            errbit.Add(0x2A, "Unable to allocate memory for objects");
            errbit.Add(0x2B, "test usage");
            errbit.Add(0x2C, "Software error");
            errbit.Add(0x2D, "Object dictionary does not match the software");
            errbit.Add(0x2E, "Error in calculation of device parameters");
            errbit.Add(0x2F, "Error with access to non volatile device memory");

            sdoerrormessages.Add(0x05030000, "Toggle bit not altered");
            sdoerrormessages.Add(0x05040000, "SDO protocol timed out");
            sdoerrormessages.Add(0x05040001, "Command specifier not valid or unknown");
            sdoerrormessages.Add(0x05040002, "Invalid block size in block mode");
            sdoerrormessages.Add(0x05040003, "Invalid sequence number in block mode");
            sdoerrormessages.Add(0x05040004, "CRC error (block mode only)");
            sdoerrormessages.Add(0x05040005, "Out of memory");
            sdoerrormessages.Add(0x06010000, "Unsupported access to an object");
            sdoerrormessages.Add(0x06010001, "Attempt to read a write only object");
            sdoerrormessages.Add(0x06010002, "Attempt to write a read only object");
            sdoerrormessages.Add(0x06020000, "Object does not exist");
            sdoerrormessages.Add(0x06040041, "Object cannot be mapped to the PDO");
            sdoerrormessages.Add(0x06040042, "Number and length of object to be mapped exceeds PDO length");
            sdoerrormessages.Add(0x06040043, "General parameter incompatibility reasons");
            sdoerrormessages.Add(0x06040047, "General internal incompatibility in device");
            sdoerrormessages.Add(0x06060000, "Access failed due to hardware error");
            sdoerrormessages.Add(0x06070010, "Data type does not match, length of service parameter does not match");
            sdoerrormessages.Add(0x06070012, "Data type does not match, length of service parameter too high");
            sdoerrormessages.Add(0x06070013, "Data type does not match, length of service parameter too short");
            sdoerrormessages.Add(0x06090011, "Sub index does not exist");
            sdoerrormessages.Add(0x06090030, "Invalid value for parameter (download only).");
            sdoerrormessages.Add(0x06090031, "Value range of parameter written too high");
            sdoerrormessages.Add(0x06090032, "Value range of parameter written too low");
            sdoerrormessages.Add(0x06090036, "Maximum value is less than minimum value.");
            sdoerrormessages.Add(0x060A0023, "Resource not available: SDO connection");
            sdoerrormessages.Add(0x08000000, "General error");
            sdoerrormessages.Add(0x08000020, "Data cannot be transferred or stored to application");
            sdoerrormessages.Add(0x08000021, "Data cannot be transferred or stored to application because of local control");
            sdoerrormessages.Add(0x08000022, "Data cannot be transferred or stored to application because of present device state");
            sdoerrormessages.Add(0x08000023, "Object dictionary not present or dynamic generation fails");
            sdoerrormessages.Add(0x08000024, "No data available");


        }



        private void button_open_Click(object sender, EventArgs e)
        {
            try
            {
                lco.close();

                if (button_open.Text == "Close")
                {


                    if (sw != null)
                        sw.Close();

                    button_open.BackColor = Color.Green;
                    button_open.Text = "Open";

                    textBox_info.AppendText("PORT CLOSED\r\n");
                    return;
                }

                if (comboBox_port.SelectedItem == null)
                {
                    comboBox_port.Text = "";
                    return;
                }

#if false
                driverport dp = (driverport)comboBox_port.SelectedItem;

                textBox_info.AppendText(String.Format("Trying to open port {0} using driver {1} \r\n", dp.port, dp.driver));

                int rate = comboBox_rate.SelectedIndex;


                string port = dp.port;

                if (dp.port.Contains("USB"))
                {
                    sComPortModel s = cpm.requestSerialPortById(dp.VID, dp.PID, "", "");
                    s.DeviceConnected += S_DeviceConnected;
                    s.DeviceDisconnected += S_DeviceDisconnected;
                    port = s.port;
                }

                lco.open(port, (BUSSPEED)rate, dp.driver);
#else
                lco.open(this.m_PcanHandle, this.m_PcanBaudRate);
#endif

                if (lco.isopen())
                {
                    button_open.Text = "Close";
                    button_open.BackColor = Color.Red;
                    //fixme make this user selectable from GUI
                    //sw = new StreamWriter("canlog.txt", true);

                    textBox_info.AppendText("Success port open\r\n");

#if false
                    Properties.Settings.Default.lastport = dp.port;
                    Properties.Settings.Default.lastdriver = dp.driver;
                    Properties.Settings.Default.lastrate = comboBox_rate.Text;
#endif

                    Properties.Settings.Default.Save();

                    foreach (KeyValuePair<string, object> kvp in plugins)
                    {
                        IInterfaceService iis = (IInterfaceService)kvp.Value;
                        iis.setlco(lco);
                    }

                }
                else
                {
                    button_open.Text = "Open";
                    button_open.BackColor = Color.Green;
                    textBox_info.AppendText("ERROR opening port\r\n");
                }
            }
            catch (Exception ex)
            {
                textBox_info.AppendText("ERROR opening port " + ex.ToString() + "\r\n");
                MessageBox.Show("Setup error " + ex.ToString());

            }

            button_open.Enabled = true;
        }

        private void S_DeviceDisconnected(object sender, EventArgs e)
        {
            comboBox_rate.Invoke(new MethodInvoker(delegate
            {

                if (lco.isopen())
                    lco.close();

                button_open.Text = "Open";

            }));
        }

        private void S_DeviceConnected(object sender, EventArgs e)
        {

            comboBox_rate.Invoke(new MethodInvoker(delegate
            {

                sComPortModel s = (sComPortModel)sender;
                int rate = comboBox_rate.SelectedIndex;

                lco.close();

#if false /// TODO
                //FIXME hardcoded driver
                lco.open(s.port, (BUSSPEED)rate, "drivers\\can_canusb_win32");
#endif

                button_open.Text = "Close";
                textBox_info.AppendText("Success port open\r\n");

            }));

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void button_clear_Click(object sender, EventArgs e)
        {



            lock (EMClistitems)
            {
                EMClistitems.Clear();
                listView_emcy.Items.Clear();
            }

            lock (dirtyNMTstates)
            {
                NMTstate.Clear();
                dirtyNMTstates.Clear();
                listView_nmt.Items.Clear();
            }

        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBox_port_SelectedIndexChanged(object sender, EventArgs e)
        {
            SettingsMgr.settings.options.selectedport = comboBox_port.SelectedItem.ToString();

            string strTemp = comboBox_port.Text;
            strTemp = strTemp.Substring(strTemp.IndexOf('B') + 1, 1);
            strTemp = "0x5" + strTemp;
            this.m_PcanHandle = Convert.ToByte(strTemp, 16);
        }

        private void comboBox_rate_SelectedIndexChanged(object sender, EventArgs e)
        {
            SettingsMgr.settings.options.selectedrate = comboBox_rate.SelectedIndex;

            TPCANBaudrate[] baudValue = (TPCANBaudrate[])Enum.GetValues(typeof(TPCANBaudrate));
            this.m_PcanBaudRate = baudValue[this.comboBox_rate.SelectedIndex];
        }

        private void Form1_Load(object sender, EventArgs e)
        {



            comboBox_rate.SelectedIndex = SettingsMgr.settings.options.selectedrate;
            comboBox_port.SelectedItem = SettingsMgr.settings.options.selectedport;

            //read git version string, show in title bar 
            //(https://stackoverflow.com/a/15145121)
            string gitVersion = String.Empty;
            using (Stream stream = System.Reflection.Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("CanMonitor." + "version.txt"))
            using (StreamReader reader = new StreamReader(stream))
            {
                gitVersion = reader.ReadToEnd();
            }
            if (gitVersion == "")
            {
                gitVersion = "Unknown";
            }
            this.Text += " -- " + gitVersion;
            this.gitVersion = gitVersion;

            var mruFilePath = Path.Combine(appdatafolder, "PLUGINMRU.txt");
            if (System.IO.File.Exists(mruFilePath))
                _mru.AddRange(System.IO.File.ReadAllLines(mruFilePath));

            populateMRU();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.autoscroll)
            {
                if (listView1.Items.Count > 1)
                    listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }



        #region pluginloader

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

        }


        public void dosave(string filename)
        {
            XElement xeRoot = new XElement("CanOpenMonitor");

            foreach (ListViewItem i in listView1.Items)
            {
                XElement xeRow = new XElement("Packet", new XAttribute("backcol", i.BackColor.Name), new XAttribute("forcol", i.ForeColor.Name));
                int x = 0;

                foreach (ListViewItem.ListViewSubItem subItem in i.SubItems)
                {
                    XElement xeCol = new XElement(listView1.Columns[x].Text);
                    xeCol.Value = subItem.Text;
                    xeRow.Add(xeCol);
                    // To add attributes use XAttributes
                    x++;
                }
                xeRoot.Add(xeRow);

            }

            xeRoot.Save(filename);


        }


        public void doload()
        {
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.ShowHelp = true;
            sfd.Filter = "(*.xml)|*.xml";

            try
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {

                    listView1.Items.Clear();

                    XElement xeRoot = XElement.Load(sfd.FileName);
                    XName Packet = XName.Get("Packet");


                    foreach (var packetelement in xeRoot.Elements(Packet))
                    {
                        XName XTimestamp = XName.Get("Timestamp");
                        XName XType = XName.Get("Type");
                        XName XCob = XName.Get("COB");
                        XName XNodeT = XName.Get("Node");
                        XName XPayload = XName.Get("Payload");
                        XName XInfo = XName.Get("Info");

                        string[] bits = new string[6];

                        bits[0] = packetelement.Element(XTimestamp).Value;
                        bits[1] = packetelement.Element(XType).Value;
                        bits[2] = packetelement.Element(XCob).Value;
                        bits[3] = packetelement.Element(XNodeT).Value;
                        bits[4] = packetelement.Element(XPayload).Value;
                        bits[5] = packetelement.Element(XInfo).Value;

                        string cobx = bits[2].ToUpper();
                        UInt16 cob = Convert.ToUInt16(cobx, 16);

                        byte[] b = new byte[bits[4].Length / 2];
                        for (int x = 0; x < bits[4].Length / 2; x++)
                        {
                            string s = bits[4].Substring(x * 2, 2);
                            b[x] = byte.Parse(s, System.Globalization.NumberStyles.HexNumber);
                        }

                        canpacket[] p = new canpacket[1];
                        p[0] = new canpacket();
                        p[0].cob = cob;
                        p[0].data = b;
                        p[0].len = (byte)b.Length;



                        string d = bits[0];
                        string[] d2 = d.Split(' ');
                        string[] d3 = d2[0].Split('/');
                        string[] d4 = d2[1].Split(':');


                        //  11/16/2023 17:56:53.474
                        //  "2018-08-18T07:22:16.0000000Z"
                        string d5 = $"{d3[2]}-{d3[0]}-{d3[1]}T{d4[0]}:{d4[1]}:{d4[2]}";


                        DateTime dt = DateTime.Parse(d5);

                        switch (bits[1])
                        {
                            case "PDO":
                                log_PDO(p, dt);
                                break;

                            case "SDO":
                                log_SDO(p[0], dt);
                                break;

                            case "NMT":
                                log_NMT(p[0], dt);
                                break;

                            case "NMTEC":
                                log_NMTEC(p[0], dt);
                                break;

                            case "EMCY":
                                log_EMCY(p[0], dt);
                                break;
                        }

                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());
                listView1.EndUpdate();

            }

        }


        public void doload2()
        {
            OpenFileDialog sfd = new OpenFileDialog();
            sfd.ShowHelp = true;
            sfd.Filter = "(*.log)|*.log";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string[] lines = File.ReadAllLines(sfd.FileName);

                foreach (string line in lines)
                {
                    try
                    {

                        string[] bits = line.Split(',');
                        UInt16 cob = Convert.ToUInt16(bits[1], 16);
                        byte len = Convert.ToByte(bits[2], 16);


                        canpacket[] p = new canpacket[1];
                        p[0] = new canpacket();
                        p[0].cob = cob;


                        p[0].data = new byte[len];
                        for (int x = 0; x < len; x++)
                        {
                            p[0].data[x] = Convert.ToByte(bits[3].Substring(x * 2, 2), 16);
                        }

                        p[0].len = len;

                        DateTime dt = DateTime.Parse(bits[0]);

                        if (cob < 0x80)
                            log_NMT(p[0], dt);

                        if (cob >= 0x80 && cob < 0x100)
                            log_EMCY(p[0], dt);

                        if (cob >= 0x180 && cob < 0x580)
                            log_PDO(p, dt);

                        if (cob >= 0x580 && cob < 0x700)
                            log_PDO(p, dt);

                        if (cob >= 0x700)
                            log_NMTEC(p[0], dt);


                    }
                    catch (Exception)
                    {

                    }

                }
            }
        }




        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button_refresh_Click(object sender, EventArgs e)
        {
#if false
            enumerateports();
#else
            this.enumerateports((SupportedVendor)this.comboBox_vendor.SelectedIndex);
#endif
        }


        private void enumerateports()
        {
            comboBox_port.Text = "";
            comboBox_port.Items.Clear();

            textBox_info.AppendText("\r\nEnumerating ports....\r\n\r\n");

            foreach (string s in drivers)
            {
                textBox_info.AppendText(String.Format("Attempting to enumerate with driver {0}\r\n", s));

                try
                {
                    lco.enumerate(s);
                }
                catch (Exception e)
                {
                    textBox_info.AppendText(e.ToString() + "\r\n");
                }
            }

            textBox_info.AppendText("\r\n");

            foreach (KeyValuePair<string, List<string>> kvp in lco.ports)
            {
                List<string> ps = kvp.Value;


                textBox_info.AppendText($"Driver {kvp.Key} has {kvp.Value.Count} ports\r\n");

                foreach (string s in ps)
                {
                    textBox_info.AppendText(string.Format("Found port {0}\r\n", s));
                    driverport dp = new driverport();
                    dp.port = s;
                    dp.driver = kvp.Key;
                    comboBox_port.Items.Add(dp);
                }

                textBox_info.AppendText("\r\n");
            }

            textBox_info.AppendText("\r\n");

            List<sComPortModel> psx = cpm.GetPorts();

            foreach (sComPortModel p in psx)
            {
                driverport dp = new driverport();
                dp.port = string.Format($"USB/VID_{p.vid}/PID_{p.pid}");
                dp.PID = p.pid;
                dp.VID = p.vid;
                dp.driver = "drivers\\can_canusb_win32";
                comboBox_port.Items.Add(dp);
            }
        }

        private void enumerateports(SupportedVendor vendor)
        {
            comboBox_port.Text = "";
            comboBox_port.Items.Clear();


            textBox_info.AppendText("Enumerating ports...." + Environment.NewLine);

            switch (vendor)
            {
#if false
                case SupportedVendor.CANFESTIVAL: {
                    foreach (string s in drivers) {
                        textBox_info.AppendText(String.Format("Attempting to enumerate with driver {0}", s) + Environment.NewLine);

                        lco.enumerate(s);
                    }

                    foreach (KeyValuePair<string, List<string>> kvp in lco.ports) {
                        List<string> ps = kvp.Value;

                        foreach (string s in ps) {

                            textBox_info.AppendText(string.Format("Found port {0}", s) + Environment.NewLine);

                            driverport dp = new driverport();
                            dp.port = s;
                            dp.driver = kvp.Key;

                            comboBox_port.Items.Add(dp);
                        }
                    }
                    break;
                }
#endif
                case SupportedVendor.PEAK:
                    {
                        UInt32 iBuffer;
                        TPCANStatus stsResult;
                        try
                        {
                            for (int i = 0; i < m_HandlesArray.Length; i++)
                            {
                                // Includes all no-Plug&Play Handles
                                if ((m_HandlesArray[i] >= PCANBasic.PCAN_USBBUS1) &&
                                    (m_HandlesArray[i] <= PCANBasic.PCAN_USBBUS8))
                                {
                                    stsResult = PCANBasic.GetValue(m_HandlesArray[i], TPCANParameter.PCAN_CHANNEL_CONDITION, out iBuffer, sizeof(UInt32));
                                    if ((stsResult == TPCANStatus.PCAN_ERROR_OK) && (iBuffer == PCANBasic.PCAN_CHANNEL_AVAILABLE))
                                    {
                                        this.comboBox_port.Items.Add(string.Format("PCAN-USB{0}", m_HandlesArray[i] & 0x0F));
                                    }
                                }
                            }
                            this.comboBox_port.SelectedIndex = this.comboBox_port.Items.Count - 1;

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Environment.Exit(-1);
                        }

                        string[] baudName = Enum.GetNames(typeof(TPCANBaudrate));
                        this.comboBox_rate.Items.Clear();
                        for (int i = 0; i < baudName.Length; i++)
                        {
                            this.comboBox_rate.Items.Add(baudName[i]);
                        }
                        this.comboBox_rate.SelectedIndex = 3;
                        break;
                    }

                default:
                    break;
            }
        }

        private void preferencesToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Prefs p = new Prefs(appdatafolder, assemblyfolder);
            p.ShowDialog();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                // Hide();
                notifyIcon1.Visible = true;
                notifyIcon1.BalloonTipText = "Can monitor is still running";
                notifyIcon1.BalloonTipTitle = "Can Monitor";
                notifyIcon1.ShowBalloonTip(2000);
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }




        #endregion

        #region controls


        #endregion
#endregion

    }

}
