using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Timers;
using DeskHeight.Properties;
using mcp2221_dll_m;
using System.Data.OleDb;
using System.IO;
using System.Data.SQLite;

namespace DeskHeight {
    class ProcessIcon:IDisposable {
        NotifyIcon ni;
        IntPtr Mcp;
        bool active = true;
        string dbFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DeskHeight\\DH.sqlite";
        bool deskUp = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProcessIcon"/> class.
        /// </summary>
        public ProcessIcon() {
            // Instantiate the NotifyIcon object.
            ni = new NotifyIcon();
        }

        /// <summary>
        /// Displays the icon in the system tray.
        /// </summary>
        public void Display() {
            // Put the icon in the system tray and allow it react to mouse clicks.			
            ni.MouseClick += new MouseEventHandler(ni_MouseClick);
            ni.Icon = Resources.DeskDown;
            ni.Text = "Desk Height Data Collection Program";
            ni.Visible = true;

            // Attach a context menu.
            ni.ContextMenuStrip = new ContextMenus().Create();

            //Adding the timer here (trying)
            System.Timers.Timer aTimer = new System.Timers.Timer();
            aTimer.Elapsed += new ElapsedEventHandler(onTimedEvent);
            aTimer.Interval = 5000;
            aTimer.Enabled = true;

            // Connect to MCP2221
            uint nDevices = 0;
            
            MCP2221.M_Mcp2221_GetConnectedDevices(0x04D8, 0x00DD, ref nDevices);
            //Mcp = MCP2221.M_Mcp2221_OpenBySN(0x04D8, 0x00DD, "0000071521"); // This returned a pointer of -1, Opening by index did not
                                // Did double-check the serial number, it matches the MCP2221 utility.
            Mcp = MCP2221.M_Mcp2221_OpenByIndex(0x04D8, 0x00DD, 0);

            // Turn off when computer is locked
            Microsoft.Win32.SystemEvents.SessionSwitch += new Microsoft.Win32.SessionSwitchEventHandler(SystemEvents_SessionSwitch);

            // Check for SQL file, check tables
            
            try {
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DeskHeight")) {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DeskHeight");
                }

                SQLiteConnection.CreateFile(dbFile);
                SQLiteConnection connDH = new SQLiteConnection("Data Source=" + dbFile + ";Version=3;");
                connDH.Open();
                string strCommand = Resources.createTables_0001;
                SQLiteCommand cmdDH = new SQLiteCommand();
                cmdDH.Connection = connDH;
                cmdDH.CommandText = strCommand;
                cmdDH.ExecuteNonQuery();
                connDH.Close();
            } catch (Exception inex) {
                Debug.Print(inex.Message);
            }
            writeEvent(5);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        public void Dispose() {
            // When the application closes, this will remove the icon from the system tray immediately.
            ni.Dispose();
            MCP2221.M_Mcp2221_Close(Mcp);
            writeEvent(6);
        }

        /// <summary>
        /// Handles the MouseClick event of the ni control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Forms.MouseEventArgs"/> instance containing the event data.</param>
        void ni_MouseClick(object sender, MouseEventArgs e) {
            // Handle mouse button clicks.
            if (e.Button == MouseButtons.Left) {
                // Start Windows Explorer.
                Process.Start("explorer", null);
            }
        }

        void onTimedEvent(object source, ElapsedEventArgs e) {

            if (!active)
                return;

            

            Byte[] i2cRxData = { 0 };
            Byte[] i2cWriteAlert = { 0 };
            int i2cWriteSuccess = MCP2221.M_Mcp2221_I2cWrite(Mcp, 1, 0x9A, 0, i2cWriteAlert);

            if (i2cWriteSuccess == 0) {
                int i2cReadSuccess = MCP2221.M_Mcp2221_I2cRead(Mcp, 1, 0x9B, 0, i2cRxData);
                if (i2cReadSuccess == 0) {
                    double tempC = Convert.ToDouble(i2cRxData[0]);
                    double tempF = 1.8 * tempC + 32;
                }
            }

            uint[] adcData = new uint[3];
            
            int adcReadSuccess = MCP2221.M_Mcp2221_GetAdcData(Mcp, adcData);

            // 150 is the break point
            // - 95/96 when up, cx when down
            int adc = Convert.ToInt32(adcData[0]);
            if (adc < 300) {
                ni.Icon = Resources.DeskUp;
                if (!deskUp) {
                    writeEvent(2);
                    deskUp = true;
                }

            } else {
                ni.Icon = Resources.DeskDown;
                if (deskUp) {
                    writeEvent(3);
                    deskUp = false;
                }
            }

            int a = 1;

            // Ref: http://stackoverflow.com/a/17937051/877387
            
        }

        void writeEvent(int eventNote) {
            SQLiteConnection connDH = new SQLiteConnection("Data Source=" + dbFile + ";Version=3;");
            try {
                connDH.Open();
            } catch (Exception ex) {
                Debug.Print(ex.Message);
            }


            double dtNow = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
            string strCommand = "INSERT INTO Today (EventID, EventTime) VALUES(" + eventNote + ", " + dtNow + ");";
            SQLiteCommand cmdDH = new SQLiteCommand();
            cmdDH.Connection = connDH;
            cmdDH.CommandText = strCommand;
            cmdDH.ExecuteNonQuery();
            connDH.Close();
        }

        void SystemEvents_SessionSwitch(object sender, Microsoft.Win32.SessionSwitchEventArgs e) {
            if (e.Reason == Microsoft.Win32.SessionSwitchReason.SessionLock) {
                active = false;
                writeEvent(4);
            }
            if (e.Reason == Microsoft.Win32.SessionSwitchReason.SessionUnlock) {
                active = true;
                writeEvent(1);
            }
        }
    }
}
