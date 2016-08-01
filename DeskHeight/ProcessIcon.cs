using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Timers;
using DeskHeight.Properties;

namespace DeskHeight {
    class ProcessIcon:IDisposable {
        NotifyIcon ni;

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
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        public void Dispose() {
            // When the application closes, this will remove the icon from the system tray immediately.
            ni.Dispose();
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
            // Ref: http://stackoverflow.com/a/17937051/877387
            //TODO: This is where the update code would be called and if the desk is up, change the icon
            DateTime now = DateTime.Now;

            if(now.Minute % 2 == 0){
                ni.Icon = Resources.DeskUp;

            } else {
                ni.Icon = Resources.DeskDown;
            }
        }
    }
}
