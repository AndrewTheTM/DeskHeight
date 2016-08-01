using System;
using System.Windows.Forms;

namespace DeskHeight {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (ProcessIcon pi = new ProcessIcon()) {
                pi.Display();

                // Make sure the application runs!
                Application.Run();
                
            }
            
        }

    }
}
