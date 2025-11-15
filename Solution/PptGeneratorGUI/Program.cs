using System;
using System.Windows.Forms;

namespace PptGeneratorGUI
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Application.Run(new frmMain());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore non gestito: {ex}", "Errore non gestito", MessageBoxButtons.OK);
            }
        }
    }
}