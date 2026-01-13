using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PptGeneratorGUI
{
    public partial class frmWaiting : Form
    {
        public frmWaiting()
        {
            InitializeComponent();
        }

        private void frmWaiting_Load(object sender, EventArgs e)
        {
            // Dimensione del form
            //this.Size = new Size(200, 100);

            this.Text = "Processing in progress...";

            // Ottiene l'area visibile dello schermo (esclude taskbar)
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;

            // Calcola la posizione in basso a destra
            this.Location = new Point(
                workingArea.Right - this.Width,
                workingArea.Bottom - this.Height
            );
        }
    }
}
