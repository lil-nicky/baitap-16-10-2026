
using System.Drawing;
using System.Windows.Forms;

namespace ClinicAppointmentManager
{
    partial class Form1
    {
        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.AutoScaleDimensions = new SizeF(8F, 16F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1200, 720);
            this.MinimumSize = new Size(950, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Clinic Appointment Manager";

            this.ResumeLayout(false);
        }
    }
}
