using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace arbioApp.Utils.Controls
{
    internal class RoundedTableLayoutPanel : TableLayoutPanel
    {

        public int BorderRadius { get; set; } = 20; // Rayon des coins

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Création d'un GraphicsPath pour dessiner les coins arrondis
            using (GraphicsPath path = new GraphicsPath())
            {
                path.AddArc(new Rectangle(0, 0, BorderRadius, BorderRadius), 180, 90);
                path.AddArc(new Rectangle(Width - BorderRadius, 0, BorderRadius, BorderRadius), 270, 90);
                path.AddArc(new Rectangle(Width - BorderRadius, Height - BorderRadius, BorderRadius, BorderRadius), 0, 90);
                path.AddArc(new Rectangle(0, Height - BorderRadius, BorderRadius, BorderRadius), 90, 90);
                path.CloseAllFigures();

                // Appliquer le GraphicsPath à la région du TableLayoutPanel
                this.Region = new Region(path);

                // Désactivation de la bordure
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ResumeLayout(false);

        }
    }
}
