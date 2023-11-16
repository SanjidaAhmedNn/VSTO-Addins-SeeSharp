using System.Drawing;
using System.Windows.Forms;

namespace VSTO_Addins
{

    public class CustomButton : Button
    {

        private Color _borderColor = Color.Black;
        private int _borderWidth = 1;

        public Color BorderColor
        {
            get
            {
                return _borderColor;
            }
            set
            {
                _borderColor = value;
                Invalidate(); // Forces control to be redrawn
            }
        }

        public int BorderWidth
        {
            get
            {
                return _borderWidth;
            }
            set
            {
                _borderWidth = value;
                Invalidate(); // Forces control to be redrawn
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Create border using BorderColor and BorderWidth properties
            var borderPen = new Pen(_borderColor, _borderWidth);
            var borderRectangle = new Rectangle(0, 0, ClientSize.Width - 1, ClientSize.Height - 1);

            // Draw border
            e.Graphics.DrawRectangle(borderPen, borderRectangle);
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ResumeLayout(false);

        }
    }
}