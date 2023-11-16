using System;
using System.Drawing;
using System.Windows.Forms;

namespace VSTO_Addins
{

    public class CustomLabel : Label
    {

        private Color _borderColor = Color.Black;
        private double _borderWidth = 0.1d;

        public Color BorderColor
        {
            get
            {
                return _borderColor;
            }
            set
            {
                _borderColor = value;
                Invalidate(); // Redraw the control
            }
        }

        public int BorderWidth
        {
            get
            {
                return (int)Math.Round(_borderWidth);
            }
            set
            {
                _borderWidth = value;
                Invalidate(); // Redraw the control
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var borderRect = new Rectangle(new Point((int)Math.Round(_borderWidth), (int)Math.Round(_borderWidth)), new Size((int)Math.Round(ClientSize.Width - 2d * _borderWidth), (int)Math.Round(ClientSize.Height - 2d * _borderWidth)));

            // Draw border
            using (var borderPen = new Pen(_borderColor, (float)_borderWidth))
            {
                e.Graphics.DrawRectangle(borderPen, borderRect);
            }
        }
    }
}