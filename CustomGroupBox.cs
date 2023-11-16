
using System.Drawing;
using System.Windows.Forms;

namespace VSTO_Addins
{

    public class CustomGroupBox : GroupBox
    {

        private readonly TextFormatFlags flags = TextFormatFlags.Top | TextFormatFlags.Left | TextFormatFlags.LeftAndRightPadding | TextFormatFlags.EndEllipsis;
        private Color _Bordercolor = SystemColors.Window;

        public Color BorderColor
        {
            get
            {
                return _Bordercolor;

            }
            set
            {
                _Bordercolor = value;
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            int mTxt = TextRenderer.MeasureText(e.Graphics, Text, Font, ClientSize).Height / 2 + 2;
            var r = new Rectangle(0, mTxt, ClientSize.Width, ClientSize.Height - mTxt);
            ControlPaint.DrawBorder(e.Graphics, r, BorderColor, ButtonBorderStyle.Solid);

            var textrect = Rectangle.Inflate(ClientRectangle, -4, 0);
            TextRenderer.DrawText(e.Graphics, Text, Font, textrect, ForeColor, BackColor, flags);

        }

        private void InitializeComponent()
        {
            SuspendLayout();
            ResumeLayout(false);

        }
    }
}