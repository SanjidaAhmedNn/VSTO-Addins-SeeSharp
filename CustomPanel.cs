using System;
using System.ComponentModel;
using System.Drawing;

namespace VSTO_Addins
{

    public class CustomPanel : System.Windows.Forms.Panel
    {
        public CustomPanel()
        {
            BorderStyle = System.Windows.Forms.BorderStyle.None;
            Paint += MyPanel_Paint;
        }
        private int bWidth;
        [Category("Appearance")]
        [Description("Change border width")]
        public int BorderWidth
        {
            get
            {
                return bWidth;
            }
            set
            {
                bWidth = Math.Abs(value);
                Refresh();
            }
        }
        private Color bColor;
        [Category("Appearance")]
        [Description("Change border color")]
        public Color BorderColor
        {
            get
            {
                return bColor;
            }
            set
            {
                bColor = value;
                Refresh();
            }
        }
        public virtual void MyPanel_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            e.Graphics.DrawRectangle(new Pen(bColor, bWidth), ClientRectangle);
        }
    }
}