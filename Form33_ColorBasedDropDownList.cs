using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form33_ColorBasedDropDownList
    {
        private Excel.Application _excelApp;

        private Excel.Application excelApp
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _excelApp;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_excelApp != null)
                {
                    _excelApp.SheetSelectionChange -= excelApp_SheetSelectionChange;
                }

                _excelApp = value;
                if (_excelApp != null)
                {
                    _excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;
                }
            }
        }
        private Excel.Workbook workBook;
        public static Excel.Worksheet workSheet;
        private Excel.Worksheet workSheet2;
        private Excel.Worksheet workSheet3;
        private Excel.Range src_rng;
        public Excel.Range des_rng;
        private Excel.Range selectedRange;
        public string ax;

        private int opened;
        private Point objectPosition = new Point(); // For 2D
        public object mybtn;
        public bool focuschange;
        public Form42 form = null;
        private bool flag = false;
        public Form43 form2 = null;
        private bool flag2 = false;

        public Form33_ColorBasedDropDownList()
        {
            InitializeComponent();
        }



        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;
        // Declare the tooltip at class level
        private ToolTip tooltip = new ToolTip();
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_OK.PerformClick();


            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // ReDim mybtn(List_Preview.Items.Count)
            // MsgBox
            KeyPreview = true;
            TB_des_rng.Enabled = false;
            Selection_destination.Enabled = false;

            // Define the first 42 colors from the Visual Studio Custom tab with their names
            var vsColors = new Dictionary<string, Color>() { { "White", Color.FromArgb(255, 255, 255) }, { "Aqua Light", Color.FromArgb(228, 239, 240) }, { "Blue Light", Color.FromArgb(127, 127, 127) }, { "Rose", Color.FromArgb(250, 214, 212) }, { "Light Yellow", Color.FromArgb(255, 255, 235) }, { "Lavender", Color.FromArgb(255, 153, 255) }, { "Lime", Color.FromArgb(233, 249, 198) }, { "Light Gray", Color.FromArgb(217, 217, 217) }, { "Aqua", Color.FromArgb(188, 215, 218) }, { "Light Torquoise", Color.FromArgb(102, 204, 255) }, { "Light Red", Color.FromArgb(245, 174, 169) }, { "Light Medium Yellow", Color.FromArgb(255, 255, 204) }, { "Pink", Color.FromArgb(255, 153, 204) }, { "Light Green", Color.FromArgb(200, 249, 207) }, { "Gray", Color.FromArgb(166, 166, 166) }, { "Teal", Color.FromArgb(122, 174, 181) }, { "Blue", Color.FromArgb(51, 102, 255) }, { "Medium Red", Color.FromArgb(241, 133, 127) }, { "Yellow", Color.FromArgb(255, 255, 153) }, { "Medium Pink", Color.FromArgb(255, 51, 204) }, { "Medium Green", Color.FromArgb(91, 138, 212) }, { "Dark Gray", Color.FromArgb(20, 26, 26) }, { "Aqua Medium", Color.FromArgb(71, 121, 128) }, { "Royal Blue", Color.FromArgb(0, 0, 255) }, { "Red", Color.FromArgb(183, 30, 21) }, { "Medium Yellow", Color.FromArgb(255, 255, 102) }, { "Dark Pink", Color.FromArgb(204, 51, 153) }, { "Green", Color.FromArgb(51, 153, 102) }, { "Black", Color.FromArgb(64, 64, 64) }, { "Dark Aqua", Color.FromArgb(49, 83, 88) }, { "Medium Dark Blue", Color.FromArgb(0, 51, 153) }, { "Dark Red", Color.FromArgb(122, 20, 14) }, { "Dark Yellow", Color.FromArgb(255, 255, 0) }, { "Plum", Color.FromArgb(153, 44, 98) }, { "Medium Dark Green", Color.FromArgb(15, 141, 33) }, { "Dark Black", Color.FromArgb(13, 13, 13) }, { "Dark Teal", Color.FromArgb(20, 26, 26) }, { "Dark Blue", Color.FromArgb(0, 32, 96) }, { "Brown", Color.FromArgb(106, 53, 12) }, { "Gold", Color.FromArgb(255, 204, 0) }, { "Dark Purple", Color.FromArgb(128, 0, 128) }, { "Dark Green", Color.FromArgb(10, 94, 22) } };    // 1
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 2
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 3
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 4
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 5
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 6
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 7
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 8
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 9
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 10
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 11
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 12
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 13
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 14
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 15
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 16
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 17
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 18
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 19
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 20
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 21
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 22
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 23
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 24
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 25
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 26
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 27
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 28
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 29
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 30
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 31
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 32
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 33
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 34
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 35
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 36
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 37
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 38
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 39
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 40
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 41
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // 42
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // ... Add more colors if you have them
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 // }

            // Generate color palette buttons
            foreach (var colorEntry in vsColors)
            {
                var btn = new Button()
                {
                    Width = 20,
                    Height = 20,
                    BackColor = colorEntry.Value,
                    Tag = colorEntry.Key,  // Store color name in the Tag property
                    FlatStyle = FlatStyle.Popup
                };

                // Attach events to the button
                btn.Click += ColorButton_Click;
                btn.MouseHover += ColorButton_MouseHover;

                // Add the button to the flow layout panel
                FlowLayoutPanel1.Controls.Add(btn);
            }



            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                opened = opened + 1;

                if (excelApp.Selection is not null)
                {
                    selectedRange = (Excel.Range)excelApp.Selection;
                    src_rng = selectedRange;
                    TB_src_rng.Text = selectedRange.get_Address();

                }
                TB_src_rng.Focus();
            }

            catch (Exception ex)
            {
                TB_src_rng.Focus();
            }


        }

        private void ColorButton_Click(object sender, EventArgs e)
        {

            Button clickedButton = (Button)sender;
            objectPosition = clickedButton.Location;
            Color c;
            // MsgBox(8)
            c = clickedButton.BackColor;


            int index = List_Preview.SelectedIndex;
            if (index < 0)
                return;

            ColoredItem item = (ColoredItem)List_Preview.Items[index];
            item.Color = c;
            clickedButton.Focus();

            this.mybtn((object)List_Preview.SelectedIndex) = clickedButton;
            Btn_color.BackColor = c;
            Refresh();
        }
        private void List_Box_IndexChanged()
        {

            ColoredItem item = (ColoredItem)List_Preview.Items[List_Preview.SelectedIndex];

            if (item.Color == Color.White)
            {
                Btn_NC.Focus();
            }
            else
            {
                this.mybtn((object)List_Preview.SelectedIndex).Focus();
            }
        }

        private void ListBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
                return;

            ColoredItem item = (ColoredItem)List_Preview.Items[e.Index];

            var textColor = Color.Black;
            var backColor = item.Color;

            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected & item.Color == Color.White)
            {
                // If item is selected, we'll use system colors to highlight.
                backColor = SystemColors.Highlight;
                textColor = SystemColors.HighlightText;
            }

            // Use the determined colors.
            using (var brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            // Draw the text in the determined text color.
            using (var brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(item.Text, e.Font, brush, e.Bounds.Left, e.Bounds.Top);
            }

            e.DrawFocusRectangle();
        }

        private void ColorButton_MouseHover(object sender, EventArgs e)
        {
            Button hoveredButton = (Button)sender;
            // Display the color name from the Tag property of the button
            tooltip.SetToolTip(hoveredButton, hoveredButton.Tag.ToString());
        }


        private void Button2_Click(object sender, EventArgs e)
        {
            if (ColorDialog1.ShowDialog() != DialogResult.Cancel)
            {
                // Label1.ForeColor = ColorDialog1.Color

                Button clickedButton = (Button)sender;
                objectPosition = clickedButton.Location;
                int index = List_Preview.SelectedIndex;

                ColoredItem item = (ColoredItem)List_Preview.Items[index];
                item.Color = ColorDialog1.Color;
                Button2.Focus();

                this.mybtn((object)List_Preview.SelectedIndex) = Button2;
                Btn_color.BackColor = ColorDialog1.Color;
                Refresh();
            }
        }

        private void Selection_source_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedRange is null)
                {
                }
                else
                {

                    // MsgBox(List_Preview.Items.Count)
                    TB_src_rng.Text = selectedRange.get_Address();


                    // FocusedTextBox = 1
                    Hide();

                    excelApp = Globals.ThisAddIn.Application;
                    workBook = excelApp.ActiveWorkbook;

                    Excel.Range userInput = (Excel.Range)excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type: 8);
                    src_rng = userInput;

                    string sheetName;
                    sheetName = Strings.Split(src_rng.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true), "]")[1];
                    sheetName = Strings.Split(sheetName, "!")[0];

                    if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                    {
                        sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                    }
                    workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                    workSheet.Activate();

                    src_rng.Select();

                    TB_src_rng.Text = src_rng.get_Address();

                    Show();
                    TB_src_rng.Focus();

                    Excel.Range ran = (Excel.Range)src_rng[1, 1];



                    // Clear the ListBox
                    List_Preview.Items.Clear();
                    // MsgBox(ran.Address)

                    // If range.Validation.Type = Excel.XlDVType.xlValidateList Then
                    string formula = ran.Validation.Formula1;
                    // MsgBox(formula)
                    var items = new List<string>();
                    if (formula.Contains(":"))
                    {
                        var range = excelApp.get_Range(formula);
                        foreach (var r in range)
                            items.Add(r.Value.ToString());
                    }
                    else
                    {
                        // Else, split the formula to get the individual items
                        items.AddRange(formula.Split(','));
                    }


                    foreach (string item in items)
                        List_Preview.Items.Add(new ColoredItem(item.Trim()));
                    ;


                    // Next


#error Cannot convert ReDimStatementSyntax - see comment for details
                    /* Cannot convert ReDimStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Symbols.Metadata.PE.PENamedTypeSymbolWithEmittedNamespaceName' to type 'Microsoft.CodeAnalysis.IArrayTypeSymbol'.
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.CreateNewArrayAssignment(ExpressionSyntax vbArrayExpression, ExpressionSyntax csArrayExpression, List`1 convertedBounds)
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<ConvertRedimClauseAsync>d__42.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<<VisitReDimStatement>b__41_0>d.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.Common.AsyncEnumerableTaskExtensions.<SelectAsync>d__3`2.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.Common.AsyncEnumerableTaskExtensions.<SelectManyAsync>d__0`2.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<VisitReDimStatement>d__41.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.PerScopeStateVisitorDecorator.<AddLocalVariablesAsync>d__6.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.<DefaultVisitInnerAsync>d__3.MoveNext()

                                        Input:


                                                        'Next

                                                        ReDim Me.mybtn(Me.List_Preview.Items.Count)

                                         */
                }
            }

            catch (Exception ex)
            {

                Show();
                TB_src_rng.Focus();

            }
        }

        private void Label_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Label lbl = (System.Windows.Forms.Label)sender;
            // Reset all labels to their default colors
            foreach (Control control in Controls)
            {
                if (control is System.Windows.Forms.Label)
                {
                    System.Windows.Forms.Label lbl1 = (System.Windows.Forms.Label)control;
                    lbl1.BackColor = Color.White;
                    lbl1.ForeColor = Color.Black;

                }
            }

            lbl.BackColor = SystemColors.Highlight;
            lbl.ForeColor = SystemColors.HighlightText;
            Btn_NC.Focus();

        }



        private void Selection_destination_Click(object sender, EventArgs e)
        {
            if (selectedRange is null)
            {
            }
            else
            {
                // TB_src_range.Text = selectedRange.Address


                Hide();

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;

                // Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


                Excel.Range userInput = (Excel.Range)excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type: 8);
                des_rng = userInput;

                string sheetName;
                sheetName = Strings.Split(des_rng.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                workSheet.Activate();

                des_rng.Select();
                // MsgBox(src_rng.Address)

                TB_des_rng.Text = des_rng.get_Address();

                Show();
                TB_des_rng.Focus();

            }
        }

        private void excelApp_SheetSelectionChange(object Sh, Excel.Range selectionRange1)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                if (focuschange == false)
                {

                    if (ReferenceEquals(ActiveControl, TB_des_rng))
                    {
                        des_rng = selectionRange1;
                        // This will run on the Excel thread, so you need to use Invoke to update the UI
                        // Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                        Activate();
                        BeginInvoke(new System.Action(() =>
                            {
                                TB_des_rng.Text = des_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                    }

                    else if (ReferenceEquals(ActiveControl, TB_src_rng))
                    {
                        src_rng = selectionRange1;
                        Activate();


                        BeginInvoke(new System.Action(() =>
                            {
                                TB_src_rng.Text = src_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                    }

                }
            }

            catch (Exception ex)
            {

            }

        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // Check if Enter is pressed and btn1 is focused
            if (keyData == Keys.Enter && ReferenceEquals(ActiveControl, Btn_NC))
            {
                btn_OK.PerformClick(); // Perform the btn2 click operation
                return true; // The key is handled
            }
            foreach (Control ctrl in FlowLayoutPanel1.Controls)
            {
                if (keyData == Keys.Enter && ctrl is Button)
                {
                    btn_OK.PerformClick(); // Perform the btn2 click operation
                    return true; // The key is handled
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }


        private void btn_OK_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            // Try
            if (des_rng is not null)
            {
                des_rng.FormatConditions.Delete();
            }
            if (RB_Row.Checked == false)
            {
                des_rng = null;
            }


            if (string.IsNullOrEmpty(TB_src_rng.Text) & RB_cell.Checked == true)
            {
                MessageBox.Show("Select all necessary options", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (!string.IsNullOrEmpty(TB_src_rng.Text) & IsValidExcelCellReference(TB_src_rng.Text) == false)
            {
                MessageBox.Show("Select a valid data validation range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (string.IsNullOrEmpty(TB_src_rng.Text) & RB_Row.Checked == true & string.IsNullOrEmpty(TB_des_rng.Text))
            {
                MessageBox.Show("Select all necessary options", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }


            else if (string.IsNullOrEmpty(TB_des_rng.Text) & RB_Row.Checked == true)
            {
                MessageBox.Show("Select the Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();
                // Me.Close()
                return;
            }


            else if (IsValidExcelCellReference(TB_des_rng.Text) == false & RB_Row.Checked == true)
            {

                MessageBox.Show("Select a valid range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();
                // Me.Close()
                return;
            }

            else if (src_rng.Areas.Count > 1)
            {
                MessageBox.Show("Multiple selection is not possible in the Data Validation Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                return;
            }




            // ElseIf RB_Row.Checked = True And src_rng.row <> des_rng.row Then
            // MsgBox("so")
            // ElseIf RB_Row.Checked = True And (src_rng.Row >= des_rng.Row) AndAlso
            // ((src_rng.Row + src_rng.Rows.Count - 1) <= (des_rng.Row + des_rng.Rows.Count - 1)) AndAlso
            // (src_rng.Column >= des_rng.Column) AndAlso
            // ((src_rng.Column + src_rng.Columns.Count - 1) <= (des_rng.Column + des_rng.Columns.Count - 1)) Then



            else if (RB_Row.Checked == true)
            {
                if ((workSheet3.Name ?? "") != (des_rng.Worksheet.Name ?? ""))
                {
                    MessageBox.Show("Please select the range of the same worksheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TB_des_rng.Focus();
                    return;
                }

                else if (src_rng.Row + src_rng.Rows.Count - 1 < des_rng.Row + des_rng.Rows.Count - 1 | src_rng.Row != des_rng.Row | excelApp.Intersect(src_rng, des_rng) is null)
                {


                    MessageBox.Show("Please select the range of the same data table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TB_des_rng.Focus();
                    // Me.Close()
                    return;
                }

                else if (RB_Row.Checked == true && des_rng.Areas.Count > 1)
                {
                    MessageBox.Show("Select Case a valid range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TB_src_rng.Focus();
                    return;
                }
                else
                {
                    goto GotoExpression;
                }
            }

            else
            {

GotoExpression:
                ;


                if (Backup_sheet.Checked == true)
                {
                    worksheet.Copy(After: workbook.Sheets[worksheet.Name]);
                }

                workbook.Sheets[worksheet.Name].Activate();

                // Retrieve data validation items

                // Clear any existing conditional formats

                string formula = src_rng.Validation.Formula1;
                string[] items = formula.Split(',');

                src_rng.FormatConditions.Delete();

                if (RB_cell.Checked == true)
                {
                    foreach (ColoredItem item in List_Preview.Items)
                        // Only drop-down cell

                        AddColorCondition(src_rng, item.ToString(), item.Color);
                }

                else if (RB_Row.Checked == true)
                {
                    foreach (ColoredItem item in List_Preview.Items)
                    {
                        // Color the destination range
                        int i = 0;
                        foreach (var cell in src_rng)
                        {
                            i = i + 1;
                            if (item.Color != Color.White)
                            {
                                AddColorCondition2((Excel.Range)des_rng.Rows[i], (Excel.Range)cell, item.ToString(), item.Color);
                            }
                        }
                    }
                }


            }
            src_rng.Select();


            foreach (var cell in src_rng)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(cell.Validation.Type, Excel.XlDVType.xlValidateList, false)))
                {
                    flag = true;
                    break;
                }

            }


            foreach (var item in List_Preview.Items)
            {
                if (Operators.ConditionalCompareObjectEqual(item.color, Color.White, false) | Operators.ConditionalCompareObjectEqual(item.color, SystemColors.Highlight, false))
                {
                    flag2 = false;
                }
                else
                {
                    flag2 = true;
                    break;
                }
            }


            if (flag == false & GlobalModule.sessionflag2 == true)
            {
                form = new Form42();
                form.Show();
                Hide();
            }

            else if (flag2 == false & GlobalModule.sessionflag1 == true)
            {
                form2 = new Form43();
                form2.Show();
                Hide();
            }


            else
            {

                Close();

            }

            // Me.Show()

            // Catch ex As Exception
            // If flag = False Then
            // Me.Hide()
            // form = New Form42
            // form.Show()
            // If form.IsDisposed Or form Is Nothing Then
            // Me.Show()
            // End If

            // ElseIf flag2 = False Then
            // Me.Hide()
            // form2 = New Form43
            // form2.Show()
            // If form2.IsDisposed Or form2 Is Nothing Then
            // Me.Show()
            // End If
            // End If
            Close();
            // End Try
        }
        private Color GetColorForItem(int index)
        {
            // This function maps an index to a color
            // You can adjust or expand this as needed
            Color[] colors = new[] { Color.Red, Color.Green, Color.Blue, Color.Yellow, Color.Purple };

            if (index >= 0 & index < colors.Length)
            {
                return colors[index];
            }
            else
            {
                return Color.White;
            }
        }

        private void AddColorCondition(Excel.Range targetRange, string value, Color color)
        {
            Excel.FormatCondition condition = (Excel.FormatCondition)targetRange.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlCellValue, Operator: Excel.XlFormatConditionOperator.xlEqual, Formula1: value);
            condition.Interior.Color = ColorTranslator.ToOle(color);
        }

        private void AddColorCondition2(Excel.Range targetRange, Excel.Range controlCell, string value, Color color)
        {
            if (Information.IsNumeric(value) == false)
            {
                string formula = "=" + controlCell.get_Address() + " = \"" + value + "\"";
                Excel.FormatCondition condition = (Excel.FormatCondition)targetRange.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Formula1: formula);
                condition.Interior.Color = ColorTranslator.ToOle(color);
            }
            else
            {
                string formula = "=" + controlCell.get_Address() + " = " + value + "";
                Excel.FormatCondition condition = (Excel.FormatCondition)targetRange.FormatConditions.Add(Type: Excel.XlFormatConditionType.xlExpression, Formula1: formula);
                condition.Interior.Color = ColorTranslator.ToOle(color);
            }

        }


        private void RB_Row_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Row.Checked == true)
            {
                Selection_destination.Enabled = true;
                TB_des_rng.Enabled = true;
                TB_des_rng.Focus();
            }
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void TB_src_rng_TextChanged(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                var workbook = excelApp.ActiveWorkbook;
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                TB_src_rng.Focus();


                if (TB_src_rng.Text is not null & IsValidExcelCellReference(TB_src_rng.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    src_rng = excelApp.get_Range(TB_src_rng.Text);
                    src_rng.Select();
                    var range = src_rng;

                    Activate();
                    // TB_src_range.Focus()
                    TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                    focuschange = false;



                    Excel.Range ran = (Excel.Range)src_rng[1, 1];

                    // Clear the ListBox
                    List_Preview.Items.Clear();
                    // MsgBox(ran.Address)

                    // If range.Validation.Type = Excel.XlDVType.xlValidateList Then
                    string formula = ran.Validation.Formula1;
                    // MsgBox(formula)
                    var items = new List<string>();
                    if (formula.Contains(":"))
                    {
                        range = excelApp.get_Range(formula);
                        foreach (var r in range)
                            items.Add(r.Value.ToString());
                    }
                    else
                    {
                        // Else, split the formula to get the individual items
                        items.AddRange(formula.Split(','));
                    }


                    foreach (string item in items)
                        // MsgBox(item.ToString)
                        List_Preview.Items.Add(new ColoredItem(item.Trim()));

                    TB_src_rng.Focus();
                    ;
#error Cannot convert ReDimStatementSyntax - see comment for details
                    /* Cannot convert ReDimStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Symbols.Metadata.PE.PENamedTypeSymbolWithEmittedNamespaceName' to type 'Microsoft.CodeAnalysis.IArrayTypeSymbol'.
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.CreateNewArrayAssignment(ExpressionSyntax vbArrayExpression, ExpressionSyntax csArrayExpression, List`1 convertedBounds)
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<ConvertRedimClauseAsync>d__42.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<<VisitReDimStatement>b__41_0>d.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.Common.AsyncEnumerableTaskExtensions.<SelectAsync>d__3`2.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.Common.AsyncEnumerableTaskExtensions.<SelectManyAsync>d__0`2.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.MethodBodyExecutableStatementVisitor.<VisitReDimStatement>d__41.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.PerScopeStateVisitorDecorator.<AddLocalVariablesAsync>d__6.MoveNext()
                                        --- End of stack trace from previous location where exception was thrown ---
                                           at System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw()
                                           at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.<DefaultVisitInnerAsync>d__3.MoveNext()

                                        Input:
                                                        ReDim Me.mybtn(Me.List_Preview.Items.Count)

                                         */
                    workSheet3 = worksheet;

                }
            }

            catch (Exception ex)
            {
                TB_src_rng.Focus();
            }
        }




        // Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub RB_cell_KeyDown(sender As Object, e As KeyEventArgs) Handles RB_cell.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub


        // Private Sub RB_row_KeyDown(sender As Object, e As KeyEventArgs) Handles Btn_NC.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // btn_OK.PerformClick()

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub
        // Private Sub Sample_Image_KeyDown(sender As Object, e As KeyEventArgs) Handles Btn_NC.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception
        // MsgBox(1)
        // End Try

        // End Sub


        // Private Sub TextBox_Destination_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_des_rng.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub TextBox_Source_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        private void Button1_Click(object sender, EventArgs e)
        {
            int index = 0;
            foreach (var item in List_Preview.Items)
            {
                item = (ColoredItem)List_Preview.Items[index];
                item.Color = (object)Color.White;
                index = index + 1;
            }
            Refresh();
        }

        private void Btn_color_Click(object sender, EventArgs e)
        {
            Button clickedButton = (Button)sender;
            objectPosition = clickedButton.Location;
            Color c;
            // MsgBox(8)
            c = Btn_color.BackColor;


            int index = List_Preview.SelectedIndex;
            if (index < 0)
                return;

            ColoredItem item = (ColoredItem)List_Preview.Items[index];
            item.Color = c;
            clickedButton.Focus();

            this.mybtn((object)List_Preview.SelectedIndex) = clickedButton;
            Btn_color.BackColor = c;
            Refresh();
        }

        private bool IsValidExcelCellReference(string cellReference)
        {

            // Regular expression pattern for a valid sheet name. This is a simplified version and might not cover all edge cases.
            // Excel sheet names cannot contain the characters \, /, *, [, ], :, ?, and cannot be 'History'.
            string sheetNamePattern = @"(?i)(?![\/*[\]:?])(?!History)[^\/\[\]*?:\\]+";

            // Regular expression pattern for a cell reference.
            // This pattern will match references like A1, $A$1, etc.
            string cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";

            // Regular expression pattern for an Excel reference.
            // This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
            string singleReferencePattern = cellPattern + "(:" + cellPattern + ")?";

            // Regular expression pattern to allow the sheet name, followed by '!', before the cell reference
            string fullPattern = "^(" + sheetNamePattern + "!)?(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$";

            // Create a regex object with the pattern.
            var regex = new Regex(fullPattern);

            // Test the input string against the regex pattern.
            return regex.IsMatch(cellReference.ToUpper());

        }

        // Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_des_rng.KeyDown
        // 'If Enter key is pressed then check if the text is a valid address
        // If IsValidExcelCellReference(TB_des_rng.Text) = True And e.KeyCode = Keys.Enter Then
        // des_rng = excelApp.Range(TB_des_rng.Text)
        // TB_des_rng.Focus()
        // des_rng.Select()

        // Call btn_OK_Click(sender, e)   'OK button click event called

        // 'MsgBox(des_rng.Address)
        // ElseIf IsValidExcelCellReference(TB_des_rng.Text) = False And e.KeyCode = Keys.Enter Then
        // MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        // TB_des_rng.Text = ""
        // TB_des_rng.Focus()
        // 'Me.Close()
        // Exit Sub
        // End If
        // End Sub

        // Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown
        // 'If Enter key is pressed then check if the text is a valid address

        // If IsValidExcelCellReference(TB_src_rng.Text) = True And e.KeyCode = Keys.Enter Then
        // src_rng = excelApp.Range(TB_src_rng.Text)
        // TB_src_rng.Focus()
        // src_rng.Select()

        // Call btn_OK_Click(sender, e)   'OK button click event called

        // 'MsgBox(des_rng.Address)
        // ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False And e.KeyCode = Keys.Enter Then
        // MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        // TB_src_rng.Text = ""
        // TB_src_rng.Focus()
        // 'Me.Close()
        // Exit Sub
        // End If
        // End Sub
        // Private Sub FlowLayout_KeyDown(sender As Object, e As KeyEventArgs) Handles FlowLayoutPanel1.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub



        // Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter And TB_src_rng.Focus = False And TB_des_rng.Focus = False Then

        // Call btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        private void TB_des_rng_TextChanged(object sender, EventArgs e)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workbook = excelApp.ActiveWorkbook;
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                // workSheet3 = worksheet

                if (TB_des_rng.Text is not null & IsValidExcelCellReference(TB_des_rng.Text) == true)
                {
                    focuschange = true;
                    try
                    {
                        // Define the range of cells to read (for example, cells A1 to A10)
                        des_rng = excelApp.get_Range(TB_des_rng.Text);
                        // Dim range As Excel.Range = des_rng
                        des_rng.Select();
                    }
                    catch
                    {
                        // Split the string into sheet name and cell address
                        string[] parts = TB_des_rng.Text.Split('!');
                        string sheetName = parts[0];
                        string cellAddress = parts[1];

                        des_rng = excelApp.get_Range(cellAddress);
                        des_rng.Select();
                    }

                    if ((worksheet.Name ?? "") != (workSheet3.Name ?? ""))
                    {
                        TB_des_rng.Text = worksheet.Name + "!" + des_rng.get_Address();
                        des_rng = excelApp.get_Range(TB_des_rng.Text);
                    }
                    Activate();
                    // TB_src_range.Focus()
                    TB_des_rng.SelectionStart = TB_des_rng.Text.Length;
                    focuschange = false;
                    ax = worksheet.Name;
                }
            }

            catch (Exception ex)
            {
                ax = "";
            }
        }


        private void Form33_ColorBasedDropDownList_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form33_ColorBasedDropDownList_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form33_ColorBasedDropDownList_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    TB_src_rng.Text = src_rng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
            TB_src_rng.Focus();
        }

        private void RB_cell_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_cell.Checked == true)
            {
                Selection_destination.Enabled = false;
                TB_des_rng.Enabled = false;
            }
        }

        private void Btn_NC_Click(object sender, EventArgs e)
        {
            Button clickedButton = (Button)sender;
            objectPosition = clickedButton.Location;
            int index = List_Preview.SelectedIndex;

            ColoredItem item = (ColoredItem)List_Preview.Items[index];
            item.Color = Color.White;
            Button2.Focus();

            this.mybtn((object)List_Preview.SelectedIndex) = Button2;
            Btn_color.BackColor = Color.White;
            Refresh();
        }

        private void List_Box_IndexChanged(object sender, EventArgs e) => List_Box_IndexChanged();


    }

    public class ColoredItem
    {
        public string Text { get; set; }
        public Color Color { get; set; } = Color.White;

        public ColoredItem(string t)
        {
            Text = t;
        }

        public ColoredItem(string t, Color c)
        {
            Text = t;
            Color = c;
        }

        public override string ToString()
        {
            return Text;
        }
    }
}