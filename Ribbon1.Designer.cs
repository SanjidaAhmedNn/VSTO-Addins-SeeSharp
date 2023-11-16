using System.Runtime.CompilerServices;
using Office = Microsoft.Office.Core;

namespace VSTO_Addins
{
    public partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {

        [System.Diagnostics.DebuggerNonUserCode()]
        public Ribbon1(System.ComponentModel.IContainer container) : this()
        {

            // Required for Windows.Forms Class Composition Designer support
            if (container is not null)
            {
                container.Add(this);
            }

        }

        [System.Diagnostics.DebuggerNonUserCode()]
        public Ribbon1() : base(Globals.Factory.GetRibbonFactory())
        {

            // This call is required by the Component Designer.
            InitializeComponent();
            Load += Ribbon1_Load;

        }

        // Component overrides dispose to clean up the component list.
        [System.Diagnostics.DebuggerNonUserCode()]
        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing && components is not null)
                {
                    components.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        // Required by the Component Designer
        private System.ComponentModel.IContainer components;

        // NOTE: The following procedure is required by the Component Designer
        // It can be modified using the Component Designer.
        // Do not modify it using the code editor.
        [System.Diagnostics.DebuggerStepThrough()]
        private void InitializeComponent()
        {
            _Tab1 = Factory.CreateRibbonTab();
            _Group1 = Factory.CreateRibbonGroup();
            _Separator1 = Factory.CreateRibbonSeparator();
            _Group2 = Factory.CreateRibbonGroup();
            _Group5 = Factory.CreateRibbonGroup();
            _Group3 = Factory.CreateRibbonGroup();
            _Separator2 = Factory.CreateRibbonSeparator();
            _Group4 = Factory.CreateRibbonGroup();
            _Button1 = Factory.CreateRibbonButton();
            _Button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button1_Click);
            _Button8 = Factory.CreateRibbonButton();
            _Button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button8_Click);
            _Button3 = Factory.CreateRibbonButton();
            _Button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button3_Click);
            _Button7 = Factory.CreateRibbonButton();
            _Button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button7_Click);
            _Button13 = Factory.CreateRibbonButton();
            _Button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button13_Click);
            _Menu2 = Factory.CreateRibbonMenu();
            _Button12 = Factory.CreateRibbonButton();
            _Button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button12_Click_1);
            _Button10 = Factory.CreateRibbonButton();
            _Button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button10_Click_1);
            _Button14 = Factory.CreateRibbonButton();
            _Button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button14_Click);
            _Button5 = Factory.CreateRibbonButton();
            _Button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button5_Click);
            _Menu5 = Factory.CreateRibbonMenu();
            _Button16 = Factory.CreateRibbonButton();
            _Button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button16_Click);
            _Button17 = Factory.CreateRibbonButton();
            _Button17.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button17_Click);
            _Button18 = Factory.CreateRibbonButton();
            _Button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button18_Click);
            _Menu8 = Factory.CreateRibbonMenu();
            _Button20 = Factory.CreateRibbonButton();
            _Button20.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button20_Click_1);
            _Button21 = Factory.CreateRibbonButton();
            _Button21.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button21_Click_1);
            _Button6 = Factory.CreateRibbonButton();
            _Button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button6_Click);
            _Menu10 = Factory.CreateRibbonMenu();
            _Button23 = Factory.CreateRibbonButton();
            _Button23.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button23_Click_1);
            _Button22 = Factory.CreateRibbonButton();
            _Button22.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button22_Click_1);
            _Menu9 = Factory.CreateRibbonMenu();
            _Button45 = Factory.CreateRibbonButton();
            _Button45.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button45_Click);
            _Button46 = Factory.CreateRibbonButton();
            _Button46.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button46_Click);
            _Button47 = Factory.CreateRibbonButton();
            _Button47.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button47_Click);
            _Button15 = Factory.CreateRibbonButton();
            _Button15.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button15_Click);
            _Menu11 = Factory.CreateRibbonMenu();
            _Button54 = Factory.CreateRibbonButton();
            _Button54.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button54_Click);
            _Button11 = Factory.CreateRibbonButton();
            _Button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button11_Click_1);
            _Menu4 = Factory.CreateRibbonMenu();
            _Button31 = Factory.CreateRibbonButton();
            _Button31.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button31_Click);
            _Button32 = Factory.CreateRibbonButton();
            _Button32.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button32_Click);
            _Menu7 = Factory.CreateRibbonMenu();
            _Button37 = Factory.CreateRibbonButton();
            _Button37.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button37_Click);
            _Button38 = Factory.CreateRibbonButton();
            _Button38.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button38_Click);
            _Button39 = Factory.CreateRibbonButton();
            _Button39.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button39_Click);
            _Button40 = Factory.CreateRibbonButton();
            _Button40.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button40_Click);
            _Menu6 = Factory.CreateRibbonMenu();
            _Button33 = Factory.CreateRibbonButton();
            _Button33.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button33_Click);
            _Button34 = Factory.CreateRibbonButton();
            _Button34.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button34_Click);
            _Button35 = Factory.CreateRibbonButton();
            _Button35.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button35_Click);
            _Button36 = Factory.CreateRibbonButton();
            _Button36.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button36_Click);
            _Button41 = Factory.CreateRibbonButton();
            _Button41.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button41_Click);
            _Button19 = Factory.CreateRibbonButton();
            _Button19.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button19_Click);
            _Menu1 = Factory.CreateRibbonMenu();
            _Button2 = Factory.CreateRibbonButton();
            _Button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button2_Click_2);
            _Button9 = Factory.CreateRibbonButton();
            _Button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button9_Click_1);
            _Button49 = Factory.CreateRibbonButton();
            _Button49.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button49_Click);
            _Menu3 = Factory.CreateRibbonMenu();
            _Button28 = Factory.CreateRibbonButton();
            _Button28.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button28_Click);
            _Button29 = Factory.CreateRibbonButton();
            _Button29.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button29_Click);
            _Button30 = Factory.CreateRibbonButton();
            _Button30.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button30_Click);
            _Button4 = Factory.CreateRibbonButton();
            _SplitButton7 = Factory.CreateRibbonSplitButton();
            _Button24 = Factory.CreateRibbonButton();
            _Button24.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button24_Click_1);
            _Button25 = Factory.CreateRibbonButton();
            _Button25.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button25_Click_1);
            _Button26 = Factory.CreateRibbonButton();
            _Button26.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button26_Click_1);
            _Button27 = Factory.CreateRibbonButton();
            _Button27.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(Button27_Click_1);
            _Tab1.SuspendLayout();
            _Group1.SuspendLayout();
            _Group2.SuspendLayout();
            _Group5.SuspendLayout();
            _Group3.SuspendLayout();
            _Group4.SuspendLayout();
            SuspendLayout();
            // 
            // Tab1
            // 
            _Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            _Tab1.Groups.Add(_Group1);
            _Tab1.Groups.Add(_Group2);
            _Tab1.Groups.Add(_Group5);
            _Tab1.Groups.Add(_Group3);
            _Tab1.Groups.Add(_Group4);
            _Tab1.Label = "TabAddIns";
            _Tab1.Name = "_Tab1";
            // 
            // Group1
            // 
            _Group1.Items.Add(_Button1);
            _Group1.Items.Add(_Button8);
            _Group1.Items.Add(_Button3);
            _Group1.Items.Add(_Separator1);
            _Group1.Items.Add(_Button7);
            _Group1.Items.Add(_Button13);
            _Group1.Items.Add(_Menu2);
            _Group1.Items.Add(_Button14);
            _Group1.Label = "Range";
            _Group1.Name = "_Group1";
            // 
            // Separator1
            // 
            _Separator1.Name = "_Separator1";
            // 
            // Group2
            // 
            _Group2.Items.Add(_Button5);
            _Group2.Items.Add(_Menu5);
            _Group2.Items.Add(_Menu8);
            _Group2.Items.Add(_Button6);
            _Group2.Items.Add(_Menu10);
            _Group2.Items.Add(_Menu9);
            _Group2.Items.Add(_Button15);
            _Group2.Label = "Merge && Unmerge";
            _Group2.Name = "_Group2";
            // 
            // Group5
            // 
            _Group5.Items.Add(_Menu11);
            _Group5.Items.Add(_Menu4);
            _Group5.Label = "Hide && Unhide";
            _Group5.Name = "_Group5";
            // 
            // Group3
            // 
            _Group3.Items.Add(_Menu7);
            _Group3.Items.Add(_Menu6);
            _Group3.Items.Add(_Button41);
            _Group3.Items.Add(_Separator2);
            _Group3.Items.Add(_Button19);
            _Group3.Label = "Remove Blanks";
            _Group3.Name = "_Group3";
            // 
            // Separator2
            // 
            _Separator2.Name = "_Separator2";
            // 
            // Group4
            // 
            _Group4.Items.Add(_Menu1);
            _Group4.Items.Add(_Button49);
            _Group4.Items.Add(_Menu3);
            _Group4.Items.Add(_SplitButton7);
            _Group4.Label = "Drop-down List";
            _Group4.Name = "_Group4";
            // 
            // Button1
            // 
            _Button1.Label = "    Flip";
            _Button1.Name = "_Button1";
            // 
            // Button8
            // 
            _Button8.Label = "  Swap";
            _Button8.Name = "_Button8";
            // 
            // Button3
            // 
            _Button3.Label = "Transpose";
            _Button3.Name = "_Button3";
            // 
            // Button7
            // 
            _Button7.Label = "Transform";
            _Button7.Name = "_Button7";
            // 
            // Button13
            // 
            _Button13.Label = "Compare Cells";
            _Button13.Name = "_Button13";
            // 
            // Menu2
            // 
            _Menu2.Items.Add(_Button12);
            _Menu2.Items.Add(_Button10);
            _Menu2.Label = "Specify Scroll Area";
            _Menu2.Name = "_Menu2";
            // 
            // Button12
            // 
            _Button12.Label = "Set";
            _Button12.Name = "_Button12";
            _Button12.ShowImage = true;
            // 
            // Button10
            // 
            _Button10.Label = "Clear";
            _Button10.Name = "_Button10";
            _Button10.ShowImage = true;
            // 
            // Button14
            // 
            _Button14.ControlSize = Office.RibbonControlSize.RibbonControlSizeLarge;
            _Button14.Label = "Paste into Visible Range";
            _Button14.Name = "_Button14";
            _Button14.ShowImage = true;
            // 
            // Button5
            // 
            _Button5.ControlSize = Office.RibbonControlSize.RibbonControlSizeLarge;
            _Button5.Label = "Merge Cells with Same Value";
            _Button5.Name = "_Button5";
            _Button5.ShowImage = true;
            // 
            // Menu5
            // 
            _Menu5.Items.Add(_Button16);
            _Menu5.Items.Add(_Button17);
            _Menu5.Items.Add(_Button18);
            _Menu5.Label = "Combine Range";
            _Menu5.Name = "_Menu5";
            // 
            // Button16
            // 
            _Button16.Label = "Combine Ranges into Column";
            _Button16.Name = "_Button16";
            _Button16.ShowImage = true;
            // 
            // Button17
            // 
            _Button17.Label = "Combine Ranges into Row";
            _Button17.Name = "_Button17";
            _Button17.ShowImage = true;
            // 
            // Button18
            // 
            _Button18.Label = "Combine Ranges into Cells";
            _Button18.Name = "_Button18";
            _Button18.ShowImage = true;
            // 
            // Menu8
            // 
            _Menu8.Items.Add(_Button20);
            _Menu8.Items.Add(_Button21);
            _Menu8.Label = "Combine Duplicate";
            _Menu8.Name = "_Menu8";
            // 
            // Button20
            // 
            _Button20.Label = "Combine Duplicate Rows";
            _Button20.Name = "_Button20";
            _Button20.ShowImage = true;
            // 
            // Button21
            // 
            _Button21.Label = "Combine Duplicate Columns";
            _Button21.Name = "_Button21";
            _Button21.ShowImage = true;
            // 
            // Button6
            // 
            _Button6.Label = "Unmerge Cells with Value";
            _Button6.Name = "_Button6";
            // 
            // Menu10
            // 
            _Menu10.Items.Add(_Button23);
            _Menu10.Items.Add(_Button22);
            _Menu10.Label = "Split Data";
            _Menu10.Name = "_Menu10";
            // 
            // Button23
            // 
            _Button23.Label = "Split Range";
            _Button23.Name = "_Button23";
            _Button23.ShowImage = true;
            // 
            // Button22
            // 
            _Button22.Label = "Split Cells";
            _Button22.Name = "_Button22";
            _Button22.ShowImage = true;
            // 
            // Menu9
            // 
            _Menu9.Items.Add(_Button45);
            _Menu9.Items.Add(_Button46);
            _Menu9.Items.Add(_Button47);
            _Menu9.Label = "Split Text";
            _Menu9.Name = "_Menu9";
            // 
            // Button45
            // 
            _Button45.Label = "Split Text by Characters";
            _Button45.Name = "_Button45";
            _Button45.ShowImage = true;
            // 
            // Button46
            // 
            _Button46.Label = "Split Text by Strings";
            _Button46.Name = "_Button46";
            _Button46.ShowImage = true;
            // 
            // Button47
            // 
            _Button47.Label = "Split Text by Patterns";
            _Button47.Name = "_Button47";
            _Button47.ShowImage = true;
            // 
            // Button15
            // 
            _Button15.Label = "Divide Names";
            _Button15.Name = "_Button15";
            // 
            // Menu11
            // 
            _Menu11.Items.Add(_Button54);
            _Menu11.Items.Add(_Button11);
            _Menu11.Label = "Hide Ranges";
            _Menu11.Name = "_Menu11";
            // 
            // Button54
            // 
            _Button54.Label = "Hide only the selected range";
            _Button54.Name = "_Button54";
            _Button54.ShowImage = true;
            // 
            // Button11
            // 
            _Button11.Label = "Hide all except the selected range";
            _Button11.Name = "_Button11";
            _Button11.ShowImage = true;
            // 
            // Menu4
            // 
            _Menu4.Items.Add(_Button31);
            _Menu4.Items.Add(_Button32);
            _Menu4.Label = "Unhide Ranges";
            _Menu4.Name = "_Menu4";
            // 
            // Button31
            // 
            _Button31.Label = "Unhide All Ranges";
            _Button31.Name = "_Button31";
            _Button31.ShowImage = true;
            // 
            // Button32
            // 
            _Button32.Label = "Unhide Ranges from the Selection";
            _Button32.Name = "_Button32";
            _Button32.ShowImage = true;
            // 
            // Menu7
            // 
            _Menu7.Items.Add(_Button37);
            _Menu7.Items.Add(_Button38);
            _Menu7.Items.Add(_Button39);
            _Menu7.Items.Add(_Button40);
            _Menu7.Label = "Empty Rows";
            _Menu7.Name = "_Menu7";
            _Menu7.ShowImage = true;
            // 
            // Button37
            // 
            _Button37.Label = "From Selected Range";
            _Button37.Name = "_Button37";
            _Button37.ShowImage = true;
            // 
            // Button38
            // 
            _Button38.Label = "From Active Sheet";
            _Button38.Name = "_Button38";
            _Button38.ShowImage = true;
            // 
            // Button39
            // 
            _Button39.Label = "From Selected Sheets";
            _Button39.Name = "_Button39";
            _Button39.ShowImage = true;
            // 
            // Button40
            // 
            _Button40.Label = "From All Sheets";
            _Button40.Name = "_Button40";
            _Button40.ShowImage = true;
            // 
            // Menu6
            // 
            _Menu6.Items.Add(_Button33);
            _Menu6.Items.Add(_Button34);
            _Menu6.Items.Add(_Button35);
            _Menu6.Items.Add(_Button36);
            _Menu6.Label = "Empty Columns";
            _Menu6.Name = "_Menu6";
            _Menu6.ShowImage = true;
            // 
            // Button33
            // 
            _Button33.Label = "From Selected Range";
            _Button33.Name = "_Button33";
            _Button33.ShowImage = true;
            // 
            // Button34
            // 
            _Button34.Label = "From Active Sheet";
            _Button34.Name = "_Button34";
            _Button34.ShowImage = true;
            // 
            // Button35
            // 
            _Button35.Label = "From Selected Sheets";
            _Button35.Name = "_Button35";
            _Button35.ShowImage = true;
            // 
            // Button36
            // 
            _Button36.Label = "From All Sheets";
            _Button36.Name = "_Button36";
            _Button36.ShowImage = true;
            // 
            // Button41
            // 
            _Button41.Label = "Empty Sheets";
            _Button41.Name = "_Button41";
            _Button41.ShowImage = true;
            // 
            // Button19
            // 
            _Button19.ControlSize = Office.RibbonControlSize.RibbonControlSizeLarge;
            _Button19.Label = "Fill Empty Cells";
            _Button19.Name = "_Button19";
            _Button19.ShowImage = true;
            // 
            // Menu1
            // 
            _Menu1.Items.Add(_Button2);
            _Menu1.Items.Add(_Button9);
            _Menu1.Label = "Create Drop-down List";
            _Menu1.Name = "_Menu1";
            // 
            // Button2
            // 
            _Button2.Label = "Simple Drop-down List";
            _Button2.Name = "_Button2";
            _Button2.ShowImage = true;
            // 
            // Button9
            // 
            _Button9.Label = "Picture Based Drop-down List";
            _Button9.Name = "_Button9";
            _Button9.ShowImage = true;
            // 
            // Button49
            // 
            _Button49.Label = "Color Based Drop-down List";
            _Button49.Name = "_Button49";
            // 
            // Menu3
            // 
            _Menu3.Items.Add(_Button28);
            _Menu3.Items.Add(_Button29);
            _Menu3.Items.Add(_Button30);
            _Menu3.Items.Add(_Button4);
            _Menu3.Label = "Dynamic Drop-down List";
            _Menu3.Name = "_Menu3";
            // 
            // Button28
            // 
            _Button28.Label = "Create";
            _Button28.Name = "_Button28";
            _Button28.ShowImage = true;
            // 
            // Button29
            // 
            _Button29.Label = "Update";
            _Button29.Name = "_Button29";
            _Button29.ShowImage = true;
            // 
            // Button30
            // 
            _Button30.Label = "Extend";
            _Button30.Name = "_Button30";
            _Button30.ShowImage = true;
            // 
            // Button4
            // 
            _Button4.Label = "Button4";
            _Button4.Name = "_Button4";
            _Button4.ShowImage = true;
            // 
            // SplitButton7
            // 
            _SplitButton7.ControlSize = Office.RibbonControlSize.RibbonControlSizeLarge;
            _SplitButton7.Items.Add(_Button24);
            _SplitButton7.Items.Add(_Button25);
            _SplitButton7.Items.Add(_Button26);
            _SplitButton7.Items.Add(_Button27);
            _SplitButton7.Label = "Advanced Drop-down List";
            _SplitButton7.Name = "_SplitButton7";
            // 
            // Button24
            // 
            _Button24.Label = "Multi-selection Based Drop-down List";
            _Button24.Name = "_Button24";
            _Button24.ShowImage = true;
            // 
            // Button25
            // 
            _Button25.Label = "Drop-down List with Checkbox";
            _Button25.Name = "_Button25";
            _Button25.ShowImage = true;
            // 
            // Button26
            // 
            _Button26.Label = "Drop-down List with Search Option";
            _Button26.Name = "_Button26";
            _Button26.ShowImage = true;
            // 
            // Button27
            // 
            _Button27.Label = "Remove Advanced Drop-down List";
            _Button27.Name = "_Button27";
            _Button27.ShowImage = true;
            // 
            // Ribbon1
            // 
            Name = "Ribbon1";
            RibbonType = "Microsoft.Excel.Workbook";
            Tabs.Add(_Tab1);
            _Tab1.ResumeLayout(false);
            _Tab1.PerformLayout();
            _Group1.ResumeLayout(false);
            _Group1.PerformLayout();
            _Group2.ResumeLayout(false);
            _Group2.PerformLayout();
            _Group5.ResumeLayout(false);
            _Group5.PerformLayout();
            _Group3.ResumeLayout(false);
            _Group3.PerformLayout();
            _Group4.ResumeLayout(false);
            _Group4.PerformLayout();
            ResumeLayout(false);

        }

        private Microsoft.Office.Tools.Ribbon.RibbonTab _Tab1;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonTab Tab1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Tab1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Tab1 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonGroup _Group1;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonGroup Group1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Group1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Group1 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button1;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button1 != null)
                {
                    _Button1.Click -= Button1_Click;
                }

                _Button1 = value;
                if (_Button1 != null)
                {
                    _Button1.Click += Button1_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button3;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button3 != null)
                {
                    _Button3.Click -= Button3_Click;
                }

                _Button3 = value;
                if (_Button3 != null)
                {
                    _Button3.Click += Button3_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button5;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button5 != null)
                {
                    _Button5.Click -= Button5_Click;
                }

                _Button5 = value;
                if (_Button5 != null)
                {
                    _Button5.Click += Button5_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button6;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button6
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button6;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button6 != null)
                {
                    _Button6.Click -= Button6_Click;
                }

                _Button6 = value;
                if (_Button6 != null)
                {
                    _Button6.Click += Button6_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button7;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button7
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button7;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button7 != null)
                {
                    _Button7.Click -= Button7_Click;
                }

                _Button7 = value;
                if (_Button7 != null)
                {
                    _Button7.Click += Button7_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonGroup _Group2;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonGroup Group2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Group2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Group2 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button8;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button8
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button8;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button8 != null)
                {
                    _Button8.Click -= Button8_Click;
                }

                _Button8 = value;
                if (_Button8 != null)
                {
                    _Button8.Click += Button8_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonGroup _Group3;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonGroup Group3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Group3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Group3 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button13;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button13
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button13;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button13 != null)
                {
                    _Button13.Click -= Button13_Click;
                }

                _Button13 = value;
                if (_Button13 != null)
                {
                    _Button13.Click += Button13_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button14;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button14
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button14;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button14 != null)
                {
                    _Button14.Click -= Button14_Click;
                }

                _Button14 = value;
                if (_Button14 != null)
                {
                    _Button14.Click += Button14_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button15;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button15
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button15;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button15 != null)
                {
                    _Button15.Click -= Button15_Click;
                }

                _Button15 = value;
                if (_Button15 != null)
                {
                    _Button15.Click += Button15_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button19;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button19
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button19;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button19 != null)
                {
                    _Button19.Click -= Button19_Click;
                }

                _Button19 = value;
                if (_Button19 != null)
                {
                    _Button19.Click += Button19_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonGroup _Group4;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonGroup Group4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Group4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Group4 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu3;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu3
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu3;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu3 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button28;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button28
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button28;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button28 != null)
                {
                    _Button28.Click -= Button28_Click;
                }

                _Button28 = value;
                if (_Button28 != null)
                {
                    _Button28.Click += Button28_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button29;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button29
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button29;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button29 != null)
                {
                    _Button29.Click -= Button29_Click;
                }

                _Button29 = value;
                if (_Button29 != null)
                {
                    _Button29.Click += Button29_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button30;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button30
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button30;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button30 != null)
                {
                    _Button30.Click -= Button30_Click;
                }

                _Button30 = value;
                if (_Button30 != null)
                {
                    _Button30.Click += Button30_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu4;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu4 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button31;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button31
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button31;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button31 != null)
                {
                    _Button31.Click -= Button31_Click;
                }

                _Button31 = value;
                if (_Button31 != null)
                {
                    _Button31.Click += Button31_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button32;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button32
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button32;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button32 != null)
                {
                    _Button32.Click -= Button32_Click;
                }

                _Button32 = value;
                if (_Button32 != null)
                {
                    _Button32.Click += Button32_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu6;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu6
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu6;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu6 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button33;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button33
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button33;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button33 != null)
                {
                    _Button33.Click -= Button33_Click;
                }

                _Button33 = value;
                if (_Button33 != null)
                {
                    _Button33.Click += Button33_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button34;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button34
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button34;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button34 != null)
                {
                    _Button34.Click -= Button34_Click;
                }

                _Button34 = value;
                if (_Button34 != null)
                {
                    _Button34.Click += Button34_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button35;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button35
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button35;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button35 != null)
                {
                    _Button35.Click -= Button35_Click;
                }

                _Button35 = value;
                if (_Button35 != null)
                {
                    _Button35.Click += Button35_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button36;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button36
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button36;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button36 != null)
                {
                    _Button36.Click -= Button36_Click;
                }

                _Button36 = value;
                if (_Button36 != null)
                {
                    _Button36.Click += Button36_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu7;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu7
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu7;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu7 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button41;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button41
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button41;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button41 != null)
                {
                    _Button41.Click -= Button41_Click;
                }

                _Button41 = value;
                if (_Button41 != null)
                {
                    _Button41.Click += Button41_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button37;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button37
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button37;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button37 != null)
                {
                    _Button37.Click -= Button37_Click;
                }

                _Button37 = value;
                if (_Button37 != null)
                {
                    _Button37.Click += Button37_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button38;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button38
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button38;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button38 != null)
                {
                    _Button38.Click -= Button38_Click;
                }

                _Button38 = value;
                if (_Button38 != null)
                {
                    _Button38.Click += Button38_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button39;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button39
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button39;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button39 != null)
                {
                    _Button39.Click -= Button39_Click;
                }

                _Button39 = value;
                if (_Button39 != null)
                {
                    _Button39.Click += Button39_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button40;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button40
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button40;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button40 != null)
                {
                    _Button40.Click -= Button40_Click;
                }

                _Button40 = value;
                if (_Button40 != null)
                {
                    _Button40.Click += Button40_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonGroup _Group5;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonGroup Group5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Group5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Group5 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button49;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button49
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button49;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button49 != null)
                {
                    _Button49.Click -= Button49_Click;
                }

                _Button49 = value;
                if (_Button49 != null)
                {
                    _Button49.Click += Button49_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonSplitButton _SplitButton7;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonSplitButton SplitButton7
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _SplitButton7;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _SplitButton7 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonSeparator _Separator1;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonSeparator Separator1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Separator1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Separator1 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonSeparator _Separator2;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonSeparator Separator2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Separator2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Separator2 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu1;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu1
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu1;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu1 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button2;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button2 != null)
                {
                    _Button2.Click -= Button2_Click_2;
                }

                _Button2 = value;
                if (_Button2 != null)
                {
                    _Button2.Click += Button2_Click_2;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button9;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button9
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button9;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button9 != null)
                {
                    _Button9.Click -= Button9_Click_1;
                }

                _Button9 = value;
                if (_Button9 != null)
                {
                    _Button9.Click += Button9_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu2;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu2
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu2;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu2 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button12;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button12
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button12;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button12 != null)
                {
                    _Button12.Click -= Button12_Click_1;
                }

                _Button12 = value;
                if (_Button12 != null)
                {
                    _Button12.Click += Button12_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button10;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button10
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button10;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button10 != null)
                {
                    _Button10.Click -= Button10_Click_1;
                }

                _Button10 = value;
                if (_Button10 != null)
                {
                    _Button10.Click += Button10_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu5;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu5
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu5;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu5 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button16;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button16
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button16;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button16 != null)
                {
                    _Button16.Click -= Button16_Click;
                }

                _Button16 = value;
                if (_Button16 != null)
                {
                    _Button16.Click += Button16_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button17;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button17
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button17;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button17 != null)
                {
                    _Button17.Click -= Button17_Click;
                }

                _Button17 = value;
                if (_Button17 != null)
                {
                    _Button17.Click += Button17_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button18;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button18
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button18;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button18 != null)
                {
                    _Button18.Click -= Button18_Click;
                }

                _Button18 = value;
                if (_Button18 != null)
                {
                    _Button18.Click += Button18_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu8;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu8
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu8;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu8 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button20;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button20
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button20;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button20 != null)
                {
                    _Button20.Click -= Button20_Click_1;
                }

                _Button20 = value;
                if (_Button20 != null)
                {
                    _Button20.Click += Button20_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button21;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button21
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button21;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button21 != null)
                {
                    _Button21.Click -= Button21_Click_1;
                }

                _Button21 = value;
                if (_Button21 != null)
                {
                    _Button21.Click += Button21_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu10;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu10
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu10;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu10 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button23;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button23
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button23;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button23 != null)
                {
                    _Button23.Click -= Button23_Click_1;
                }

                _Button23 = value;
                if (_Button23 != null)
                {
                    _Button23.Click += Button23_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button22;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button22
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button22;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button22 != null)
                {
                    _Button22.Click -= Button22_Click_1;
                }

                _Button22 = value;
                if (_Button22 != null)
                {
                    _Button22.Click += Button22_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu9;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu9
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu9;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu9 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button45;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button45
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button45;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button45 != null)
                {
                    _Button45.Click -= Button45_Click;
                }

                _Button45 = value;
                if (_Button45 != null)
                {
                    _Button45.Click += Button45_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button46;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button46
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button46;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button46 != null)
                {
                    _Button46.Click -= Button46_Click;
                }

                _Button46 = value;
                if (_Button46 != null)
                {
                    _Button46.Click += Button46_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button47;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button47
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button47;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button47 != null)
                {
                    _Button47.Click -= Button47_Click;
                }

                _Button47 = value;
                if (_Button47 != null)
                {
                    _Button47.Click += Button47_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonMenu _Menu11;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonMenu Menu11
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Menu11;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Menu11 = value;
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button54;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button54
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button54;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button54 != null)
                {
                    _Button54.Click -= Button54_Click;
                }

                _Button54 = value;
                if (_Button54 != null)
                {
                    _Button54.Click += Button54_Click;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button11;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button11
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button11;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button11 != null)
                {
                    _Button11.Click -= Button11_Click_1;
                }

                _Button11 = value;
                if (_Button11 != null)
                {
                    _Button11.Click += Button11_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button24;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button24
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button24;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button24 != null)
                {
                    _Button24.Click -= Button24_Click_1;
                }

                _Button24 = value;
                if (_Button24 != null)
                {
                    _Button24.Click += Button24_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button25;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button25
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button25;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button25 != null)
                {
                    _Button25.Click -= Button25_Click_1;
                }

                _Button25 = value;
                if (_Button25 != null)
                {
                    _Button25.Click += Button25_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button26;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button26
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button26;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button26 != null)
                {
                    _Button26.Click -= Button26_Click_1;
                }

                _Button26 = value;
                if (_Button26 != null)
                {
                    _Button26.Click += Button26_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button27;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button27
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button27;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Button27 != null)
                {
                    _Button27.Click -= Button27_Click_1;
                }

                _Button27 = value;
                if (_Button27 != null)
                {
                    _Button27.Click += Button27_Click_1;
                }
            }
        }
        private Microsoft.Office.Tools.Ribbon.RibbonButton _Button4;

        internal virtual Microsoft.Office.Tools.Ribbon.RibbonButton Button4
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Button4;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _Button4 = value;
            }
        }
    }

    internal partial class ThisRibbonCollection
    {

        [System.Diagnostics.DebuggerNonUserCode()]
        internal Ribbon1 Ribbon1
        {
            get
            {
                return GetRibbon<Ribbon1>();
            }
        }
    }
}