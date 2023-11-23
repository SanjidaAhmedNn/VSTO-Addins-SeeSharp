using System.Runtime.CompilerServices;

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
            _Tab1.SuspendLayout();
            SuspendLayout();
            // 
            // Tab1
            // 
            _Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            _Tab1.Groups.Add(_Group1);
            _Tab1.Label = "TabAddIns";
            _Tab1.Name = "_Tab1";
            // 
            // Group1
            // 
            _Group1.Label = "Group1";
            _Group1.Name = "_Group1";
            // 
            // Ribbon1
            // 
            Name = "Ribbon1";
            RibbonType = "Microsoft.Excel.Workbook";
            Tabs.Add(_Tab1);
            _Tab1.ResumeLayout(false);
            _Tab1.PerformLayout();
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