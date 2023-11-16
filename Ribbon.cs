using System;
using Office = Microsoft.Office.Core;

namespace VSTO_Addins
{
    // TODO:  Follow these steps to enable the Ribbon (XML) item:

    // 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

    // Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
    // Return New Ribbon()
    // End Function

    // 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
    // actions, such as clicking a button. Note: if you have exported this Ribbon from the
    // Ribbon designer, move your code from the event handlers to the callback methods and
    // modify the code to work with the Ribbon extensibility (RibbonX) programming model.

    // 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

    // For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

    [System.Runtime.InteropServices.ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {

        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("VSTO_Addins.Ribbon.xml");
        }

        #region Ribbon Callbacks
        // Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }



        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0, loopTo = resourceNames.Length - 1; i <= loopTo; i++)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resourceReader = new System.IO.StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader is not null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

    }
}