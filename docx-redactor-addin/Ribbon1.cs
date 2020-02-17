using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace docx_redactor_addin {
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility {
        private Office.IRibbonUI _ribbon;

        public Ribbon1() {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId) {
            return GetResourceText("docx_redactor_addin.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void GetButtonId(Office.IRibbonControl control) {
            var currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Text = "This text was added by the context menu named My Button.";
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
            _ribbon = ribbonUi;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var t in resourceNames) {
                if (string.Compare(resourceName, t, StringComparison.OrdinalIgnoreCase) != 0) continue;
                using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(t) ?? throw new InvalidOperationException())) {
                    return resourceReader.ReadToEnd();
                }
            }
            return null;
        }

        #endregion
    }
}