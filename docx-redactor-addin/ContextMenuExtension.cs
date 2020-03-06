using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace docx_redactor_addin {
    [ComVisible(true)]
    public class ContextMenuExtension : Office.IRibbonExtensibility {
        public delegate void RedactEvent(Range range);

        public event RedactEvent Redact;
        public event RedactEvent RedactLikeThis;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId) {
            return GetResourceText("docx_redactor_addin.ContextMenuExtension.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void RedactAction(Office.IRibbonControl control) {
            Range selectedRange = Globals.DocxRedactorAddIn.Application.Selection.Range;

            if (selectedRange?.Text == null) return;
            if (selectedRange.Text.Length == 0) return;
            
            switch (control.Id)
            {
                case "redact":
                    Redact?.Invoke(selectedRange);
                    break;
                case "redactLikeThis":
                    RedactLikeThis?.Invoke(selectedRange);
                    break;
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string t in resourceNames) {
                if (string.Compare(resourceName, t, StringComparison.OrdinalIgnoreCase) != 0) continue;
                using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(t) ?? throw new InvalidOperationException())) {
                    return resourceReader.ReadToEnd();
                }
            }
            return null;
        }

        #endregion
    }
}