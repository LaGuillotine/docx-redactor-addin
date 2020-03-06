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

        public event RedactEvent RedactLikeThis;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId) {
            return GetResourceText("docx_redactor_addin.ContextMenuExtension.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void RedactAction(Office.IRibbonControl control) {
            Range selectedRange = Globals.DocxRedactorAddIn.Application.Selection.Range;

            selectedRange = ExtendByOneIfLengthIsZero(selectedRange);
            selectedRange = ShrinkRangeUntilOnlyOneHighlight(selectedRange);

            switch (control.Id)
            {
                case "redactLikeThis":
                    RedactLikeThis?.Invoke(selectedRange);
                    break;
            }
        }

        #endregion

        #region Helpers

        private static Range ExtendByOneIfLengthIsZero(Range range) {
            if (range.Length() > 0) return range;

            int documentLength = Globals.DocxRedactorAddIn.Application.ActiveDocument.Content.End;
            if (range.End == documentLength) range.Start -= 1;
            else range.End += 1;

            return range;
        }

        private static Range ShrinkRangeUntilOnlyOneHighlight(Range range) {
            bool isHighlightDefined(Range r) => Enum.IsDefined(typeof(WdColorIndex), r.HighlightColorIndex);

            while (range.Length() > 0 && !isHighlightDefined(range)) range.End -= range.Length() / 2;

            return range;
        }

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

    internal static class RangeX10 {
        public static int Length(this Range range) {
            return range.End - range.Start;
        }
    }
}