using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace docx_redactor_addin
{
    public partial class ThisAddIn
    {
        private static void OnStartup(object sender, System.EventArgs e) {
        }

        private static void OnShutdown(object sender, System.EventArgs e) {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup() {
            Startup += OnStartup;
            Shutdown += OnShutdown;
        }

        #endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return new Ribbon1();
        }

        private void FindColors(Word.Document doc) {
            Word.Paragraphs paragraphs = doc.Paragraphs;
            HashSet<Word.WdColorIndex> colors = new HashSet<Word.WdColorIndex>();

            foreach (var obj in paragraphs) {
                if (!(obj is Word.Paragraph paragraph)) continue;

                Word.Range range = paragraph.Range;
                Word.WdColorIndex color = range.HighlightColorIndex;
                colors.Add(color);
            }

            foreach (Word.WdColorIndex color in colors)
                Debug.WriteLine(color);
        }
    }
}
