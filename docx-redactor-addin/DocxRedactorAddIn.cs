using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace docx_redactor_addin {
    public partial class DocxRedactorAddIn {
        private readonly Dictionary<Document, Redactor> _redactors = new Dictionary<Document, Redactor>();
        private ContextMenuExtension _contextMenu;

        private void OnStartup(object sender, System.EventArgs e) {
            Application.DocumentOpen += OnDocumentOpen;
            Application.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        private static void OnShutdown(object sender, System.EventArgs e) { /* Nothing to do here */ }

        private void OnDocumentOpen(Document document) {
            Redactor redactor = new Redactor(document);
            _redactors[document] = redactor;

            _contextMenu.RedactLikeThis += redactor.RedactLikeRange;
        }

        private void OnDocumentBeforeClose(Document document, ref bool cancel) {
            Redactor redactor = _redactors[document];
            _redactors.Remove(document);

            _contextMenu.RedactLikeThis -= redactor.RedactLikeRange;
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

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            _contextMenu = new ContextMenuExtension();
            return _contextMenu;
        }
    }
}
