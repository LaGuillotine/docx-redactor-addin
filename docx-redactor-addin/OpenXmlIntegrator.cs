using System;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using Document = Microsoft.Office.Interop.Word.Document;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace docx_redactor_addin {
    public class OpenXmlIntegrator {
        private readonly Document _document;

        private XmlDocument _xmlDocument;
        private XmlNode _documentNode;

        private OpenXmlDocument _documentElement;

        public OpenXmlIntegrator(Document document) {
            _document = document;
        }

        public OpenXmlDocument ExtractDocumentElement() {
            string xml = _document.Content.WordOpenXML;

            using (TextReader tr = new StringReader(xml)) {
                _xmlDocument = new XmlDocument();
                _xmlDocument.Load(tr);
            }

            _documentNode = FindDocumentNode(_xmlDocument);
            if (_documentNode == null) throw new NotSupportedException($"The OpenXml returned by the {_document.GetType()} object is of unexpected format");
            
            _documentElement = new OpenXmlDocument(_documentNode.OuterXml);
            return _documentElement;
        }

        public string IntegrateAndReturnOpenXml(OpenXmlElement openXmlElement) {
            if (openXmlElement == null)
                throw new ArgumentNullException(nameof(openXmlElement));
            if (_documentElement == null) 
                throw new InvalidOperationException("Changes can only be applied once the document element has been extracted");
            if (!(openXmlElement is OpenXmlDocument document))
                throw new ArgumentException($"Argument must be of type '{typeof(OpenXmlDocument)}' but type was '{openXmlElement.GetType()}");
            if (document != _documentElement)
                throw new InvalidOperationException("Changes can not be applied. The argument does not match the original document element");
            if (_documentNode?.ParentNode == null)
                throw new InvalidOperationException("The document node does not have a parent. This should never happen. You stumbled on a bug.");

            _documentNode.ParentNode.InnerXml = document.OuterXml;
            return _xmlDocument.OuterXml;
        }

        private static XmlNode FindDocumentNode(XmlNode xmlDocument) {
            XmlNode searchDocument(XmlNode node) {
                if (node.LocalName == "document") return node;

                XmlNode searchedNode = null;
                foreach (XmlNode child in node.ChildNodes) {
                    XmlNode searchResult = searchDocument(child);

                    if (searchResult?.LocalName != "document") continue;
                    searchedNode = searchResult;
                    break;
                }

                return searchedNode;
            }

            XmlNode currentRoot = xmlDocument.FirstChild;
            XmlNode documentNode = null;
            while (documentNode == null && currentRoot != null) {
                documentNode = searchDocument(currentRoot);
                currentRoot = currentRoot.NextSibling;
            }

            return documentNode;
        }
    }
}
