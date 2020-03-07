using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using Document = Microsoft.Office.Interop.Word.Document;

namespace docx_redactor_addin {
    public class Redactor {
        private readonly Document _document;
        private readonly Dictionary<HighlightColorValues, List<Run>> _color2Run = new Dictionary<HighlightColorValues, List<Run>>();

        private XmlDocument _xmlDocument;
        private XmlNode _documentNode;

        public Redactor(Document document) {
            _document = document;
            BuildColorMap();
        }

        public bool KeepHighlight { get; set; } = false;

        public string Replacement { get; set; } = "redacted";

        private void BuildColorMap() {
            string xml = _document.Content.WordOpenXML;
            string documentXml = ExtractDocumentXmlElement(xml);

            OpenXmlDocument doc = new OpenXmlDocument(documentXml);

            void addRun(Run run) {
                if (run.RunProperties?.Highlight == null) return;

                HighlightColorValues color = run.RunProperties.Highlight.Val.Value;
                bool exists = _color2Run.TryGetValue(color, out List<Run> runs);

                if (!exists) {
                    runs = new List<Run>();
                    _color2Run[color] = runs;
                }

                runs.Add(run);
            }

            foreach (Run run in CollectRuns(doc)) {
                addRun(run);
            }
        }

        private string ExtractDocumentXmlElement(string xml) {
            using (TextReader tr = new StringReader(xml)) {
                _xmlDocument = new XmlDocument();
                _xmlDocument.Load(tr);
            }

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


            XmlNode currentRoot = _xmlDocument.FirstChild;
            while (_documentNode == null && currentRoot != null) {
                _documentNode = searchDocument(currentRoot);
                currentRoot = currentRoot.NextSibling;
            }

            return _documentNode != null ? _documentNode.OuterXml : string.Empty;
        }

        private static IEnumerable<Run> CollectRuns(OpenXmlElement document) {
            List<Run> runs = new List<Run>();

            void collect(OpenXmlElement element) {
                if (element is Run run) {
                    runs.Add(run);
                    return;
                }

                foreach (OpenXmlElement child in element.ChildElements)
                    collect(child);
            }

            collect(document);

            return runs;
        }

        public void RedactLikeRange(Range range) {
            HighlightColorValues color = ColorIndex2Value(range.HighlightColorIndex);
            if (!_color2Run.TryGetValue(color, out List<Run> runs)) return;

            foreach (Run currentRun in runs) {
                PerformReplacement(currentRun);
            }

            if (!KeepHighlight) _color2Run.Remove(color);
            
            ApplyChanges();
        }

        private void PerformReplacement(Run run) {
            run.InnerXml = Replacement;

            if (KeepHighlight) return;
            run.RunProperties?.Highlight?.Remove();
        }

        private void ApplyChanges() {
            OpenXmlElement findRoot() {
                Run someRun = _color2Run.Values.First().First();

                OpenXmlElement node = someRun;
                do {
                    node = node.Parent;
                } while (node.Parent != null);

                return node;
            }

            OpenXmlElement root = findRoot();

            if (_documentNode.ParentNode == null) return;
            _documentNode.ParentNode.InnerXml = root.OuterXml;

            string xml = _xmlDocument.OuterXml;
            _document.Content.InsertXML(xml);
        }

        private static HighlightColorValues ColorIndex2Value(WdColorIndex colorIndex) {
            switch (colorIndex)
            {
                case WdColorIndex.wdNoHighlight:
                    return HighlightColorValues.None;
                case WdColorIndex.wdBlack:
                    return HighlightColorValues.Black;
                case WdColorIndex.wdBlue:
                    return HighlightColorValues.Blue;
                case WdColorIndex.wdBrightGreen:
                    return HighlightColorValues.Green;
                case WdColorIndex.wdDarkBlue:
                    return HighlightColorValues.DarkBlue;
                case WdColorIndex.wdDarkRed:
                    return HighlightColorValues.DarkRed;
                case WdColorIndex.wdDarkYellow:
                    return HighlightColorValues.DarkYellow;
                case WdColorIndex.wdGray25:
                    return HighlightColorValues.LightGray;
                case WdColorIndex.wdGray50:
                    return HighlightColorValues.DarkGray;
                case WdColorIndex.wdGreen:
                    return HighlightColorValues.DarkGreen;
                case WdColorIndex.wdPink:
                    return HighlightColorValues.Magenta;
                case WdColorIndex.wdRed:
                    return HighlightColorValues.Red;
                case WdColorIndex.wdTeal:
                    return HighlightColorValues.DarkCyan;
                case WdColorIndex.wdTurquoise:
                    return HighlightColorValues.Cyan;
                case WdColorIndex.wdWhite:
                    return HighlightColorValues.White;
                case WdColorIndex.wdViolet:
                    return HighlightColorValues.DarkMagenta;
                case WdColorIndex.wdYellow:
                    return HighlightColorValues.Yellow;
                case WdColorIndex.wdByAuthor:
                    return HighlightColorValues.None;
                default:
                    return HighlightColorValues.None;
            }
        } 
    }
}