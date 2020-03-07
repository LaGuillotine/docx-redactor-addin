using System.Collections.Generic;
using DocumentFormat.OpenXml;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using Document = Microsoft.Office.Interop.Word.Document;

namespace docx_redactor_addin {
    public class Redactor {
        private readonly Document _document;

        public Redactor(Document document) {
            _document = document;
        }

        public void RedactLikeRange(Range range) {
            HighlightColorValues color = ColorIndex2Value(range.HighlightColorIndex);

            OpenXmlIntegrator integrator = new OpenXmlIntegrator(_document);
            OpenXmlDocument document = integrator.ExtractDocumentElement();
            
            IEnumerable<Run> runs = GetRunsWithColor(document, color);
            foreach (Run currentRun in runs) PerformReplacement(currentRun);

            string xml = integrator.IntegrateAndReturnOpenXml(document);
            _document.Content.InsertXML(xml);
        }

        private static void PerformReplacement(Run run) {
            HighlightColorValues? color = run.RunProperties?.Highlight?.Val.Value;
            run.RunProperties?.Remove();

            foreach (OpenXmlElement child in run.ChildElements) {
                if (!(child is Text text)) continue;
                text.Remove();
            }

            run.AppendChild(new Text(Settings.Replacement));

            if (!Settings.KeepHighlight || color == null) return;
            run.RunProperties = new RunProperties { Highlight = new Highlight { Val = color.Value } };
        }

        private static IEnumerable<Run> GetRunsWithColor(OpenXmlElement document, HighlightColorValues color) {
            List<Run> runs = new List<Run>();

            foreach (Run run in CollectRuns(document)) {
                if (run.RunProperties?.Highlight == null) continue;

                HighlightColorValues highlightColor = run.RunProperties.Highlight.Val.Value;
                if (highlightColor == color) runs.Add(run);
            }

            return runs;
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

        public static class Settings {
            public static bool KeepHighlight { get; set; } = true;
            public static string Replacement { get; set; } = "redacted";
        }
    }
}