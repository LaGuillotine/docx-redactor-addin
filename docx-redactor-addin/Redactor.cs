using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace docx_redactor_addin {
    public class Redactor {
        private readonly Document _document;

        private readonly Dictionary<WdColorIndex, List<Range>> _color2Run = new Dictionary<WdColorIndex, List<Range>>();

        public Redactor(Document document) {
            _document = document;
            BuildColorMap();
        }

        public bool KeepHighlight { get; set; } = false;

        public string Replacement { get; set; } = "redacted";

        private void BuildColorMap() {
            Paragraphs paragraphs = _document.Paragraphs;
            Tables tables = _document.Tables;

            void addRange(Range range) {
                WdColorIndex color = range.HighlightColorIndex;

                bool exists = _color2Run.TryGetValue(color, out List<Range> ranges);

                if (!exists) {
                    ranges = new List<Range>();
                    _color2Run[color] = ranges;
                }

                ranges.Add(range);
            }

            foreach (object obj in paragraphs) {
                if (!(obj is Paragraph paragraph)) continue;

                Range range = paragraph.Range;
                addRange(range);
            }

            foreach (Table table in tables) {
                Range range = table.Range;
                addRange(range);
            }
        }

        public void Redact(Range range) {
            PerformReplacement(range);
            RemoveFromColorDictionary(range);
        }

        public void RedactLikeRange(Range range) {
            WdColorIndex color = range.HighlightColorIndex;
            if (!_color2Run.TryGetValue(color, out List<Range> ranges)) return;

            foreach (Range currentRange in ranges) {
                PerformReplacement(currentRange);
            }

            _color2Run.Remove(color);
        }

        private void PerformReplacement(Range range) {
            range.Text = Replacement;
            if (KeepHighlight) return;
            range.HighlightColorIndex = WdColorIndex.wdNoHighlight;
        }

        private void RemoveFromColorDictionary(Range range) {
            WdColorIndex color = range.HighlightColorIndex;
            if (!_color2Run.TryGetValue(color, out List<Range> ranges)) return;

            ranges.Remove(range);

            if (ranges.Count == 0) _color2Run.Remove(color);
        }
    }
}