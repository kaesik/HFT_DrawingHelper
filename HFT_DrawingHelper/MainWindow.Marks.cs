using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        private static readonly HashSet<string> KnownWeldMarkSuffixes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "-#",
                "-o",
                "-u",
                string.Empty
            };

        private void SetWeldMarkSuffixOButton_Click(object sender, RoutedEventArgs e) {
            ApplyWeldMarkSuffix("-o");
        }

        private void SetWeldMarkSuffixUButton_Click(object sender, RoutedEventArgs e) {
            ApplyWeldMarkSuffix("-u");
        }

        private void ClearWeldMarkSuffixButton_Click(object sender, RoutedEventArgs e) {
            ApplyWeldMarkSuffix(string.Empty);
        }

        private void RestoreWeldMarkSuffixButton_Click(object sender, RoutedEventArgs e) {
            ApplyWeldMarkSuffix("-#");
        }

        private void HideSelectedDrawingObjectsButton_Click(object sender, RoutedEventArgs e) {
            SetSelectedDrawingObjectsVisibility(false);
        }

        private void ShowSelectedDrawingObjectsButton_Click(object sender, RoutedEventArgs e) {
            SetSelectedDrawingObjectsVisibility(true);
        }

        private void ApplyWeldMarkSuffix(string newValue) {
            var drawingHandler = GetConnectedDrawingHandler();
            if (drawingHandler == null)
                return;

            var activeDrawing = drawingHandler.GetActiveDrawing();
            if (activeDrawing == null) {
                SetMarksStatus("Brak otwartego rysunku.");
                return;
            }

            var selectedObjects = drawingHandler.GetDrawingObjectSelector().GetSelected();
            var changed = 0;
            var skipped = 0;
            var errors = 0;

            while (selectedObjects.MoveNext()) {
                var mark = selectedObjects.Current as TSD.MarkBase;
                if (mark == null)
                    continue;

                var suffixTarget = FindSuffixTarget(mark);
                if (suffixTarget == null) {
                    if (!string.Equals(newValue, "-#", StringComparison.OrdinalIgnoreCase)) {
                        skipped++;
                        continue;
                    }

                    try {
                        if (!TryAddSuffixToMark(mark, newValue)) {
                            skipped++;
                            continue;
                        }

                        mark.Modify();
                        changed++;
                    }
                    catch {
                        errors++;
                    }

                    continue;
                }

                if (string.Equals(suffixTarget.CurrentSuffix, newValue, StringComparison.OrdinalIgnoreCase)) {
                    skipped++;
                    continue;
                }

                try {
                    suffixTarget.SetSuffix(newValue);
                    mark.Modify();
                    changed++;
                }
                catch {
                    errors++;
                }
            }

            if (changed > 0)
                activeDrawing.CommitChanges();

            if (changed == 0 && skipped == 0 && errors == 0) {
                SetMarksStatus("Nie zaznaczono żadnych marek.");
                return;
            }

            SetMarksStatus(BuildResultMessage(changed, skipped, errors, "Zmieniono"));
        }

        private void SetSelectedDrawingObjectsVisibility(bool show) {
            var drawingHandler = GetConnectedDrawingHandler();
            if (drawingHandler == null)
                return;

            var activeDrawing = drawingHandler.GetActiveDrawing();
            if (activeDrawing == null) {
                SetMarksStatus("Brak otwartego rysunku.");
                return;
            }

            var selectedObjects = drawingHandler.GetDrawingObjectSelector().GetSelected();
            var changed = 0;
            var errors = 0;

            while (selectedObjects.MoveNext()) {
                var drawingObject = selectedObjects.Current as TSD.DrawingObject;
                if (drawingObject == null)
                    continue;

                try {
                    ApplyDrawingObjectVisibility(drawingObject, show);
                    drawingObject.Modify();
                    changed++;
                }
                catch {
                    errors++;
                }
            }

            if (changed > 0)
                activeDrawing.CommitChanges();

            if (changed == 0 && errors == 0) {
                SetMarksStatus("Nie zaznaczono żadnych obiektów.");
                return;
            }

            SetMarksStatus(BuildResultMessage(changed, 0, errors, show ? "Pokazano" : "Ukryto"));
        }

        private static void ApplyDrawingObjectVisibility(TSD.DrawingObject drawingObject, bool show) {
            var hideable = drawingObject.GetType().GetProperty("Hideable", BindingFlags.Instance | BindingFlags.Public)?.GetValue(drawingObject, null);
            if (hideable == null)
                throw new InvalidOperationException();

            var methodName = show ? "ShowInDrawingView" : "HideFromDrawingView";
            var method = hideable.GetType().GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public);
            if (method == null)
                throw new InvalidOperationException();

            method.Invoke(hideable, null);
        }

        private TSD.DrawingHandler GetConnectedDrawingHandler() {
            var drawingHandler = new TSD.DrawingHandler();
            if (drawingHandler.GetConnectionStatus())
                return drawingHandler;

            SetMarksStatus("Brak połączenia z Tekla.");
            return null;
        }

        private static SuffixTarget FindSuffixTarget(TSD.MarkBase mark) {
            var textElements = new List<TSD.TextElement>();
            CollectTextElements(GetMarkContent(mark), textElements);

            for (var i = textElements.Count - 1; i >= 0; i--) {
                var textElement = textElements[i];
                var value = textElement.Value;
                if (value == null)
                    value = string.Empty;

                if (KnownWeldMarkSuffixes.Contains(value))
                    return new SuffixTarget(textElement, value, value.Length, true);

                foreach (var suffix in KnownWeldMarkSuffixes) {
                    if (string.IsNullOrEmpty(suffix))
                        continue;

                    if (!value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                        continue;

                    return new SuffixTarget(textElement, suffix, suffix.Length, false);
                }
            }

            return null;
        }

        private static bool TryAddSuffixToMark(TSD.MarkBase mark, string value) {
            var textElements = new List<TSD.TextElement>();
            CollectTextElements(GetMarkContent(mark), textElements);

            for (var i = textElements.Count - 1; i >= 0; i--) {
                var textElement = textElements[i];
                if (textElement == null)
                    continue;

                var currentValue = textElement.Value;
                if (currentValue == null)
                    currentValue = string.Empty;

                textElement.Value = currentValue + value;
                return true;
            }

            return false;
        }

        private static object GetMarkContent(TSD.MarkBase mark) {
            if (mark == null || mark.Attributes == null)
                return null;

            var attributesType = mark.Attributes.GetType();
            var contentProperty = attributesType.GetProperty("Content", BindingFlags.Instance | BindingFlags.Public);
            return contentProperty == null ? null : contentProperty.GetValue(mark.Attributes, null);
        }

        private static void CollectTextElements(object container, ICollection<TSD.TextElement> results) {
            if (container == null)
                return;

            IEnumerator enumerator;
            try {
                var enumerable = container as IEnumerable;
                if (enumerable != null)
                    enumerator = enumerable.GetEnumerator();
                else
                    enumerator = container.GetType().GetMethod("GetEnumerator", Type.EmptyTypes)?.Invoke(container, null) as IEnumerator;
            }
            catch {
                return;
            }

            if (enumerator == null)
                return;

            while (enumerator.MoveNext()) {
                var element = enumerator.Current;
                if (element == null)
                    continue;

                var textElement = element as TSD.TextElement;
                if (textElement != null)
                    results.Add(textElement);

                var nestedContainer = element as TSD.ContainerElement;
                if (nestedContainer != null)
                    CollectTextElements(nestedContainer, results);
            }
        }

        private static string BuildResultMessage(int changed, int skipped, int errors, string changedLabel) {
            var parts = new List<string>();

            if (changed > 0)
                parts.Add(changedLabel + ": " + changed);

            if (skipped > 0)
                parts.Add("Pominięto: " + skipped);

            if (errors > 0)
                parts.Add("Błędy: " + errors);

            return string.Join("  |  ", parts);
        }

        private void SetMarksStatus(string message) {
            if (_marksStatusTextBlock == null) {
                MessageBox.Show(message);
                return;
            }

            _marksStatusTextBlock.Text = message;
            _marksStatusTextBlock.Foreground = TryFindResource("TextSecondaryBrush") as Brush ?? _marksStatusTextBlock.Foreground;
        }

        private sealed class SuffixTarget {
            private readonly TSD.TextElement _textElement;
            private readonly int _suffixLength;
            private readonly bool _replaceWholeValue;

            public SuffixTarget(TSD.TextElement textElement, string currentSuffix, int suffixLength, bool replaceWholeValue) {
                _textElement = textElement;
                CurrentSuffix = currentSuffix;
                _suffixLength = suffixLength;
                _replaceWholeValue = replaceWholeValue;
            }

            public string CurrentSuffix { get; private set; }

            public void SetSuffix(string suffix) {
                var value = _textElement.Value;
                if (value == null)
                    value = string.Empty;

                if (_replaceWholeValue) {
                    _textElement.Value = suffix;
                    CurrentSuffix = suffix;
                    return;
                }

                var baseValueLength = Math.Max(0, value.Length - _suffixLength);
                _textElement.Value = value.Substring(0, baseValueLength) + suffix;
                CurrentSuffix = suffix;
            }
        }
    }
}
