using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
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

                var textElement = FindSuffixTextElement(mark, out var currentValue);
                if (textElement == null) {
                    if (!string.Equals(newValue, "-#", StringComparison.OrdinalIgnoreCase)) {
                        skipped++;
                        continue;
                    }

                    try {
                        if (!TryAppendSuffixTextElement(mark, newValue)) {
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

                if (currentValue == null)
                    currentValue = string.Empty;

                if (!KnownWeldMarkSuffixes.Contains(currentValue)) {
                    skipped++;
                    continue;
                }

                if (string.Equals(currentValue, newValue, StringComparison.OrdinalIgnoreCase)) {
                    skipped++;
                    continue;
                }

                try {
                    SetTextElementValue(textElement, newValue);
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
            var hideable = drawingObject.GetType().GetProperty("Hideable", BindingFlags.Instance | BindingFlags.Public)
                ?.GetValue(drawingObject, null);
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

        private static object FindSuffixTextElement(TSD.MarkBase mark, out string value) {
            value = null;
            var textElements = new List<object>();
            CollectTextElements(GetMarkContent(mark), textElements);

            if (textElements.Count == 0)
                return null;

            for (var i = textElements.Count - 1; i >= 0; i--) {
                var currentTextElement = textElements[i];
                var currentValue = GetTextElementValue(currentTextElement);
                if (currentValue == null)
                    currentValue = string.Empty;

                if (!KnownWeldMarkSuffixes.Contains(currentValue))
                    continue;

                value = currentValue;
                return currentTextElement;
            }

            return null;
        }

        private static bool TryAppendSuffixTextElement(TSD.MarkBase mark, string value) {
            var content = GetMarkContent(mark);
            if (content == null)
                return false;

            var textElements = new List<object>();
            CollectTextElements(content, textElements);
            var referenceTextElement = textElements.Count > 0 ? textElements[textElements.Count - 1] : null;
            var textElement = CreateTextElement(value, referenceTextElement);
            if (textElement == null)
                return false;

            return TryAddElementToContainer(content, textElement);
        }

        private static object CreateTextElement(string value, object referenceTextElement) {
            var drawingAssembly = typeof(TSD.MarkBase).Assembly;
            var textElementType = drawingAssembly.GetType("Tekla.Structures.Drawing.TextElement");
            if (textElementType == null)
                return null;

            object textElement = null;

            try {
                textElement = Activator.CreateInstance(textElementType, value);
            }
            catch {
                try {
                    textElement = Activator.CreateInstance(textElementType);
                }
                catch {
                    return null;
                }
            }

            CopyTextElementStyle(referenceTextElement, textElement);
            SetTextElementValue(textElement, value);
            return textElement;
        }

        private static void CopyTextElementStyle(object source, object target) {
            if (source == null || target == null)
                return;

            var sourceType = source.GetType();
            var targetType = target.GetType();
            foreach (var sourceProperty in sourceType.GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
                if (!sourceProperty.CanRead || sourceProperty.GetIndexParameters().Length > 0)
                    continue;

                if (string.Equals(sourceProperty.Name, "Value", StringComparison.OrdinalIgnoreCase))
                    continue;

                var targetProperty = targetType.GetProperty(sourceProperty.Name, BindingFlags.Instance | BindingFlags.Public);
                if (targetProperty == null || !targetProperty.CanWrite || targetProperty.GetIndexParameters().Length > 0)
                    continue;

                if (!targetProperty.PropertyType.IsAssignableFrom(sourceProperty.PropertyType))
                    continue;

                try {
                    targetProperty.SetValue(target, sourceProperty.GetValue(source, null), null);
                }
                catch {
                }
            }
        }

        private static bool TryAddElementToContainer(object container, object element) {
            var containerType = container.GetType();
            var elementType = element.GetType();

            foreach (var method in containerType.GetMethods(BindingFlags.Instance | BindingFlags.Public)) {
                if (!string.Equals(method.Name, "Add", StringComparison.OrdinalIgnoreCase))
                    continue;

                var parameters = method.GetParameters();
                if (parameters.Length != 1)
                    continue;

                if (!parameters[0].ParameterType.IsAssignableFrom(elementType))
                    continue;

                method.Invoke(container, new[] { element });
                return true;
            }

            return false;
        }

        private static object GetMarkContent(TSD.MarkBase mark) {
            if (mark == null || mark.Attributes == null)
                return null;

            return mark.Attributes.GetType().GetProperty("Content", BindingFlags.Instance | BindingFlags.Public)
                ?.GetValue(mark.Attributes, null);
        }

        private static void CollectTextElements(object container, ICollection<object> results) {
            if (container == null)
                return;

            IEnumerator enumerator;

            try {
                if (container is IEnumerable enumerable)
                    enumerator = enumerable.GetEnumerator();
                else
                    enumerator = container.GetType().GetMethod("GetEnumerator", Type.EmptyTypes)
                        ?.Invoke(container, null) as IEnumerator;
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

                var typeName = element.GetType().Name;
                if (string.Equals(typeName, "TextElement", StringComparison.OrdinalIgnoreCase))
                    results.Add(element);
                else if (string.Equals(typeName, "ContainerElement", StringComparison.OrdinalIgnoreCase))
                    CollectTextElements(element, results);
            }
        }

        private static string GetTextElementValue(object textElement) {
            return textElement?.GetType().GetProperty("Value", BindingFlags.Instance | BindingFlags.Public)
                ?.GetValue(textElement, null) as string;
        }

        private static void SetTextElementValue(object textElement, string value) {
            textElement.GetType().GetProperty("Value", BindingFlags.Instance | BindingFlags.Public)
                ?.SetValue(textElement, value, null);
        }

        private static string BuildResultMessage(int changed, int skipped, int errors, string changedLabel) {
            var parts = new List<string>();

            if (changed > 0)
                parts.Add($"{changedLabel}: {changed}");

            if (skipped > 0)
                parts.Add($"Pominięto: {skipped}");

            if (errors > 0)
                parts.Add($"Błędy: {errors}");

            return string.Join("  |  ", parts);
        }

        private void SetMarksStatus(string message) {
            var marksStatusTextBlock = FindNamedDescendant<TextBlock>("MarksStatusTextBlock");
            if (marksStatusTextBlock == null) {
                MessageBox.Show(message);
                return;
            }

            marksStatusTextBlock.Text = message;
        }
    }
}
