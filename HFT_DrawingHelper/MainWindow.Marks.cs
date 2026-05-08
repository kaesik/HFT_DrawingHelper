using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using TS = Tekla.Structures;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;

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

        private void AutoWeldMarkSuffixButton_Click(object sender, RoutedEventArgs e) {
            ApplyAutomaticWeldMarkSuffix();
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
                if (!(selectedObjects.Current is TSD.MarkBase mark))
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

        private void ApplyAutomaticWeldMarkSuffix() {
            var drawingHandler = GetConnectedDrawingHandler();

            if (drawingHandler == null)
                return;

            var activeDrawing = drawingHandler.GetActiveDrawing();

            if (activeDrawing == null) {
                SetMarksStatus("Brak otwartego rysunku.");
                return;
            }

            var selectedObjects = drawingHandler.GetDrawingObjectSelector().GetSelected();

            var selectedMarks = 0;
            var changed = 0;
            var skipped = 0;
            var errors = 0;

            var withoutHashSuffix = 0;
            var withoutModelPart = 0;
            var withoutView = 0;
            var withoutMainPart = 0;
            var withoutDepthDecision = 0;

            while (selectedObjects.MoveNext()) {
                if (!(selectedObjects.Current is TSD.MarkBase mark))
                    continue;

                selectedMarks++;

                try {
                    var suffixTarget = FindSuffixTarget(mark);

                    if (suffixTarget == null ||
                        !string.Equals(suffixTarget.CurrentSuffix, "-#", StringComparison.OrdinalIgnoreCase)) {
                        skipped++;
                        withoutHashSuffix++;
                        continue;
                    }

                    if (!TryGetModelPartFromMark(mark, out var markPart) || markPart == null) {
                        skipped++;
                        withoutModelPart++;
                        continue;
                    }

                    if (!(markPart is TSM.ContourPlate)) {
                        MessageBox.Show("działa poprawnie tylko na counterplate", "Automatyczne markowanie",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        SetMarksStatus("Przerwano: działa poprawnie tylko na counterplate.");
                        return;
                    }

                    if (!TryGetDrawingView(mark, out var view) || view == null) {
                        skipped++;
                        withoutView++;
                        continue;
                    }

                    if (!TryGetMainPartFromView(view, markPart, out var mainPart) || mainPart == null) {
                        skipped++;
                        withoutMainPart++;
                        continue;
                    }

                    var newSuffix = GetAutomaticSuffix(markPart, mainPart);

                    if (string.IsNullOrEmpty(newSuffix)) {
                        skipped++;
                        withoutDepthDecision++;
                        continue;
                    }

                    suffixTarget.SetSuffix(newSuffix);
                    mark.Modify();
                    changed++;
                }
                catch {
                    errors++;
                }
            }

            if (changed > 0)
                activeDrawing.CommitChanges();

            if (selectedMarks == 0) {
                SetMarksStatus("Nie zaznaczono żadnych marek.");
                return;
            }

            var diagnostics = new List<string>();

            if (withoutHashSuffix > 0)
                diagnostics.Add("bez -#: " + withoutHashSuffix);

            if (withoutModelPart > 0)
                diagnostics.Add("bez elementu: " + withoutModelPart);

            if (withoutView > 0)
                diagnostics.Add("bez widoku: " + withoutView);

            if (withoutMainPart > 0)
                diagnostics.Add("bez głównego: " + withoutMainPart);

            if (withoutDepthDecision > 0)
                diagnostics.Add("bez decyzji Z: " + withoutDepthDecision);

            var message = BuildResultMessage(changed, skipped, errors, "Automatycznie zmieniono");

            if (diagnostics.Count > 0)
                message += "  |  " + string.Join(", ", diagnostics);

            SetMarksStatus(message);
        }

        private static string GetAutomaticSuffix(TSM.Part markPart, TSM.Part mainPart) {
            var markZ = GetPartCenterGlobalZ(markPart);
            var mainZ = GetPartCenterGlobalZ(mainPart);

            const double tolerance = 1.0;

            if (markZ > mainZ + tolerance)
                return "-o";

            if (markZ < mainZ - tolerance)
                return "-u";

            return string.Empty;
        }

        private static double GetPartCenterGlobalZ(TSM.Part part) {
            var solid = part.GetSolid();

            return (solid.MinimumPoint.Z + solid.MaximumPoint.Z) / 2.0;
        }

        private static bool TryGetMainPartFromView(TSD.View view, TSM.Part markPart, out TSM.Part mainPart) {
            mainPart = null;

            var parts = new List<TSM.Part>();
            var drawingObjects = view.GetAllObjects(typeof(TSD.Part));

            while (drawingObjects.MoveNext()) {
                if (!(drawingObjects.Current is TSD.Part drawingPart))
                    continue;

                try {
                    if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden)
                        continue;
                }
                catch {
                }

                TSM.Part part = null;

                try {
                    part = MyModel.SelectModelObject(drawingPart.ModelIdentifier) as TSM.Part;
                }
                catch {
                    part = null;
                }

                if (part == null)
                    continue;

                if (IsSamePart(part, markPart))
                    continue;

                if (IsMarkSupportPartName(part.Name))
                    continue;

                parts.Add(part);
            }

            mainPart = parts
                .OrderByDescending(GetPartBoundingBoxVolumeSafe)
                .FirstOrDefault();

            return mainPart != null;
        }

        private static bool IsMarkSupportPartName(string name) {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            return name.StartsWith("MS", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("WS", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSamePart(TSM.Part firstPart, TSM.Part secondPart) {
            if (firstPart == null || secondPart == null)
                return false;

            try {
                return firstPart.Identifier.ID == secondPart.Identifier.ID;
            }
            catch {
                return ReferenceEquals(firstPart, secondPart);
            }
        }

        private static double GetPartBoundingBoxVolumeSafe(TSM.Part part) {
            try {
                var solid = part.GetSolid();

                var sizeX = Math.Abs(solid.MaximumPoint.X - solid.MinimumPoint.X);
                var sizeY = Math.Abs(solid.MaximumPoint.Y - solid.MinimumPoint.Y);
                var sizeZ = Math.Abs(solid.MaximumPoint.Z - solid.MinimumPoint.Z);

                return sizeX * sizeY * sizeZ;
            }
            catch {
                return 0.0;
            }
        }

        private static bool TryGetModelPartFromMark(TSD.MarkBase mark, out TSM.Part part) {
            part = null;

            var drawingObjects = new List<object>();
            CollectRelatedDrawingObjects(mark, drawingObjects);

            foreach (var drawingObject in drawingObjects) {
                if (drawingObject is TSD.Part drawingPart) {
                    part = MyModel.SelectModelObject(drawingPart.ModelIdentifier) as TSM.Part;

                    if (part != null)
                        return true;
                }

                if (drawingObject is TSD.ModelObject modelObject) {
                    part = MyModel.SelectModelObject(modelObject.ModelIdentifier) as TSM.Part;

                    if (part != null)
                        return true;
                }
            }

            var identifiers = new List<TS.Identifier>();
            CollectIdentifiers(mark, identifiers, new HashSet<object>(), 0);

            foreach (var identifier in identifiers) {
                part = MyModel.SelectModelObject(identifier) as TSM.Part;

                if (part != null)
                    return true;
            }

            return false;
        }

        private static void CollectRelatedDrawingObjects(object source, ICollection<object> results) {
            if (source == null)
                return;

            var type = source.GetType();

            var methodNames = new[] {
                "GetRelatedObjects",
                "GetObjects",
                "GetModelObjects",
                "GetDrawingObjects"
            };

            foreach (var methodName in methodNames) {
                var method = type.GetMethod(
                    methodName,
                    BindingFlags.Instance | BindingFlags.Public,
                    null,
                    Type.EmptyTypes,
                    null
                );

                if (method == null)
                    continue;

                object value;

                try {
                    value = method.Invoke(source, null);
                }
                catch {
                    continue;
                }

                CollectEnumeratorItems(value, results);
            }
        }

        private static void CollectEnumeratorItems(object value, ICollection<object> results) {
            if (value == null)
                return;

            var enumerator = value as IEnumerator;

            if (enumerator == null && value is IEnumerable enumerable)
                enumerator = enumerable.GetEnumerator();

            if (enumerator == null)
                return;

            while (enumerator.MoveNext())
                if (enumerator.Current != null)
                    results.Add(enumerator.Current);
        }

        private static void CollectIdentifiers(
            object source,
            ICollection<TS.Identifier> identifiers,
            HashSet<object> visited,
            int depth
        ) {
            if (source == null || depth > 4)
                return;

            var sourceType = source.GetType();

            if (source is string || sourceType.IsValueType)
                return;

            if (!visited.Add(source))
                return;

            if (source is TS.Identifier identifier) {
                identifiers.Add(identifier);
                return;
            }

            if (source is IEnumerable enumerable) {
                IEnumerator enumerator;

                try {
                    enumerator = enumerable.GetEnumerator();
                }
                catch {
                    enumerator = null;
                }

                if (enumerator != null)
                    while (enumerator.MoveNext())
                        CollectIdentifiers(enumerator.Current, identifiers, visited, depth + 1);
            }

            foreach (var property in sourceType.GetProperties(BindingFlags.Instance | BindingFlags.Public)) {
                if (property.GetIndexParameters().Length > 0)
                    continue;

                if (property.PropertyType == typeof(string))
                    continue;

                object value;

                try {
                    value = property.GetValue(source, null);
                }
                catch {
                    continue;
                }

                if (value is TS.Identifier valueIdentifier) {
                    identifiers.Add(valueIdentifier);
                    continue;
                }

                if (
                    property.Name.IndexOf("Identifier", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    property.Name.IndexOf("Object", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    property.Name.IndexOf("Part", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    property.Name.IndexOf("Model", StringComparison.OrdinalIgnoreCase) >= 0
                )
                    CollectIdentifiers(value, identifiers, visited, depth + 1);
            }
        }

        private static bool TryGetDrawingView(TSD.MarkBase mark, out TSD.View view) {
            view = null;

            var method = mark.GetType().GetMethod(
                "GetView",
                BindingFlags.Instance | BindingFlags.Public,
                null,
                Type.EmptyTypes,
                null
            );

            if (method == null)
                return false;

            try {
                view = method.Invoke(mark, null) as TSD.View;
            }
            catch {
                view = null;
            }

            return view != null;
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
                var drawingObject = selectedObjects.Current;

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

            SetMarksStatus(BuildResultMessage(changed, 0, errors, show ? "Pokazano" : "Ukryto"));
        }

        private static void ApplyDrawingObjectVisibility(TSD.DrawingObject drawingObject, bool show) {
            var hideable =
                drawingObject.GetType()
                    .GetProperty("Hideable", BindingFlags.Instance | BindingFlags.Public)
                    ?.GetValue(drawingObject, null);

            if (hideable == null)
                throw new InvalidOperationException();

            var methodName = show ? "ShowInDrawingView" : "HideFromDrawingView";

            var method = hideable.GetType().GetMethod(
                methodName,
                BindingFlags.Instance | BindingFlags.Public
            );

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
                var value = textElement.Value ?? string.Empty;

                if (KnownWeldMarkSuffixes.Contains(value))
                    return new SuffixTarget(textElement, value, value.Length, true);

                foreach (var suffix in KnownWeldMarkSuffixes
                             .Where(suffix => !string.IsNullOrEmpty(suffix))
                             .Where(suffix => value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)))
                    return new SuffixTarget(textElement, suffix, suffix.Length, false);
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

                var currentValue = textElement.Value ?? string.Empty;

                textElement.Value = currentValue + value;

                return true;
            }

            return false;
        }

        private static object GetMarkContent(TSD.MarkBase mark) {
            if (mark?.Attributes == null)
                return null;

            var attributesType = mark.Attributes.GetType();

            var contentProperty = attributesType.GetProperty(
                "Content",
                BindingFlags.Instance | BindingFlags.Public
            );

            return contentProperty == null
                ? null
                : contentProperty.GetValue(mark.Attributes, null);
        }

        private static void CollectTextElements(
            object container,
            ICollection<TSD.TextElement> results
        ) {
            if (container == null)
                return;

            IEnumerator enumerator;

            try {
                if (container is IEnumerable enumerable)
                    enumerator = enumerable.GetEnumerator();
                else
                    enumerator =
                        container.GetType()
                            .GetMethod("GetEnumerator", Type.EmptyTypes)
                            ?.Invoke(container, null) as IEnumerator;
            }
            catch {
                return;
            }

            if (enumerator == null)
                return;

            while (enumerator.MoveNext()) {
                var element = enumerator.Current;

                switch (element) {
                    case null:
                        continue;

                    case TSD.TextElement textElement:
                        results.Add(textElement);
                        break;

                    case TSD.ContainerElement nestedContainer:
                        CollectTextElements(nestedContainer, results);
                        break;
                }
            }
        }

        private static string BuildResultMessage(
            int changed,
            int skipped,
            int errors,
            string changedLabel
        ) {
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
            if (_marksStatusTextBlock == null)
                _marksStatusTextBlock =
                    FindNamedDescendant<TextBlock>("MarksStatusTextBlock");

            if (_marksStatusTextBlock == null) {
                MessageBox.Show(message);
                return;
            }

            _marksStatusTextBlock.Text = message;

            _marksStatusTextBlock.Foreground =
                TryFindResource("TextSecondaryBrush") as Brush
                ?? _marksStatusTextBlock.Foreground;
        }

        private sealed class SuffixTarget {
            private readonly bool _replaceWholeValue;
            private readonly int _suffixLength;
            private readonly TSD.TextElement _textElement;

            public SuffixTarget(
                TSD.TextElement textElement,
                string currentSuffix,
                int suffixLength,
                bool replaceWholeValue
            ) {
                _textElement = textElement;
                CurrentSuffix = currentSuffix;
                _suffixLength = suffixLength;
                _replaceWholeValue = replaceWholeValue;
            }

            public string CurrentSuffix { get; private set; }

            public void SetSuffix(string suffix) {
                var value = _textElement.Value ?? string.Empty;

                if (_replaceWholeValue) {
                    _textElement.Value = suffix;
                    CurrentSuffix = suffix;
                    return;
                }

                var baseValueLength = Math.Max(0, value.Length - _suffixLength);

                _textElement.Value = 
                    value.Substring(0, baseValueLength) + suffix;

                CurrentSuffix = suffix;
            }
        }
    }
}