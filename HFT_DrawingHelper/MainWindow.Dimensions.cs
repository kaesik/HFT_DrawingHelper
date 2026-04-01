using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Bounds Calculation

        private static PartBounds GetAssemblyBounds(List<PartBounds> parts) {
            if (parts == null || parts.Count == 0) return null;

            var minX = double.MaxValue;
            var minY = double.MaxValue;
            var maxX = double.MinValue;
            var maxY = double.MinValue;
            var found = false;

            foreach (var p in parts.Where(p => p != null)) {
                if (p.MinX < minX) minX = p.MinX;
                if (p.MinY < minY) minY = p.MinY;
                if (p.MaxX > maxX) maxX = p.MaxX;
                if (p.MaxY > maxY) maxY = p.MaxY;
                found = true;
            }

            var result = !found ? null : new PartBounds { MinX = minX, MaxX = maxX, MinY = minY, MaxY = maxY };
            return result;
        }

        #endregion

        #region Geometry Helpers

        private static List<double> MergeAndSort(IEnumerable<double> rawValues, double tolerance) {
            var sorted = rawValues
                .Where(v => !double.IsNaN(v) && !double.IsInfinity(v))
                .OrderBy(v => v)
                .ToList();

            if (sorted.Count == 0) return new List<double>();

            var result = new List<double>();
            var bucketStart = sorted[0];
            var bucket = new List<double> { sorted[0] };

            for (var i = 1; i < sorted.Count; i++)
                if (sorted[i] - bucketStart <= tolerance)
                    bucket.Add(sorted[i]);
                else {
                    result.Add(bucket.Average());
                    bucketStart = sorted[i];
                    bucket = new List<double> { sorted[i] };
                }

            result.Add(bucket.Average());
            return result;
        }

        private static List<TSG.Point> EnsureClosedOutline(List<TSG.Point> outline) {
            if (outline == null || outline.Count < 2) return outline;

            var first = outline[0];
            var last = outline[outline.Count - 1];

            var distance = Math.Sqrt(
                (first.X - last.X) * (first.X - last.X) +
                (first.Y - last.Y) * (first.Y - last.Y)
            );

            if (distance > DuplicateToleranceMillimeters)
                outline.Add(new TSG.Point(first.X, first.Y, 0));

            return outline;
        }

        private static List<TSG.Point> GetOpenOutlineVertices(List<TSG.Point> outline) {
            if (outline == null) return null;

            var result = RemoveNearDuplicates(new List<TSG.Point>(outline), DuplicateToleranceMillimeters);
            if (result == null || result.Count == 0) return result;

            if (result.Count > 1) {
                var first = result[0];
                var last = result[result.Count - 1];

                var distance = Math.Sqrt(
                    (first.X - last.X) * (first.X - last.X) +
                    (first.Y - last.Y) * (first.Y - last.Y)
                );

                if (distance <= DuplicateToleranceMillimeters)
                    result.RemoveAt(result.Count - 1);
            }

            return result;
        }

        private static double GetCornerAngleDegrees(TSG.Point previous, TSG.Point current, TSG.Point next) {
            var firstVectorX = current.X - previous.X;
            var firstVectorY = current.Y - previous.Y;
            var secondVectorX = next.X - current.X;
            var secondVectorY = next.Y - current.Y;

            var firstLength = Math.Sqrt(firstVectorX * firstVectorX + firstVectorY * firstVectorY);
            var secondLength = Math.Sqrt(secondVectorX * secondVectorX + secondVectorY * secondVectorY);

            if (firstLength < 1e-9 || secondLength < 1e-9) return 0.0;

            var dot = (firstVectorX * secondVectorX + firstVectorY * secondVectorY) / (firstLength * secondLength);
            if (dot > 1.0) dot = 1.0;
            if (dot < -1.0) dot = -1.0;

            return Math.Acos(dot) * 180.0 / Math.PI;
        }

        private static List<int> GetSignificantCornerIndices(List<TSG.Point> vertices) {
            var result = new List<int>();
            if (vertices == null || vertices.Count < 3) return result;

            for (var i = 0; i < vertices.Count; i++) {
                var previousIndex = i == 0 ? vertices.Count - 1 : i - 1;
                var nextIndex = i == vertices.Count - 1 ? 0 : i + 1;

                var angleDegrees = GetCornerAngleDegrees(
                    vertices[previousIndex],
                    vertices[i],
                    vertices[nextIndex]
                );

                LogDebug("corner[" + i + "] angle = " + angleDegrees.ToString("0.###"));

                if (angleDegrees >= SignificantCornerAngleDegrees)
                    result.Add(i);
            }

            return result;
        }

        private static List<TSG.Point> CollectPathBetweenCorners(
            List<TSG.Point> vertices,
            int startIndex,
            int endIndex
        ) {
            var result = new List<TSG.Point>();
            if (vertices == null || vertices.Count == 0) return result;

            var index = startIndex;
            result.Add(vertices[index]);

            while (index != endIndex) {
                index++;
                if (index >= vertices.Count) index = 0;
                result.Add(vertices[index]);
            }

            return result;
        }

        private static List<TSG.Point> PrepareSegmentPath(List<TSG.Point> points) {
            if (points == null) return null;

            var result = RemoveNearDuplicates(new List<TSG.Point>(points), DuplicateToleranceMillimeters);
            result = SimplifyPolyline(result);
            result = RemoveNearDuplicates(result, DuplicateToleranceMillimeters);

            return result;
        }

        #endregion

        #region Entry Point

        private static readonly StringBuilder DebugBuilder = new StringBuilder();
        private static string _debugFilePath;

        private static void AddDimensions(DimensionOptions options) {
            StartDebugSession();

            try {
                LogDebug("=== START AddDimensions ===");

                var drawingHandler = new TSD.DrawingHandler();
                if (!drawingHandler.GetConnectionStatus()) {
                    LogDebug("Brak połączenia z drawingHandler.");
                    FlushDebugSession();
                    return;
                }

                var activeDrawing = drawingHandler.GetActiveDrawing();
                if (activeDrawing == null) {
                    LogDebug("Brak aktywnego rysunku.");
                    FlushDebugSession();
                    return;
                }

                LogDebug("Aktywny rysunek pobrany poprawnie.");

                var selectedParts = GetSelectedDrawingParts(drawingHandler);
                LogDebug("Liczba zaznaczonych partów: " + selectedParts.Count);

                if (selectedParts.Count == 0) {
                    LogDebug("Brak zaznaczonych partów -> AddDimensionsFromSelectedView");
                    AddDimensionsFromSelectedView(drawingHandler, activeDrawing, options);
                    FlushDebugSession();
                    return;
                }

                LogDebug("Są zaznaczone party -> AddDimensionsFromSelectedParts");
                AddDimensionsFromSelectedParts(activeDrawing, selectedParts, options);
                FlushDebugSession();
            }
            catch (Exception exception) {
                LogDebug("WYJĄTEK AddDimensions: " + exception);
                FlushDebugSession();
                throw;
            }
        }

        #endregion

        #region Debug

        private static void StartDebugSession() {
            DebugBuilder.Clear();
            _debugFilePath = Path.Combine(
                Path.GetTempPath(),
                "HFT_DrawingHelper_Debug_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt"
            );

            DebugBuilder.AppendLine("HFT Drawing Helper Debug");
            DebugBuilder.AppendLine("Data: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            DebugBuilder.AppendLine(new string('-', 80));
        }

        private static void FlushDebugSession() {
            try {
                File.WriteAllText(_debugFilePath, DebugBuilder.ToString(), Encoding.UTF8);
            }
            catch (Exception exception) {
                MessageBox.Show("Nie udało się zapisać debug loga:\n" + exception.Message);
            }
        }

        private static void LogDebug(string message) {
            DebugBuilder.AppendLine("[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + message);
        }

        private static void LogPoint(string label, TSG.Point point) {
            if (point == null) {
                LogDebug(label + ": null");
                return;
            }

            LogDebug(string.Format(
                "{0}: X={1:0.###}, Y={2:0.###}, Z={3:0.###}",
                label,
                point.X,
                point.Y,
                point.Z
            ));
        }

        private static void LogBounds(string label, PartBounds bounds) {
            if (bounds == null) {
                LogDebug(label + ": null");
                return;
            }

            LogDebug(string.Format(
                "{0}: MinX={1:0.###}, MaxX={2:0.###}, MinY={3:0.###}, MaxY={4:0.###}, Width={5:0.###}, Height={6:0.###}",
                label,
                bounds.MinX,
                bounds.MaxX,
                bounds.MinY,
                bounds.MaxY,
                bounds.MaxX - bounds.MinX,
                bounds.MaxY - bounds.MinY
            ));
        }

        private static void LogPointsSummary(string label, List<TSG.Point> points, int previewCount = 10) {
            if (points == null) {
                LogDebug(label + ": null");
                return;
            }

            LogDebug(label + " count = " + points.Count);

            for (var i = 0; i < points.Count && i < previewCount; i++) {
                var point = points[i];
                LogDebug(string.Format(
                    "{0}[{1}] = X={2:0.###}, Y={3:0.###}, Z={4:0.###}",
                    label,
                    i,
                    point.X,
                    point.Y,
                    point.Z
                ));
            }

            if (points.Count > previewCount)
                LogDebug(label + ": ...");
        }

        private static void LogValuesSummary(string label, List<double> values) {
            if (values == null) {
                LogDebug(label + ": null");
                return;
            }

            LogDebug(label + " count = " + values.Count + " -> " +
                     string.Join(", ", values.Select(v => v.ToString("0.###"))));
        }

        #endregion

        #region Constants

        private const double OverallDimensionOffsetMillimeters = 20.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;
        private const double CurvedDimensionArcDepthRatio = 0.15;
        private const double DimensionMergeToleranceMillimeters = 1.0;
        private const double MinimumPartSizeForDimensionMillimeters = 10.0;
        private const double SweepStepMillimeters = 1.0;
        private const double SignificantCornerAngleDegrees = 15.0;

        #endregion

        #region Data Structures

        private enum DimensionType {
            Straight,
            Curved
        }

        private enum HorizontalDimensionPlacement {
            Below,
            Above
        }

        private enum VerticalDimensionPlacement {
            Left,
            Right
        }

        private sealed class DimensionOptions {
            public DimensionType DimensionType { get; set; }
            public HorizontalDimensionPlacement HorizontalPlacement { get; set; }
            public VerticalDimensionPlacement VerticalPlacement { get; set; }
            public bool CreateHorizontal { get; set; }
            public bool CreateVertical { get; set; }
        }

        private sealed class DrawingPartWithBounds {
            public TSD.Part DrawingPart { get; set; }
        }

        private sealed class PartBounds {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
        }

        #endregion

        #region Main Dimension Flow

        private static void AddDimensionsFromSelectedView(
            TSD.DrawingHandler drawingHandler,
            TSD.Drawing activeDrawing,
            DimensionOptions options
        ) {
            LogDebug("=== AddDimensionsFromSelectedView ===");

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) {
                LogDebug("selectedView == null");
                return;
            }

            LogDebug("Wybrany widok poprawny.");

            var allViewBounds = GetPartBoundsFromView(selectedView);
            LogDebug("Bounds z widoku: " + allViewBounds.Count);

            var allPartsBounds = GetAssemblyBounds(allViewBounds);
            LogBounds("allPartsBounds(view)", allPartsBounds);

            if (allPartsBounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obwiedni elementów w widoku.");
                return;
            }

            CreateOverallAssemblyDimensions(selectedView, allPartsBounds, allViewBounds, options);
            activeDrawing.CommitChanges();
            LogDebug("CommitChanges wykonany dla AddDimensionsFromSelectedView.");
        }

        private static void AddDimensionsFromSelectedParts(
            TSD.Drawing activeDrawing,
            List<DrawingPartWithBounds> selectedParts,
            DimensionOptions options
        ) {
            LogDebug("=== AddDimensionsFromSelectedParts ===");
            LogDebug("selectedParts.Count = " + selectedParts.Count);

            var selectedView = GetCommonViewFromSelectedParts(selectedParts);
            if (selectedView == null) {
                MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                LogDebug("selectedView == null lub party z różnych widoków.");
                return;
            }

            DrawPartEdgeOutlines(selectedView, selectedParts);

            var allViewBounds = GetPartBoundsFromView(selectedView);
            LogDebug("allViewBounds.Count = " + allViewBounds.Count);

            var allPartsBounds = GetAssemblyBounds(allViewBounds);
            LogBounds("allPartsBounds(selectedView)", allPartsBounds);

            if (allPartsBounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obwiedni wszystkich elementów w widoku.");
                return;
            }

            var selectedBounds = GetPartBoundsFromDrawingParts(selectedView, selectedParts);
            LogDebug("selectedBounds.Count = " + selectedBounds.Count);

            CreateOverallAssemblyDimensions(selectedView, allPartsBounds, selectedBounds, options);
            activeDrawing.CommitChanges();
            LogDebug("CommitChanges wykonany dla AddDimensionsFromSelectedParts.");
        }

        #endregion

        #region Dimension Creation

        private static void CreateOverallAssemblyDimensions(
            TSD.View selectedView,
            PartBounds allPartsBounds,
            List<PartBounds> parts,
            DimensionOptions options
        ) {
            if (selectedView == null || allPartsBounds == null) return;

            LogDebug("=== CreateOverallAssemblyDimensions ===");
            LogBounds("allPartsBounds", allPartsBounds);

            var selectedBounds = GetAssemblyBounds(parts);
            LogBounds("selectedBounds", selectedBounds);

            if (selectedBounds == null) return;

            var xCoordinates = MergeAndSort(
                parts
                    .Where(p => p != null)
                    .SelectMany(p => new[] { p.MinX, p.MaxX })
                    .Concat(new[] { allPartsBounds.MinX, allPartsBounds.MaxX }),
                DimensionMergeToleranceMillimeters
            );

            var yCoordinates = MergeAndSort(
                parts
                    .Where(p => p != null)
                    .SelectMany(p => new[] { p.MinY, p.MaxY })
                    .Concat(new[] { allPartsBounds.MinY, allPartsBounds.MaxY }),
                DimensionMergeToleranceMillimeters
            );

            LogValuesSummary("xCoordinates", xCoordinates);
            LogValuesSummary("yCoordinates", yCoordinates);

            var totalWidth = allPartsBounds.MaxX - allPartsBounds.MinX;
            var totalHeight = allPartsBounds.MaxY - allPartsBounds.MinY;

            LogDebug("totalWidth = " + totalWidth.ToString("0.###"));
            LogDebug("totalHeight = " + totalHeight.ToString("0.###"));

            if (options.CreateHorizontal && totalWidth >= MinimumDimensionSpanMillimeters &&
                xCoordinates.Count >= 2) {
                var anchorY = options.HorizontalPlacement == HorizontalDimensionPlacement.Above
                    ? selectedBounds.MaxY
                    : selectedBounds.MinY;

                LogDebug("Tworzenie wymiaru poziomego, anchorY = " + anchorY.ToString("0.###"));

                var horizontalPoints = new TSD.PointList();
                foreach (var x in xCoordinates)
                    horizontalPoints.Add(new TSG.Point(x, anchorY, 0));

                CreateDimensionSet(
                    selectedView,
                    horizontalPoints,
                    GetHorizontalDirectionVector(options.HorizontalPlacement),
                    OverallDimensionOffsetMillimeters,
                    options.DimensionType
                );
            }

            if (options.CreateVertical && totalHeight >= MinimumDimensionSpanMillimeters &&
                yCoordinates.Count >= 2) {
                var anchorX = options.VerticalPlacement == VerticalDimensionPlacement.Right
                    ? selectedBounds.MaxX
                    : selectedBounds.MinX;

                LogDebug("Tworzenie wymiaru pionowego, anchorX = " + anchorX.ToString("0.###"));

                var verticalPoints = new TSD.PointList();
                foreach (var y in yCoordinates)
                    verticalPoints.Add(new TSG.Point(anchorX, y, 0));

                CreateDimensionSet(
                    selectedView,
                    verticalPoints,
                    GetVerticalDirectionVector(options.VerticalPlacement),
                    OverallDimensionOffsetMillimeters,
                    options.DimensionType
                );
            }
        }

        private static TSG.Vector GetHorizontalDirectionVector(HorizontalDimensionPlacement placement) {
            return placement == HorizontalDimensionPlacement.Above
                ? new TSG.Vector(0.0, 1.0, 0.0)
                : new TSG.Vector(0.0, -1.0, 0.0);
        }

        private static TSG.Vector GetVerticalDirectionVector(VerticalDimensionPlacement placement) {
            return placement == VerticalDimensionPlacement.Right
                ? new TSG.Vector(1.0, 0.0, 0.0)
                : new TSG.Vector(-1.0, 0.0, 0.0);
        }

        private static void CreateDimensionSet(
            TSD.View selectedView,
            TSD.PointList dimensionPoints,
            TSG.Vector directionVector,
            double offsetMillimeters,
            DimensionType dimensionType
        ) {
            LogDebug("=== CreateDimensionSet ===");
            LogDebug("dimensionType = " + dimensionType);
            LogDebug("dimensionPoints.Count = " + dimensionPoints.Count);
            LogDebug(string.Format(
                "directionVector = ({0:0.###}, {1:0.###}, {2:0.###})",
                directionVector.X,
                directionVector.Y,
                directionVector.Z
            ));
            LogDebug("offsetMillimeters = " + offsetMillimeters.ToString("0.###"));

            if (dimensionType == DimensionType.Straight) {
                var straightHandler = new TSD.StraightDimensionSetHandler();
                straightHandler.CreateDimensionSet(
                    selectedView,
                    dimensionPoints,
                    directionVector,
                    offsetMillimeters
                );
                LogDebug("Utworzono StraightDimensionSet.");
                return;
            }

            var arcPoints = ComputeArcPoints(dimensionPoints, directionVector);
            LogPoint("arcStart", arcPoints.Item1);
            LogPoint("arcMid", arcPoints.Item2);
            LogPoint("arcEnd", arcPoints.Item3);

            var curvedHandler = new TSD.CurvedDimensionSetHandler();
            curvedHandler.CreateCurvedDimensionSetOrthogonal(
                selectedView,
                arcPoints.Item1,
                arcPoints.Item2,
                arcPoints.Item3,
                dimensionPoints,
                offsetMillimeters
            );
            LogDebug("Utworzono CurvedDimensionSet.");
        }

        private static Tuple<TSG.Point, TSG.Point, TSG.Point> ComputeArcPoints(
            TSD.PointList dimensionPoints,
            TSG.Vector directionVector
        ) {
            var points = new List<TSG.Point>();
            foreach (var item in dimensionPoints)
                if (item is TSG.Point point)
                    points.Add(point);

            var firstPoint = points.First();
            var lastPoint = points.Last();

            var midX = (firstPoint.X + lastPoint.X) / 2.0;
            var midY = (firstPoint.Y + lastPoint.Y) / 2.0;

            var spanX = lastPoint.X - firstPoint.X;
            var spanY = lastPoint.Y - firstPoint.Y;
            var span = Math.Sqrt(spanX * spanX + spanY * spanY);
            var arcDepth = span * CurvedDimensionArcDepthRatio;

            LogDebug("ComputeArcPoints: span = " + span.ToString("0.###") +
                     ", arcDepth = " + arcDepth.ToString("0.###"));

            return Tuple.Create(
                new TSG.Point(firstPoint.X, firstPoint.Y, 0),
                new TSG.Point(midX + directionVector.X * arcDepth, midY + directionVector.Y * arcDepth, 0),
                new TSG.Point(lastPoint.X, lastPoint.Y, 0)
            );
        }

        #endregion

        #region Part Selection And View Validation

        private static List<DrawingPartWithBounds> GetSelectedDrawingParts(TSD.DrawingHandler drawingHandler) {
            if (drawingHandler == null) return new List<DrawingPartWithBounds>();

            var objectSelector = drawingHandler.GetDrawingObjectSelector();
            if (objectSelector == null) return new List<DrawingPartWithBounds>();

            var selectedObjects = objectSelector.GetSelected();
            if (selectedObjects == null) return new List<DrawingPartWithBounds>();

            var result = new List<DrawingPartWithBounds>();
            selectedObjects.SelectInstances = true;

            while (selectedObjects.MoveNext()) {
                if (!(selectedObjects.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                result.Add(new DrawingPartWithBounds { DrawingPart = drawingPart });
            }

            return result;
        }

        private static TSD.View GetCommonViewFromSelectedParts(List<DrawingPartWithBounds> drawingParts) {
            if (drawingParts == null || drawingParts.Count == 0) return null;

            if (!(drawingParts[0].DrawingPart?.GetView() is TSD.View commonView)) return null;

            foreach (var partWithBounds in drawingParts) {
                if (!(partWithBounds?.DrawingPart?.GetView() is TSD.View currentView)) return null;
                if (!AreViewsEquivalent(commonView, currentView)) return null;
            }

            return commonView;
        }

        private static bool AreViewsEquivalent(TSD.View firstView, TSD.View secondView) {
            if (ReferenceEquals(firstView, secondView)) return true;
            if (firstView == null || secondView == null) return false;

            try {
                var firstCs = firstView.DisplayCoordinateSystem;
                var secondCs = secondView.DisplayCoordinateSystem;

                if (firstCs == null || secondCs == null) return false;

                return
                    Math.Abs(firstCs.Origin.X - secondCs.Origin.X) < 0.001 &&
                    Math.Abs(firstCs.Origin.Y - secondCs.Origin.Y) < 0.001 &&
                    Math.Abs(firstCs.Origin.Z - secondCs.Origin.Z) < 0.001 &&
                    Math.Abs(firstCs.AxisX.X - secondCs.AxisX.X) < 0.001 &&
                    Math.Abs(firstCs.AxisX.Y - secondCs.AxisX.Y) < 0.001 &&
                    Math.Abs(firstCs.AxisX.Z - secondCs.AxisX.Z) < 0.001 &&
                    Math.Abs(firstCs.AxisY.X - secondCs.AxisY.X) < 0.001 &&
                    Math.Abs(firstCs.AxisY.Y - secondCs.AxisY.Y) < 0.001 &&
                    Math.Abs(firstCs.AxisY.Z - secondCs.AxisY.Z) < 0.001;
            }
            catch {
                return false;
            }
        }

        #endregion

        #region Per-Part Outline In View Space

        private static PartBounds GetPartAabbInViewSpace(TSD.View view, TSM.Part modelPart) {
            if (view == null || modelPart == null) return null;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(view.DisplayCoordinateSystem)
                );

                var solid = modelPart.GetSolid();
                if (solid == null) return null;

                var minPoint = solid.MinimumPoint;
                var maxPoint = solid.MaximumPoint;

                if (minPoint == null || maxPoint == null) return null;

                var width = maxPoint.X - minPoint.X;
                var height = maxPoint.Y - minPoint.Y;

                if (width < MinimumPartSizeForDimensionMillimeters &&
                    height < MinimumPartSizeForDimensionMillimeters) return null;

                return new PartBounds {
                    MinX = minPoint.X,
                    MaxX = maxPoint.X,
                    MinY = minPoint.Y,
                    MaxY = maxPoint.Y
                };
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static bool ShouldUseSweepOutline(TSM.Part modelPart) {
            return modelPart is TSM.LoftedPlate || modelPart is TSM.ContourPlate;
        }

        private static List<TSG.Point> GetSweptPlateOutlinePoints(TSM.Solid solid) {
            var points = new List<TSG.Point>();
            if (solid == null) return points;

            var minX = solid.MinimumPoint.X;
            var maxX = solid.MaximumPoint.X;

            LogDebug("GetSweptPlateOutlinePoints");
            LogDebug("Plate minX = " + minX.ToString("0.###") + ", maxX = " + maxX.ToString("0.###"));

            if (maxX - minX < 1e-6) return points;

            var xValues = new List<double> { minX };
            for (var x = minX + SweepStepMillimeters; x < maxX; x += SweepStepMillimeters)
                xValues.Add(x);
            xValues.Add(maxX);

            LogValuesSummary("Sweep xValues", xValues);

            foreach (var xValue in xValues) {
                var enumerator = solid.GetAllIntersectionPoints(
                    new TSG.Point(xValue, 0, 0),
                    new TSG.Point(xValue, 1, 0),
                    new TSG.Point(xValue, 0, 1)
                );

                var localCount = 0;

                while (enumerator.MoveNext()) {
                    if (!(enumerator.Current is TSG.Point p)) continue;
                    points.Add(new TSG.Point(xValue, p.Y, 0));
                    localCount++;
                }

                LogDebug("x = " + xValue.ToString("0.###") + " -> intersections = " + localCount);
            }

            points = RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
            LogPointsSummary("Swept raw points", points);
            return points;
        }

        private static List<TSG.Point> SimplifyPolyline(
            List<TSG.Point> points,
            double toleranceMillimeters = 0.5
        ) {
            if (points == null || points.Count < 3) return points;

            var current = new List<TSG.Point>(points);
            var changed = true;

            while (changed && current.Count >= 3) {
                changed = false;
                var next = new List<TSG.Point>();

                var index = 0;
                while (index < current.Count) {
                    var remaining = current.Count - index;

                    if (remaining >= 3) {
                        var first = current[index];
                        var middle = current[index + 1];
                        var last = current[index + 2];

                        var distance = GetDistanceToSegment(middle, first, last);

                        if (distance <= toleranceMillimeters) {
                            if (next.Count == 0 || !ArePointsEqual(next[next.Count - 1], first))
                                next.Add(first);

                            next.Add(last);
                            changed = true;
                        }
                        else {
                            if (next.Count == 0 || !ArePointsEqual(next[next.Count - 1], first))
                                next.Add(first);

                            next.Add(middle);
                            next.Add(last);
                        }

                        index += 3;
                    }
                    else {
                        for (var i = index; i < current.Count; i++)
                            if (next.Count == 0 || !ArePointsEqual(next[next.Count - 1], current[i]))
                                next.Add(current[i]);

                        break;
                    }
                }

                current = RemoveNearDuplicates(next, DuplicateToleranceMillimeters);
            }

            LogDebug("SimplifyPolyline: before = " + points.Count + ", after = " + current.Count);
            return current;
        }

        private static double GetDistanceToSegment(TSG.Point point, TSG.Point start, TSG.Point end) {
            var segmentX = end.X - start.X;
            var segmentY = end.Y - start.Y;

            var segmentLengthSquared = segmentX * segmentX + segmentY * segmentY;

            if (segmentLengthSquared < 1e-12) {
                var distanceX = point.X - start.X;
                var distanceY = point.Y - start.Y;
                return Math.Sqrt(distanceX * distanceX + distanceY * distanceY);
            }

            var t = (
                (point.X - start.X) * segmentX +
                (point.Y - start.Y) * segmentY
            ) / segmentLengthSquared;

            if (t < 0.0) t = 0.0;
            if (t > 1.0) t = 1.0;

            var projectedX = start.X + t * segmentX;
            var projectedY = start.Y + t * segmentY;

            var dx = point.X - projectedX;
            var dy = point.Y - projectedY;

            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static bool ArePointsEqual(TSG.Point first, TSG.Point second) {
            if (first == null || second == null) return false;

            return
                Math.Abs(first.X - second.X) <= DuplicateToleranceMillimeters &&
                Math.Abs(first.Y - second.Y) <= DuplicateToleranceMillimeters;
        }

        private static List<TSG.Point> BuildUpperLowerChainOutline(List<TSG.Point> points) {
            if (points == null || points.Count < 2) return null;

            var byX = points
                .GroupBy(p => p.X)
                .OrderBy(g => g.Key)
                .ToList();

            if (byX.Count < 2) return null;

            var upper = byX.Select(g => new TSG.Point(g.Key, g.Max(p => p.Y), 0)).ToList();
            var lower = byX.Select(g => new TSG.Point(g.Key, g.Min(p => p.Y), 0)).ToList();

            var outline = new List<TSG.Point>();
            outline.AddRange(upper);
            outline.AddRange(Enumerable.Reverse(lower));

            LogPointsSummary("Upper chain", upper);
            LogPointsSummary("Lower chain", lower);
            LogPointsSummary("Outline before simplify", outline);

            return outline;
        }

        #endregion

        #region Part Edge Outline Drawing

        private static void DrawPartEdgeOutlines(TSD.View view, List<DrawingPartWithBounds> drawingParts) {
            if (view == null || drawingParts == null) return;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(view.DisplayCoordinateSystem)
                );

                foreach (var dp in drawingParts.Where(p => p?.DrawingPart != null)) {
                    if (!(MyModel.SelectModelObject(dp.DrawingPart.ModelIdentifier) is TSM.Part modelPart)) continue;

                    modelPart.Select();

                    var solid = modelPart.GetSolid();
                    if (solid == null) continue;

                    LogDebug(new string('-', 80));
                    LogDebug("Part: " + modelPart.Name + " | Type: " + modelPart.GetType().FullName);
                    LogPoint("solid.MinimumPoint", solid.MinimumPoint);
                    LogPoint("solid.MaximumPoint", solid.MaximumPoint);

                    List<TSG.Point> outline;

                    if (ShouldUseSweepOutline(modelPart)) {
                        LogDebug("Tryb: sweepowany plate");

                        var points = GetSweptPlateOutlinePoints(solid);
                        if (points.Count < 2) {
                            LogDebug("Za mało punktów sweep.");
                            continue;
                        }

                        outline = BuildUpperLowerChainOutline(points);
                        if (outline == null || outline.Count < 3) {
                            LogDebug("Outline sweep == null lub < 3");
                            continue;
                        }

                        outline = PrepareSegmentPath(outline);
                        outline = EnsureClosedOutline(outline);

                        if (outline == null || outline.Count < 4) {
                            LogDebug("Outline sweep po domknięciu == null lub < 4");
                            continue;
                        }

                        LogPointsSummary("Sweep final outline", outline);
                    }
                    else {
                        LogDebug("Tryb: zwykły Part - tylko zielony przekrój");

                        var backZ = solid.MinimumPoint.Z;
                        const double inward = 1.0;

                        if (solid.MaximumPoint.Z - solid.MinimumPoint.Z > inward * 2)
                            backZ += inward;
                        else
                            backZ = (solid.MinimumPoint.Z + solid.MaximumPoint.Z) * 0.5;

                        LogDebug("backZ = " + backZ.ToString("0.###"));

                        var backPoints = GetIntersectionPointsAtLocalZ(solid, backZ);
                        LogPointsSummary("backPoints", backPoints);

                        if (backPoints == null || backPoints.Count < 3) {
                            LogDebug("Za mało punktów backPoints.");
                            continue;
                        }

                        backPoints = RemoveNearDuplicates(backPoints, DuplicateToleranceMillimeters);
                        outline = BuildConvexHull2D(backPoints);

                        if (outline == null || outline.Count < 3) {
                            LogDebug("Hull backPoints == null lub < 3");
                            continue;
                        }

                        outline = PrepareSegmentPath(outline);
                        outline = EnsureClosedOutline(outline);

                        if (outline == null || outline.Count < 4) {
                            LogDebug("Outline back po domknięciu == null lub < 4");
                            continue;
                        }

                        LogPointsSummary("Back final outline", outline);
                    }

                    DrawOutlineByAngleType(view, outline);
                }
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static void DrawOutlineByAngleType(TSD.View view, List<TSG.Point> outline) {
            if (view == null || outline == null || outline.Count < 2) {
                LogDebug("DrawOutlineByAngleType: outline == null lub < 2");
                return;
            }

            var vertices = GetOpenOutlineVertices(outline);
            if (vertices == null || vertices.Count < 2) {
                LogDebug("DrawOutlineByAngleType: vertices == null lub < 2");
                return;
            }

            LogPointsSummary("Open outline vertices", vertices);

            if (vertices.Count == 2) {
                LogDebug("DrawOutlineByAngleType: outline ma tylko 2 punkty -> rysuję linię.");
                DrawStraightSegmentPrimitive(view, vertices[0], vertices[1], TSD.DrawingColors.Green);
                return;
            }

            var significantCornerIndices = GetSignificantCornerIndices(vertices);
            LogDebug("Significant corners = " + significantCornerIndices.Count);

            if (significantCornerIndices.Count == 0) {
                DrawPolylinePrimitive(view, vertices, TSD.DrawingColors.Green);
                return;
            }

            if (significantCornerIndices.Count == 1) {
                DrawPolylinePrimitive(view, vertices, TSD.DrawingColors.Green);
                return;
            }

            for (var i = 0; i < significantCornerIndices.Count; i++) {
                var startIndex = significantCornerIndices[i];
                var endIndex = significantCornerIndices[(i + 1) % significantCornerIndices.Count];

                var path = CollectPathBetweenCorners(vertices, startIndex, endIndex);
                path = PrepareSegmentPath(path);

                if (path == null || path.Count < 2) continue;

                LogPointsSummary("Path " + i, path);

                if (path.Count == 2)
                    DrawStraightSegmentPrimitive(view, path[0], path[1], TSD.DrawingColors.Red);
                else
                    DrawPolylinePrimitive(view, path, TSD.DrawingColors.Green);
            }
        }

        private static void DrawStraightSegmentPrimitive(
            TSD.View view,
            TSG.Point startPoint,
            TSG.Point endPoint,
            TSD.DrawingColors color
        ) {
            if (view == null || startPoint == null || endPoint == null) return;

            var distance = Math.Sqrt(
                (startPoint.X - endPoint.X) * (startPoint.X - endPoint.X) +
                (startPoint.Y - endPoint.Y) * (startPoint.Y - endPoint.Y)
            );

            if (distance <= DuplicateToleranceMillimeters) {
                LogDebug("DrawStraightSegmentPrimitive: odcinek za krótki.");
                return;
            }

            var points = new TSD.PointList {
                new TSG.Point(startPoint.X, startPoint.Y, 0),
                new TSG.Point(endPoint.X, endPoint.Y, 0)
            };

            LogPoint("Red start", startPoint);
            LogPoint("Red end", endPoint);

            var polyline = new TSD.Polyline(view, points);
            polyline.Attributes.Line.Color = color;
            polyline.Insert();
        }

        private static void DrawPolylinePrimitive(
            TSD.View view,
            List<TSG.Point> points,
            TSD.DrawingColors color
        ) {
            if (view == null || points == null || points.Count < 2) {
                LogDebug("DrawPolylinePrimitive: points == null lub < 2");
                return;
            }

            var cleanedPoints = RemoveNearDuplicates(
                new List<TSG.Point>(points),
                DuplicateToleranceMillimeters
            );

            if (cleanedPoints == null || cleanedPoints.Count < 2) {
                LogDebug("DrawPolylinePrimitive: cleanedPoints == null lub < 2");
                return;
            }

            if (cleanedPoints.Count == 2) {
                LogDebug("DrawPolylinePrimitive: wykryto 2 punkty -> rysuję zwykłą linię.");
                DrawStraightSegmentPrimitive(view, cleanedPoints[0], cleanedPoints[1], color);
                return;
            }

            if (cleanedPoints.Count == 3) {
                var first = cleanedPoints[0];
                var last = cleanedPoints[cleanedPoints.Count - 1];

                var distance = Math.Sqrt(
                    (first.X - last.X) * (first.X - last.X) +
                    (first.Y - last.Y) * (first.Y - last.Y)
                );

                if (distance <= DuplicateToleranceMillimeters) {
                    LogDebug("DrawPolylinePrimitive: po domknięciu są tylko 2 unikalne punkty -> rysuję zwykłą linię.");
                    DrawStraightSegmentPrimitive(view, cleanedPoints[0], cleanedPoints[1], color);
                    return;
                }
            }

            var polylinePoints = new TSD.PointList();
            foreach (var point in cleanedPoints)
                polylinePoints.Add(new TSG.Point(point.X, point.Y, 0));

            LogPoint("Polyline start", cleanedPoints.First());
            LogPoint("Polyline end", cleanedPoints.Last());
            LogDebug("Rysuję polyline, points = " + cleanedPoints.Count + ", color = " + color);

            var polyline = new TSD.Polyline(view, polylinePoints);
            polyline.Attributes.Line.Color = color;
            polyline.Insert();
        }

        #endregion

        #region Part Collection — Fast AABB Bounds

        private static List<PartBounds> GetPartBoundsFromView(TSD.View view) {
            var result = new List<PartBounds>();
            if (view == null) return result;

            var addedIds = new HashSet<int>();
            var drawingElements = view.GetAllObjects(typeof(TSD.ModelObject));
            if (drawingElements == null) return result;

            while (drawingElements.MoveNext()) {
                if (!(drawingElements.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                if (!(MyModel.SelectModelObject(drawingPart.ModelIdentifier) is TSM.Part modelPart)) continue;

                var id = modelPart.Identifier.ID;
                if (!addedIds.Add(id)) continue;

                modelPart.Select();

                var bounds = GetPartAabbInViewSpace(view, modelPart);
                if (bounds == null) continue;

                LogDebug("GetPartBoundsFromView -> " + modelPart.Name + " ID=" + id);
                LogBounds("view part bounds", bounds);

                result.Add(bounds);
            }

            return result;
        }

        private static List<PartBounds> GetPartBoundsFromDrawingParts(
            TSD.View view,
            List<DrawingPartWithBounds> drawingParts
        ) {
            var result = new List<PartBounds>();
            if (view == null || drawingParts == null) return result;

            foreach (var dp in drawingParts.Where(dp => dp?.DrawingPart != null)) {
                if (!(MyModel.SelectModelObject(dp.DrawingPart.ModelIdentifier) is TSM.Part modelPart)) continue;

                modelPart.Select();

                var bounds = GetPartAabbInViewSpace(view, modelPart);
                if (bounds == null) continue;

                LogDebug("GetPartBoundsFromDrawingParts -> " + modelPart.Name + " ID=" + modelPart.Identifier.ID);
                LogBounds("selected part bounds", bounds);

                result.Add(bounds);
            }

            return result;
        }

        #endregion
    }
}