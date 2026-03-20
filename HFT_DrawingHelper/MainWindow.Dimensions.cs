using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Entry Point

        private static void AddDimensions(DimensionOptions options) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var activeDrawing = drawingHandler.GetActiveDrawing();
            if (activeDrawing == null) return;

            var selectedParts = GetSelectedDrawingParts(drawingHandler);

            if (selectedParts == null || selectedParts.Count == 0) {
                AddDimensionsFromSelectedView(drawingHandler, activeDrawing, options);
                return;
            }

            AddDimensionsFromSelectedParts(activeDrawing, selectedParts, options);
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

        #endregion

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

            return !found ? null : new PartBounds { MinX = minX, MaxX = maxX, MinY = minY, MaxY = maxY };
        }

        #endregion

        #region Constants

        private const double OverallDimensionOffsetMillimeters = 20.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;
        private const double CurvedDimensionArcDepthRatio = 0.15;
        private const double DimensionMergeToleranceMillimeters = 1.0;
        private const double MinimumPartSizeForDimensionMillimeters = 20.0;

        // Krok X-sweep dla LoftedPlate — co ile mm wzdłuż długości próbkujemy.
        // 20mm daje wystarczającą dokładność nawet dla gwałtownych zmian krzywizny.
        private const double LoftedPlateXStepMillimeters = 20.0;

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
            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var allViewBounds = GetPartBoundsFromView(selectedView);
            var allPartsBounds = GetAssemblyBounds(allViewBounds);
            if (allPartsBounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obwiedni elementów w widoku.");
                return;
            }

            CreateOverallAssemblyDimensions(selectedView, allPartsBounds, allViewBounds, options);
            activeDrawing.CommitChanges();
        }

        private static void AddDimensionsFromSelectedParts(
            TSD.Drawing activeDrawing,
            List<DrawingPartWithBounds> selectedParts,
            DimensionOptions options
        ) {
            var selectedView = GetCommonViewFromSelectedParts(selectedParts);
            if (selectedView == null) {
                MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                return;
            }

            DrawPartEdgeOutlines(selectedView, selectedParts);

            var allViewBounds = GetPartBoundsFromView(selectedView);
            var allPartsBounds = GetAssemblyBounds(allViewBounds);
            if (allPartsBounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obwiedni wszystkich elementów w widoku.");
                return;
            }

            var selectedBounds = GetPartBoundsFromDrawingParts(selectedView, selectedParts);
            CreateOverallAssemblyDimensions(selectedView, allPartsBounds, selectedBounds, options);
            activeDrawing.CommitChanges();
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

            var selectedBounds = GetAssemblyBounds(parts);
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

            var totalWidth = allPartsBounds.MaxX - allPartsBounds.MinX;
            var totalHeight = allPartsBounds.MaxY - allPartsBounds.MinY;

            if (options.CreateHorizontal && totalWidth >= MinimumDimensionSpanMillimeters &&
                xCoordinates.Count >= 2) {
                var anchorY = options.HorizontalPlacement == HorizontalDimensionPlacement.Above
                    ? selectedBounds.MaxY
                    : selectedBounds.MinY;

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
            if (dimensionType == DimensionType.Straight) {
                var straightHandler = new TSD.StraightDimensionSetHandler();
                straightHandler.CreateDimensionSet(
                    selectedView,
                    dimensionPoints,
                    directionVector,
                    offsetMillimeters
                );
                return;
            }

            var arcPoints = ComputeArcPoints(dimensionPoints, directionVector);
            var curvedHandler = new TSD.CurvedDimensionSetHandler();
            curvedHandler.CreateCurvedDimensionSetOrthogonal(
                selectedView,
                arcPoints.Item1,
                arcPoints.Item2,
                arcPoints.Item3,
                dimensionPoints,
                offsetMillimeters
            );
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

        /// <summary>
        ///     LoftedPlate: X-sweep — pionowe płaszczyzny przy różnych X,
        ///     zbiera min/max Y z punktów przecięcia → górny+dolny łańcuch.
        ///     To jedyna metoda która zachowuje wklęsły profil (np. "dip" belki mostowej),
        ///     bo próbkuje WZDŁUŻ długości płyty a nie przez jej grubość.
        /// </summary>
        private static List<TSG.Point> GetLoftedPlateOutlinePoints(TSM.Solid solid) {
            var points = new List<TSG.Point>();
            if (solid == null) return points;

            var minX = solid.MinimumPoint.X;
            var maxX = solid.MaximumPoint.X;

            if (maxX - minX < 1e-6) return points;

            // Zawsze próbkuj minX i maxX, plus równomierne kroki pomiędzy
            var xValues = new List<double> { minX };
            for (var x = minX + LoftedPlateXStepMillimeters; x < maxX; x += LoftedPlateXStepMillimeters)
                xValues.Add(x);
            xValues.Add(maxX);

            foreach (var xValue in xValues) {
                // Pionowa płaszczyzna przy X = xValue: normalna w kierunku X
                // Trzy punkty definiujące płaszczyznę YZ przy danym X
                var enumerator = solid.GetAllIntersectionPoints(
                    new TSG.Point(xValue, 0, 0),
                    new TSG.Point(xValue, 1, 0),
                    new TSG.Point(xValue, 0, 1)
                );

                while (enumerator.MoveNext()) {
                    if (!(enumerator.Current is TSG.Point p)) continue;
                    // Zachowaj X z próbki (nie z punktu przecięcia który może się lekko różnić)
                    points.Add(new TSG.Point(xValue, p.Y, 0));
                }
            }

            return points;
        }

        /// <summary>
        ///     Kroi solid przy maxZ — przedniej ścianie solida w view CS.
        ///     W Tekla view CS oś Z rośnie w stronę obserwatora, więc maxZ
        ///     to ściana najbliżej obserwatora — dokładnie to co widać na widoku.
        ///     Fallback: jeśli maxZ nie trafia w solid, próbuje maxZ - mały offset.
        /// </summary>
        private static List<TSG.Point> GetSolidOutlinePointsByZSweep(TSM.Solid solid) {
            if (solid == null) return new List<TSG.Point>();

            var maxZ = solid.MaximumPoint.Z;

            var points = GetIntersectionPointsAtLocalZ(solid, maxZ);

            if (points.Count == 0)
                points = GetIntersectionPointsAtLocalZ(solid, maxZ - 1.0);

            return RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
        }

        /// <summary>
        ///     Usuwa punkty koliniarne z polilinii — jeśli kąt między dwoma sąsiednimi
        ///     segmentami jest mniejszy niż tolerancja (punkty leżą na jednej prostej),
        ///     środkowy punkt jest zbędny i zostaje usunięty.
        ///     Działa zarówno dla otwartych jak i zamkniętych polilinii.
        /// </summary>
        private static List<TSG.Point> SimplifyPolyline(
            List<TSG.Point> points,
            double collinearAngleToleranceDegrees = 0.5
        ) {
            if (points == null || points.Count < 3) return points;

            var result = new List<TSG.Point> { points[0] };

            for (var i = 1; i < points.Count - 1; i++) {
                var prev = result[result.Count - 1];
                var curr = points[i];
                var next = points[i + 1];

                var dx1 = curr.X - prev.X;
                var dy1 = curr.Y - prev.Y;
                var dx2 = next.X - curr.X;
                var dy2 = next.Y - curr.Y;

                var len1 = Math.Sqrt(dx1 * dx1 + dy1 * dy1);
                var len2 = Math.Sqrt(dx2 * dx2 + dy2 * dy2);

                if (len1 < 1e-9 || len2 < 1e-9) continue;

                var dot = (dx1 * dx2 + dy1 * dy2) / (len1 * len2);
                if (dot > 1.0) dot = 1.0;
                if (dot < -1.0) dot = -1.0;

                var angleDegrees = Math.Acos(dot) * 180.0 / Math.PI;

                // Kąt bliski 0° = punkty koliniarne = środkowy punkt zbędny
                if (angleDegrees > collinearAngleToleranceDegrees)
                    result.Add(curr);
            }

            result.Add(points[points.Count - 1]);
            return result;
        }

        /// <summary>
        ///     Buduje zamkniętą polilinię obrysu z listy punktów metodą górny/dolny łańcuch.
        ///     Punkty muszą być już posortowane po X (jak z X-sweep).
        ///     Nie używa smoothingu — zachowuje kształt wklęsły.
        /// </summary>
        private static List<TSG.Point> BuildUpperLowerChainOutline(List<TSG.Point> points) {
            if (points == null || points.Count < 2) return null;

            // Grupuj po X, zbierz min/max Y
            var byX = points
                .GroupBy(p => p.X)
                .OrderBy(g => g.Key)
                .ToList();

            if (byX.Count < 2) return null;

            var upper = byX.Select(g => new TSG.Point(g.Key, g.Max(p => p.Y), 0)).ToList();
            var lower = byX.Select(g => new TSG.Point(g.Key, g.Min(p => p.Y), 0)).ToList();

            // Górny łańcuch lewo→prawo + dolny prawo→lewo = zamknięty obrys
            var outline = new List<TSG.Point>();
            outline.AddRange(upper);
            outline.AddRange(Enumerable.Reverse(lower));
            outline.Add(new TSG.Point(upper[0].X, upper[0].Y, 0)); // zamknięcie

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

                    if (modelPart is TSM.LoftedPlate) {
                        var pts = GetLoftedPlateOutlinePoints(solid);
                        if (pts.Count < 2) continue;
                        var outline = BuildUpperLowerChainOutline(pts);
                        if (outline == null || outline.Count < 3) continue;
                        outline = SimplifyPolyline(outline);
                        if (outline == null || outline.Count < 3) continue;
                        DrawOutlineAsPolyline(view, outline, TSD.DrawingColors.Green);
                    }
                    else {
                        // WP = DisplayCS → Z = głębia widoku, X,Y = współrzędne papieru.
                        // Nie ma potrzeby żadnej transformacji — punkty przecięcia są
                        // bezpośrednio w układzie rysunkowym.
                        var frontZ = solid.MaximumPoint.Z;
                        var backZ = solid.MinimumPoint.Z;
                        const double inward = 1.0;

                        if (frontZ - backZ > inward * 2) {
                            frontZ -= inward;
                            backZ += inward;
                        }
                        else {
                            var centerZ = (frontZ + backZ) * 0.5;
                            frontZ = centerZ;
                            backZ = centerZ;
                        }

                        DrawHullOutline(view, GetIntersectionPointsAtLocalZ(solid, frontZ), TSD.DrawingColors.Red);
                        DrawHullOutline(view, GetIntersectionPointsAtLocalZ(solid, backZ), TSD.DrawingColors.Green);
                    }

                    var label = !string.IsNullOrWhiteSpace(modelPart.Name)
                        ? modelPart.Name
                        : modelPart.GetType().Name;

                    new TSD.Text(
                        view,
                        new TSG.Point(
                            (solid.MinimumPoint.X + solid.MaximumPoint.X) / 2.0,
                            (solid.MinimumPoint.Y + solid.MaximumPoint.Y) / 2.0,
                            0
                        ),
                        label
                    ).Insert();
                }
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static void DrawHullOutline(
            TSD.View view,
            List<TSG.Point> points,
            TSD.DrawingColors color
        ) {
            if (points == null || points.Count < 3) return;

            points = RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
            var hull = BuildConvexHull2D(points);
            if (hull.Count < 3) return;

            var outline = new List<TSG.Point>(hull) {
                new TSG.Point(hull[0].X, hull[0].Y, 0)
            };
            outline = SimplifyPolyline(outline);
            if (outline == null || outline.Count < 3) return;

            DrawOutlineAsPolyline(view, outline, color);
        }

        private static void DrawOutlineAsPolyline(
            TSD.View view,
            List<TSG.Point> outline,
            TSD.DrawingColors color
        ) {
            if (outline == null || outline.Count < 2) return;

            var polylinePoints = new TSD.PointList();
            foreach (var pt in outline)
                polylinePoints.Add(new TSG.Point(pt.X, pt.Y, 0));

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

                result.Add(bounds);
            }

            return result;
        }

        #endregion
    }
}