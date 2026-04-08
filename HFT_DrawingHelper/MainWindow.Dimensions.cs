using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Main Dimension Flow

        private static DimensionGeometrySnapshot BuildDimensionGeometrySnapshot(TSD.DrawingHandler drawingHandler) {
            var selectedParts = GetSelectedDrawingParts(drawingHandler);

            if (selectedParts.Count == 0) {
                var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
                if (selectedView == null) return null;

                var allViewBounds = GetPartBoundsFromView(selectedView);
                var allPartsBounds = GetAssemblyBounds(allViewBounds);

                if (allPartsBounds == null) return null;

                var endpoints = GetLineAndPolylineEndpoints(selectedView);

                return new DimensionGeometrySnapshot {
                    View = selectedView,
                    AllPartsBounds = allPartsBounds,
                    SelectedBounds = allPartsBounds,
                    Endpoints = endpoints,
                    ElementEndpoints = new List<ElementEndpointsSnapshot> {
                        new ElementEndpointsSnapshot {
                            Bounds = allPartsBounds,
                            Endpoints = endpoints
                        }
                    }
                };
            }

            var selectedViewFromParts = GetCommonViewFromSelectedParts(selectedParts);
            if (selectedViewFromParts == null) {
                MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                return null;
            }

            var elementEndpoints = DrawPartEdgeOutlines(selectedViewFromParts, selectedParts);

            var allViewBoundsFromParts = GetPartBoundsFromView(selectedViewFromParts);
            var allPartsBoundsFromParts = GetAssemblyBounds(allViewBoundsFromParts);

            if (allPartsBoundsFromParts == null) return null;

            var selectedBoundsList = GetPartBoundsFromDrawingParts(selectedViewFromParts, selectedParts);
            var selectedBounds = GetAssemblyBounds(selectedBoundsList);

            if (selectedBounds == null) return null;

            var mergedEndpoints = elementEndpoints
                .Where(x => x?.Endpoints != null)
                .SelectMany(x => x.Endpoints)
                .ToList();

            return new DimensionGeometrySnapshot {
                View = selectedViewFromParts,
                AllPartsBounds = allPartsBoundsFromParts,
                SelectedBounds = selectedBounds,
                Endpoints = mergedEndpoints,
                ElementEndpoints = elementEndpoints
            };
        }

        #endregion

        #region Entry Point

        private static void AddDimensions(List<DimensionOptions> optionsList) {
            try {
                var validOptions = optionsList?
                    .Where(option =>
                        option != null &&
                        (
                            (option.Axis == DimensionAxis.Horizontal &&
                             (option.Placement == DimensionPlacement.Above ||
                              option.Placement == DimensionPlacement.Below)) ||
                            (option.Axis == DimensionAxis.Vertical &&
                             (option.Placement == DimensionPlacement.Left ||
                              option.Placement == DimensionPlacement.Right))
                        ))
                    .ToList();

                if (validOptions == null || validOptions.Count == 0)
                    return;

                var drawingHandler = new TSD.DrawingHandler();
                if (!drawingHandler.GetConnectionStatus())
                    return;

                var activeDrawing = drawingHandler.GetActiveDrawing();
                if (activeDrawing == null)
                    return;

                if (validOptions.Where(o => o.DimensionType == DimensionType.Curved)
                    .Any(option => !PickCurvedArcPoints(drawingHandler, option)))
                    return;

                var snapshot = BuildDimensionGeometrySnapshot(drawingHandler);
                if (snapshot?.View == null || snapshot.AllPartsBounds == null || snapshot.SelectedBounds == null)
                    return;

                foreach (var option in validOptions) {
                    if (option.DimensionType == DimensionType.Curved) {
                        if (option.IsOverall) {
                            var curvedPoints = GetCurvedDimensionReferencePoints(
                                snapshot,
                                option,
                                option.IsHorizontal
                                    ? CurvedSideSelectionMode.Horizontal
                                    : CurvedSideSelectionMode.Vertical
                            );

                            if (curvedPoints.Count >= 2)
                                CreateCurvedDimensionSet(
                                    snapshot.View,
                                    curvedPoints,
                                    option,
                                    FarDimensionOffsetMillimeters
                                );
                        }
                        else
                            foreach (var curvedPoints in snapshot.ElementEndpoints
                                         .Select(element => GetCurvedDimensionReferencePoints(
                                             snapshot,
                                             option,
                                             option.IsHorizontal
                                                 ? CurvedSideSelectionMode.Horizontal
                                                 : CurvedSideSelectionMode.Vertical,
                                             element
                                         ))
                                         .Where(curvedPoints => curvedPoints.Count >= 2))
                                CreateCurvedDimensionSet(
                                    snapshot.View,
                                    curvedPoints,
                                    option,
                                    PerElementDimensionOffsetMillimeters
                                );

                        continue;
                    }

                    CreateStraightDimensions(snapshot, option);
                }

                activeDrawing.CommitChanges();
            }
            catch {
                // ignored
            }
        }

        #endregion

        #region Bounds Calculation

        private static PartBounds GetAssemblyBounds(List<PartBounds> parts) {
            var valid = parts?.Where(p => p != null).ToList();
            if (valid == null || valid.Count == 0) return null;

            return new PartBounds {
                MinX = valid.Min(p => p.MinX),
                MaxX = valid.Max(p => p.MaxX),
                MinY = valid.Min(p => p.MinY),
                MaxY = valid.Max(p => p.MaxY)
            };
        }

        private static PartBounds GetBoundsFromPoints(List<TSG.Point> points) {
            var valid = points?.Where(point => point != null).ToList();
            if (valid == null || valid.Count == 0) return null;

            return new PartBounds {
                MinX = valid.Min(point => point.X),
                MaxX = valid.Max(point => point.X),
                MinY = valid.Min(point => point.Y),
                MaxY = valid.Max(point => point.Y)
            };
        }

        #endregion

        #region Geometry Helpers

        private enum CurvedSideSelectionMode {
            Horizontal,
            Vertical
        }

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

        private static List<TSG.Point> GetLineAndPolylineEndpoints(TSD.View view) {
            var result = new List<TSG.Point>();
            if (view == null) return result;

            var it = view.GetAllObjects(typeof(TSD.Line));
            while (it?.MoveNext() == true)
                if (it.Current is TSD.Line line) {
                    result.Add(new TSG.Point(line.StartPoint.X, line.StartPoint.Y, 0));
                    result.Add(new TSG.Point(line.EndPoint.X, line.EndPoint.Y, 0));
                }

            it = view.GetAllObjects(typeof(TSD.Polyline));
            while (it?.MoveNext() == true) {
                if (!(it.Current is TSD.Polyline poly)) continue;

                var pts = new List<TSG.Point>();
                foreach (var item in poly.Points)
                    if (item is TSG.Point point)
                        pts.Add(point);

                if (pts.Count == 0) continue;

                result.Add(new TSG.Point(pts[0].X, pts[0].Y, 0));
                if (pts.Count > 1)
                    result.Add(new TSG.Point(pts[pts.Count - 1].X, pts[pts.Count - 1].Y, 0));
            }

            return result;
        }

        private static List<double> GetPerElementHorizontalCoordinates(
            ElementEndpointsSnapshot element,
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            if (element?.Endpoints == null || element.Endpoints.Count == 0 || options == null)
                return new List<double>();

            var localBounds = GetBoundsFromPoints(element.Endpoints);
            if (localBounds == null || snapshot?.AllPartsBounds == null)
                return new List<double>();

            var midY = (localBounds.MinY + localBounds.MaxY) / 2.0;

            return MergeAndSort(
                element.Endpoints
                    .Where(point => options.IsAbove ? point.Y >= midY : point.Y <= midY)
                    .Select(point => point.X)
                    .Concat(new[] { snapshot.AllPartsBounds.MinX, snapshot.AllPartsBounds.MaxX }),
                DimensionMergeToleranceMillimeters
            );
        }

        private static List<double> GetPerElementVerticalCoordinates(
            ElementEndpointsSnapshot element,
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            if (element?.Endpoints == null || element.Endpoints.Count == 0 || options == null)
                return new List<double>();

            var localBounds = GetBoundsFromPoints(element.Endpoints);
            if (localBounds == null || snapshot?.AllPartsBounds == null)
                return new List<double>();

            var midX = (localBounds.MinX + localBounds.MaxX) / 2.0;

            return MergeAndSort(
                element.Endpoints
                    .Where(point => options.IsRight ? point.X >= midX : point.X <= midX)
                    .Select(point => point.Y)
                    .Concat(new[] { snapshot.AllPartsBounds.MinY, snapshot.AllPartsBounds.MaxY }),
                DimensionMergeToleranceMillimeters
            );
        }

        private static List<double> GetCoordinatesForElement(
            ElementEndpointsSnapshot element,
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            if (element?.Endpoints == null || element.Endpoints.Count == 0 || options == null)
                return new List<double>();

            return options.IsHorizontal
                ? GetPerElementHorizontalCoordinates(element, snapshot, options)
                : GetPerElementVerticalCoordinates(element, snapshot, options);
        }

        private static PartBounds GetElementLocalBounds(ElementEndpointsSnapshot element) {
            return GetBoundsFromPoints(element?.Endpoints);
        }

        private static void AddHorizontalPoints(TSD.PointList points, IEnumerable<double> xCoordinates,
            double anchorY) {
            foreach (var x in xCoordinates)
                points.Add(new TSG.Point(x, anchorY, 0));
        }

        private static void AddVerticalPoints(TSD.PointList points, IEnumerable<double> yCoordinates, double anchorX) {
            foreach (var y in yCoordinates)
                points.Add(new TSG.Point(anchorX, y, 0));
        }

        private static List<TSG.Point> RemoveDuplicatePointsByDistance(
            IEnumerable<TSG.Point> points,
            double tolerance
        ) {
            var result = new List<TSG.Point>();
            if (points == null) return result;

            foreach (var point in points.Where(p => p != null)) {
                var exists = result.Any(existing =>
                    Math.Abs(existing.X - point.X) <= tolerance &&
                    Math.Abs(existing.Y - point.Y) <= tolerance
                );

                if (!exists)
                    result.Add(new TSG.Point(point.X, point.Y, 0));
            }

            return result;
        }

        private static double GetPointParameterOnSegment(TSG.Point point, TSG.Point start, TSG.Point end) {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var lengthSquared = dx * dx + dy * dy;

            if (lengthSquared < 1e-9) return 0.0;

            return ((point.X - start.X) * dx + (point.Y - start.Y) * dy) / lengthSquared;
        }

        private static List<TSG.Point> SortPointsAlongArcChord(
            IEnumerable<TSG.Point> points,
            TSG.Point arcStart,
            TSG.Point arcEnd
        ) {
            return points
                .Where(point => point != null)
                .OrderBy(point => GetPointParameterOnSegment(point, arcStart, arcEnd))
                .ToList();
        }

        private static List<TSG.Point> FilterPointsBySelectedSide(
            IEnumerable<TSG.Point> points,
            PartBounds bounds,
            DimensionOptions options,
            CurvedSideSelectionMode curvedSideSelectionMode
        ) {
            var validPoints = points?.Where(point => point != null).ToList() ?? new List<TSG.Point>();
            if (validPoints.Count == 0 || bounds == null || options == null) return new List<TSG.Point>();

            if (curvedSideSelectionMode == CurvedSideSelectionMode.Horizontal) {
                var midY = (bounds.MinY + bounds.MaxY) / 2.0;

                return validPoints
                    .Where(point => options.IsAbove ? point.Y >= midY : point.Y <= midY)
                    .Select(point => new TSG.Point(point.X, point.Y, 0))
                    .ToList();
            }

            var midX = (bounds.MinX + bounds.MaxX) / 2.0;

            return validPoints
                .Where(point => options.IsRight ? point.X >= midX : point.X <= midX)
                .Select(point => new TSG.Point(point.X, point.Y, 0))
                .ToList();
        }

        private static List<TSG.Point> GetAssemblyStartAndEndPoints(
            IEnumerable<TSG.Point> filteredPoints,
            DimensionOptions options
        ) {
            if (filteredPoints == null || !options.HasUserArcPoints)
                return new List<TSG.Point>();

            var points = filteredPoints
                .Where(point => point != null)
                .Select(point => new TSG.Point(point.X, point.Y, 0))
                .ToList();

            points = RemoveDuplicatePointsByDistance(points, DimensionMergeToleranceMillimeters);
            points = SortPointsAlongArcChord(points, options.UserArcStart, options.UserArcEnd);

            switch (points.Count) {
                case 0:
                    return new List<TSG.Point>();
                case 1:
                    return new List<TSG.Point> {
                        new TSG.Point(points[0].X, points[0].Y, 0)
                    };
                default:
                    return new List<TSG.Point> {
                        new TSG.Point(points[0].X, points[0].Y, 0),
                        new TSG.Point(points[points.Count - 1].X, points[points.Count - 1].Y, 0)
                    };
            }
        }

        private static List<TSG.Point> GetCurvedDimensionReferencePoints(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options,
            CurvedSideSelectionMode curvedSideSelectionMode,
            ElementEndpointsSnapshot element = null
        ) {
            if (snapshot == null || !options.HasUserArcPoints)
                return new List<TSG.Point>();

            var result = new List<TSG.Point>();
            var allFilteredPoints = new List<TSG.Point>();

            foreach (var currentElement in snapshot.ElementEndpoints.Where(currentElement =>
                         currentElement?.Endpoints != null)) {
                var currentBounds = GetElementLocalBounds(currentElement);
                var filteredPoints = FilterPointsBySelectedSide(
                    currentElement.Endpoints,
                    currentBounds,
                    options,
                    curvedSideSelectionMode
                );

                allFilteredPoints.AddRange(filteredPoints);
            }

            allFilteredPoints = RemoveDuplicatePointsByDistance(
                allFilteredPoints,
                DimensionMergeToleranceMillimeters
            );

            if (snapshot.AllPartsBounds != null) {
                if (curvedSideSelectionMode == CurvedSideSelectionMode.Horizontal) {
                    var anchorY = options.IsAbove
                        ? snapshot.AllPartsBounds.MaxY
                        : snapshot.AllPartsBounds.MinY;

                    allFilteredPoints.Add(new TSG.Point(snapshot.AllPartsBounds.MinX, anchorY, 0));
                    allFilteredPoints.Add(new TSG.Point(snapshot.AllPartsBounds.MaxX, anchorY, 0));
                }
                else {
                    var anchorX = options.IsRight
                        ? snapshot.AllPartsBounds.MaxX
                        : snapshot.AllPartsBounds.MinX;

                    allFilteredPoints.Add(new TSG.Point(anchorX, snapshot.AllPartsBounds.MinY, 0));
                    allFilteredPoints.Add(new TSG.Point(anchorX, snapshot.AllPartsBounds.MaxY, 0));
                }

                allFilteredPoints = RemoveDuplicatePointsByDistance(
                    allFilteredPoints,
                    DimensionMergeToleranceMillimeters
                );
            }

            var assemblyStartEnd = GetAssemblyStartAndEndPoints(allFilteredPoints, options);
            result.AddRange(assemblyStartEnd);

            if (element != null) {
                var elementBounds = GetElementLocalBounds(element);
                var filteredElementPoints = FilterPointsBySelectedSide(
                    element.Endpoints,
                    elementBounds,
                    options,
                    curvedSideSelectionMode
                );

                result.AddRange(filteredElementPoints);
            }
            else
                result.AddRange(allFilteredPoints);

            result = RemoveDuplicatePointsByDistance(result, DimensionMergeToleranceMillimeters);
            result = SortPointsAlongArcChord(result, options.UserArcStart, options.UserArcEnd);

            return result;
        }

        private static void CreateCurvedDimensionSet(
            TSD.View selectedView,
            List<TSG.Point> referencePoints,
            DimensionOptions options,
            double offsetMillimeters
        ) {
            if (selectedView == null || referencePoints == null || referencePoints.Count < 2) return;
            if (!options.HasUserArcPoints) return;

            var pointList = new TSD.PointList();
            foreach (var point in referencePoints)
                pointList.Add(new TSG.Point(point.X, point.Y, 0));

            var curvedHandler = new TSD.CurvedDimensionSetHandler();
            curvedHandler.CreateCurvedDimensionSetRadial(
                selectedView,
                options.UserArcStart,
                options.UserArcMid,
                options.UserArcEnd,
                pointList,
                offsetMillimeters
            );
        }

        #endregion

        #region Constants

        private const double FarDimensionOffsetMillimeters = 20.0;
        private const double PerElementDimensionOffsetMillimeters = 100.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;
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

        private enum DimensionAxis {
            Horizontal,
            Vertical
        }

        private enum DimensionPlacement {
            Above,
            Below,
            Left,
            Right
        }

        private enum DimensionScope {
            PerElement,
            Overall
        }

        private sealed class DimensionOptions {
            public DimensionType DimensionType { get; set; }
            public DimensionAxis Axis { get; set; }
            public DimensionPlacement Placement { get; set; }
            public DimensionScope Scope { get; set; }
            public TSG.Point UserArcStart { get; set; }
            public TSG.Point UserArcMid { get; set; }
            public TSG.Point UserArcEnd { get; set; }

            public bool HasUserArcPoints => UserArcStart != null && UserArcMid != null && UserArcEnd != null;

            public bool IsHorizontal => Axis == DimensionAxis.Horizontal;
            public bool IsOverall => Scope == DimensionScope.Overall;
            public bool IsAbove => Placement == DimensionPlacement.Above;
            public bool IsRight => Placement == DimensionPlacement.Right;
        }

        internal sealed class DrawingPartWithBounds {
            public TSD.Part DrawingPart { get; set; }
        }

        private sealed class PartBounds {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
        }

        private sealed class ElementEndpointsSnapshot {
            public PartBounds Bounds { get; set; }
            public List<TSG.Point> Endpoints { get; set; }
        }

        private sealed class DimensionGeometrySnapshot {
            public TSD.View View { get; set; }
            public PartBounds AllPartsBounds { get; set; }
            public PartBounds SelectedBounds { get; set; }
            public List<TSG.Point> Endpoints { get; set; }
            public List<ElementEndpointsSnapshot> ElementEndpoints { get; set; }
        }

        #endregion

        #region Dimension Creation

        private static void CreateStraightDimensions(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            if (snapshot?.View == null || options == null) return;
            if (snapshot.ElementEndpoints == null || snapshot.ElementEndpoints.Count == 0) return;

            if (options.IsOverall) {
                if (snapshot.AllPartsBounds == null) return;
                if (GetDimensionSpan(snapshot.AllPartsBounds, options) < MinimumDimensionSpanMillimeters) return;

                var mergedCoordinates = new List<double>();

                foreach (var element in snapshot.ElementEndpoints) {
                    var localBounds = GetElementLocalBounds(element);
                    if (localBounds == null) continue;
                    if (GetDimensionSpan(localBounds, options) < MinimumDimensionSpanMillimeters) continue;

                    mergedCoordinates.AddRange(GetCoordinatesForElement(element, snapshot, options));
                }

                var mergedAndSortedCoordinates = MergeAndSort(
                    mergedCoordinates,
                    DimensionMergeToleranceMillimeters
                );

                if (mergedAndSortedCoordinates.Count < 2) return;

                var anchorCoordinate = GetAnchorCoordinate(snapshot.AllPartsBounds, options);
                var dimensionPoints = CreateStraightDimensionPointList(
                    mergedAndSortedCoordinates,
                    anchorCoordinate,
                    options
                );

                CreateDimensionSet(
                    snapshot.View,
                    dimensionPoints,
                    GetDirectionVector(options),
                    GetOffset(options),
                    options
                );

                return;
            }

            foreach (var element in snapshot.ElementEndpoints) {
                if (element?.Endpoints == null || element.Endpoints.Count == 0) continue;

                var localBounds = GetElementLocalBounds(element);
                if (localBounds == null) continue;
                if (GetDimensionSpan(localBounds, options) < MinimumDimensionSpanMillimeters) continue;

                var coordinates = GetCoordinatesForElement(element, snapshot, options);
                if (coordinates.Count < 2) continue;

                var anchorCoordinate = GetAnchorCoordinate(localBounds, options);
                var dimensionPoints = CreateStraightDimensionPointList(
                    coordinates,
                    anchorCoordinate,
                    options
                );

                CreateDimensionSet(
                    snapshot.View,
                    dimensionPoints,
                    GetDirectionVector(options),
                    GetOffset(options),
                    options
                );
            }
        }

        private static double GetDimensionSpan(PartBounds bounds, DimensionOptions options) {
            if (bounds == null || options == null) return 0.0;

            return options.IsHorizontal
                ? bounds.MaxX - bounds.MinX
                : bounds.MaxY - bounds.MinY;
        }

        private static double GetAnchorCoordinate(PartBounds bounds, DimensionOptions options) {
            if (bounds == null || options == null) return 0.0;

            if (options.IsHorizontal)
                return options.IsAbove ? bounds.MaxY : bounds.MinY;

            return options.IsRight ? bounds.MaxX : bounds.MinX;
        }

        private static TSD.PointList CreateStraightDimensionPointList(
            IEnumerable<double> coordinates,
            double anchorCoordinate,
            DimensionOptions options
        ) {
            var pointList = new TSD.PointList();
            if (coordinates == null || options == null) return pointList;

            if (options.IsHorizontal)
                AddHorizontalPoints(pointList, coordinates, anchorCoordinate);
            else
                AddVerticalPoints(pointList, coordinates, anchorCoordinate);

            return pointList;
        }

        private static TSG.Vector GetDirectionVector(DimensionOptions options) {
            if (options == null) return new TSG.Vector(0.0, 0.0, 0.0);

            if (options.IsHorizontal)
                return options.IsAbove
                    ? new TSG.Vector(0.0, 1.0, 0.0)
                    : new TSG.Vector(0.0, -1.0, 0.0);

            return options.IsRight
                ? new TSG.Vector(1.0, 0.0, 0.0)
                : new TSG.Vector(-1.0, 0.0, 0.0);
        }

        private static double GetOffset(DimensionOptions options) {
            return options?.IsOverall == true
                ? FarDimensionOffsetMillimeters
                : PerElementDimensionOffsetMillimeters;
        }

        private static bool PickCurvedArcPoints(TSD.DrawingHandler drawingHandler, DimensionOptions options) {
            try {
                var picker = drawingHandler.GetPicker();
                var prompts = new TSD.StringList {
                    "Wskaż punkt startowy łuku",
                    "Wskaż punkt środkowy łuku",
                    "Wskaż punkt końcowy łuku"
                };

                picker.PickPoints(3, prompts, out var picked, out _);
                var pts = picked?.OfType<TSG.Point>().ToList();
                if (pts == null || pts.Count < 3) return false;

                options.UserArcStart = new TSG.Point(pts[0].X, pts[0].Y, 0);
                options.UserArcMid = new TSG.Point(pts[1].X, pts[1].Y, 0);
                options.UserArcEnd = new TSG.Point(pts[2].X, pts[2].Y, 0);
                return true;
            }
            catch {
                return false;
            }
        }

        private static void CreateDimensionSet(
            TSD.View selectedView,
            TSD.PointList dimensionPoints,
            TSG.Vector directionVector,
            double offsetMillimeters,
            DimensionOptions options
        ) {
            if (options.DimensionType == DimensionType.Straight) {
                var straightHandler = new TSD.StraightDimensionSetHandler();
                straightHandler.CreateDimensionSet(
                    selectedView,
                    dimensionPoints,
                    directionVector,
                    offsetMillimeters
                );
                return;
            }

            if (!options.HasUserArcPoints)
                return;

            var curvedHandler = new TSD.CurvedDimensionSetHandler();
            curvedHandler.CreateCurvedDimensionSetRadial(
                selectedView,
                options.UserArcStart,
                options.UserArcMid,
                options.UserArcEnd,
                dimensionPoints,
                offsetMillimeters
            );
        }

        #endregion

        #region Part Selection And View Validation

        private static List<DrawingPartWithBounds> GetSelectedDrawingParts(TSD.DrawingHandler drawingHandler) {
            if (_overrideSelectedParts != null)
                return _overrideSelectedParts
                    .Where(part => part != null)
                    .Select(part => new DrawingPartWithBounds { DrawingPart = part })
                    .ToList();

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
                    height < MinimumPartSizeForDimensionMillimeters)
                    return null;

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

            if (maxX - minX < 1e-6) return points;

            var xValues = new List<double> { minX };
            for (var x = minX + SweepStepMillimeters; x < maxX; x += SweepStepMillimeters)
                xValues.Add(x);
            xValues.Add(maxX);

            foreach (var xValue in xValues) {
                var enumerator = solid.GetAllIntersectionPoints(
                    new TSG.Point(xValue, 0, 0),
                    new TSG.Point(xValue, 1, 0),
                    new TSG.Point(xValue, 0, 1)
                );

                while (enumerator.MoveNext())
                    if (enumerator.Current is TSG.Point p)
                        points.Add(new TSG.Point(xValue, p.Y, 0));
            }

            points = RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
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

                        if (next.Count == 0 || !ArePointsEqual(next[next.Count - 1], first))
                            next.Add(first);

                        if (distance <= toleranceMillimeters) {
                            next.Add(last);
                            changed = true;
                        }
                        else {
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

            return outline;
        }

        #endregion

        #region Part Edge Outline Drawing

        private static List<ElementEndpointsSnapshot> DrawPartEdgeOutlines(
            TSD.View view,
            List<DrawingPartWithBounds> drawingParts
        ) {
            var result = new List<ElementEndpointsSnapshot>();
            if (view == null || drawingParts == null) return result;

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

                    var bounds = GetPartAabbInViewSpace(view, modelPart);
                    if (bounds == null) continue;

                    List<TSG.Point> outline;

                    if (ShouldUseSweepOutline(modelPart)) {
                        var points = GetSweptPlateOutlinePoints(solid);
                        if (points.Count < 2) continue;

                        outline = BuildUpperLowerChainOutline(points);
                        if (outline == null || outline.Count < 3) continue;

                        outline = PrepareSegmentPath(outline);
                        outline = EnsureClosedOutline(outline);

                        if (outline == null || outline.Count < 4) continue;
                    }
                    else {
                        var backZ = solid.MinimumPoint.Z;
                        const double inward = 1.0;

                        if (solid.MaximumPoint.Z - solid.MinimumPoint.Z > inward * 2)
                            backZ += inward;
                        else
                            backZ = (solid.MinimumPoint.Z + solid.MaximumPoint.Z) * 0.5;

                        var backPoints = GetIntersectionPointsAtLocalZ(solid, backZ);
                        if (backPoints == null || backPoints.Count < 3) continue;

                        backPoints = RemoveNearDuplicates(backPoints, DuplicateToleranceMillimeters);
                        outline = BuildConvexHull2D(backPoints);

                        if (outline == null || outline.Count < 3) continue;

                        outline = PrepareSegmentPath(outline);
                        outline = EnsureClosedOutline(outline);

                        if (outline == null || outline.Count < 4) continue;
                    }

                    DrawOutlineByAngleType(view, outline);

                    var realVertices = GetOpenOutlineVertices(outline);
                    if (realVertices == null || realVertices.Count == 0) continue;

                    result.Add(new ElementEndpointsSnapshot {
                        Bounds = bounds,
                        Endpoints = realVertices.Select(p => new TSG.Point(p.X, p.Y, 0)).ToList()
                    });
                }
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }

            return result;
        }

        private static void DrawOutlineByAngleType(
            TSD.View view,
            List<TSG.Point> outline
        ) {
            if (view == null || outline == null || outline.Count < 2) return;

            var vertices = GetOpenOutlineVertices(outline);
            if (vertices == null || vertices.Count < 2) return;

            if (vertices.Count == 2) {
                DrawStraightSegmentPrimitive(view, vertices[0], vertices[1], TSD.DrawingColors.Green);
                return;
            }

            var significantCornerIndices = GetSignificantCornerIndices(vertices);

            switch (significantCornerIndices.Count) {
                case 0:
                case 1:
                    DrawPolylinePrimitive(view, vertices, TSD.DrawingColors.Green);
                    return;
            }

            for (var i = 0; i < significantCornerIndices.Count; i++) {
                var startIndex = significantCornerIndices[i];
                var endIndex = significantCornerIndices[(i + 1) % significantCornerIndices.Count];

                var path = CollectPathBetweenCorners(vertices, startIndex, endIndex);
                path = PrepareSegmentPath(path);

                if (path == null || path.Count < 2) continue;

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

            if (distance <= DuplicateToleranceMillimeters) return;

            var points = new TSD.PointList {
                new TSG.Point(startPoint.X, startPoint.Y, 0),
                new TSG.Point(endPoint.X, endPoint.Y, 0)
            };

            var polyline = new TSD.Polyline(view, points);
            polyline.Attributes.Line.Color = color;
            polyline.Insert();
        }

        private static void DrawPolylinePrimitive(
            TSD.View view,
            List<TSG.Point> points,
            TSD.DrawingColors color
        ) {
            if (view == null || points == null || points.Count < 2) return;

            var cleanedPoints = RemoveNearDuplicates(
                new List<TSG.Point>(points),
                DuplicateToleranceMillimeters
            );

            if (cleanedPoints == null || cleanedPoints.Count < 2) return;

            switch (cleanedPoints.Count) {
                case 2:
                    DrawStraightSegmentPrimitive(view, cleanedPoints[0], cleanedPoints[1], color);
                    return;
                case 3: {
                    var first = cleanedPoints[0];
                    var last = cleanedPoints[cleanedPoints.Count - 1];

                    var distance = Math.Sqrt(
                        (first.X - last.X) * (first.X - last.X) +
                        (first.Y - last.Y) * (first.Y - last.Y)
                    );

                    if (distance <= DuplicateToleranceMillimeters) {
                        DrawStraightSegmentPrimitive(view, cleanedPoints[0], cleanedPoints[1], color);
                        return;
                    }

                    break;
                }
            }

            var polylinePoints = new TSD.PointList();
            foreach (var point in cleanedPoints)
                polylinePoints.Add(new TSG.Point(point.X, point.Y, 0));

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