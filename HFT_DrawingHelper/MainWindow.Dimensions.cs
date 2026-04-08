using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
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

            var outlineSnapshots = GetPartOutlineSnapshots(selectedViewFromParts, selectedParts);
            if (outlineSnapshots == null || outlineSnapshots.Count == 0) return null;

            var allViewBoundsFromParts = GetPartBoundsFromView(selectedViewFromParts);
            var allPartsBoundsFromParts = GetAssemblyBounds(allViewBoundsFromParts);

            if (allPartsBoundsFromParts == null) return null;

            var selectedBoundsList = outlineSnapshots
                .Where(snapshot => snapshot?.Bounds != null)
                .Select(snapshot => snapshot.Bounds)
                .ToList();

            var selectedBounds = GetAssemblyBounds(selectedBoundsList);
            if (selectedBounds == null) return null;

            var elementEndpoints = outlineSnapshots
                .Where(snapshot => snapshot?.Vertices != null && snapshot.Vertices.Count > 0)
                .Select(snapshot => new ElementEndpointsSnapshot {
                    Bounds = snapshot.Bounds,
                    Endpoints = snapshot.Vertices
                        .Select(point => new TSG.Point(point.X, point.Y, 0))
                        .ToList()
                })
                .ToList();

            var mergedEndpoints = elementEndpoints
                .Where(element => element?.Endpoints != null)
                .SelectMany(element => element.Endpoints)
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

                if (validOptions.Where(option => option.DimensionType == DimensionType.Curved)
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

        #region Geometry Helpers

        private enum CurvedSideSelectionMode {
            Horizontal,
            Vertical
        }

        private static List<double> MergeAndSort(IEnumerable<double> rawValues, double tolerance) {
            var sorted = rawValues
                .Where(value => !double.IsNaN(value) && !double.IsInfinity(value))
                .OrderBy(value => value)
                .ToList();

            if (sorted.Count == 0) return new List<double>();

            var result = new List<double>();
            var bucketStart = sorted[0];
            var bucket = new List<double> { sorted[0] };

            for (var index = 1; index < sorted.Count; index++)
                if (sorted[index] - bucketStart <= tolerance)
                    bucket.Add(sorted[index]);
                else {
                    result.Add(bucket.Average());
                    bucketStart = sorted[index];
                    bucket = new List<double> { sorted[index] };
                }

            result.Add(bucket.Average());
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

        private static void AddHorizontalPoints(
            TSD.PointList points,
            IEnumerable<double> xCoordinates,
            double anchorY
        ) {
            foreach (var x in xCoordinates)
                points.Add(new TSG.Point(x, anchorY, 0));
        }

        private static void AddVerticalPoints(
            TSD.PointList points,
            IEnumerable<double> yCoordinates,
            double anchorX
        ) {
            foreach (var y in yCoordinates)
                points.Add(new TSG.Point(anchorX, y, 0));
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
                var points = picked?.OfType<TSG.Point>().ToList();
                if (points == null || points.Count < 3) return false;

                options.UserArcStart = new TSG.Point(points[0].X, points[0].Y, 0);
                options.UserArcMid = new TSG.Point(points[1].X, points[1].Y, 0);
                options.UserArcEnd = new TSG.Point(points[2].X, points[2].Y, 0);
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

        #region Constants

        private const double FarDimensionOffsetMillimeters = 20.0;
        private const double PerElementDimensionOffsetMillimeters = 100.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;
        private const double DimensionMergeToleranceMillimeters = 1.0;

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
    }
}