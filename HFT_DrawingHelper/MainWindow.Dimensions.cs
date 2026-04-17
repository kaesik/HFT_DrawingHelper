using System;
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

                var drawingPartsFromView = GetDrawingPartsFromView(selectedView);
                if (drawingPartsFromView == null || drawingPartsFromView.Count == 0) return null;

                var outlineSnapshotsFromView = GetPartOutlineSnapshots(selectedView, drawingPartsFromView);
                if (outlineSnapshotsFromView == null || outlineSnapshotsFromView.Count == 0) return null;

                var allPartsBounds = GetAssemblyBounds(
                    outlineSnapshotsFromView
                        .Where(snapshot => snapshot?.Bounds != null)
                        .Select(snapshot => snapshot.Bounds)
                        .ToList()
                );

                if (allPartsBounds == null) return null;

                var elementEndpointsFromView = outlineSnapshotsFromView
                    .Where(snapshot => snapshot?.Vertices != null && snapshot.Vertices.Count > 0)
                    .Select(snapshot => new ElementEndpointsSnapshot {
                        Bounds = snapshot.Bounds,
                        Endpoints = snapshot.Vertices
                            .Select(point => new TSG.Point(point.X, point.Y, 0))
                            .ToList()
                    })
                    .ToList();

                var mergedEndpointsFromView = elementEndpointsFromView
                    .Where(element => element?.Endpoints != null)
                    .SelectMany(element => element.Endpoints)
                    .ToList();

                mergedEndpointsFromView.AddRange(GetLineAndPolylineEndpoints(selectedView));
                mergedEndpointsFromView = RemoveDuplicatePointsByDistance(
                    mergedEndpointsFromView,
                    DimensionMergeToleranceMillimeters
                );

                return new DimensionGeometrySnapshot {
                    View = selectedView,
                    AllPartsBounds = allPartsBounds,
                    SelectedBounds = allPartsBounds,
                    Endpoints = mergedEndpointsFromView,
                    ElementEndpoints = elementEndpointsFromView
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

            var selectedBounds = GetAssemblyBounds(
                outlineSnapshots
                    .Where(snapshot => snapshot?.Bounds != null)
                    .Select(snapshot => snapshot.Bounds)
                    .ToList()
            );

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
                            option.IsCurved ||
                            option.IsAngled ||
                            (
                                option.Axis == DimensionAxis.Horizontal &&
                                (option.Placement == DimensionPlacement.Above ||
                                 option.Placement == DimensionPlacement.Below)
                            ) ||
                            (
                                option.Axis == DimensionAxis.Vertical &&
                                (option.Placement == DimensionPlacement.Left ||
                                 option.Placement == DimensionPlacement.Right)
                            )
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

                foreach (var option in validOptions.Where(option => option.IsCurved))
                    if (!PickCurvedArcPoints(drawingHandler, option))
                        return;

                var firstAngledOption = validOptions.FirstOrDefault(option => option.IsAngled);
                if (firstAngledOption != null) {
                    if (!PickAngledVectorPoints(drawingHandler, firstAngledOption))
                        return;

                    foreach (var angledOption in validOptions.Where(option =>
                                 option.IsAngled && !ReferenceEquals(option, firstAngledOption))) {
                        angledOption.UserVectorStart = new TSG.Point(
                            firstAngledOption.UserVectorStart.X,
                            firstAngledOption.UserVectorStart.Y,
                            0
                        );
                        angledOption.UserVectorEnd = new TSG.Point(
                            firstAngledOption.UserVectorEnd.X,
                            firstAngledOption.UserVectorEnd.Y,
                            0
                        );
                    }
                }

                var snapshot = BuildDimensionGeometrySnapshot(drawingHandler);
                if (snapshot?.View == null || snapshot.AllPartsBounds == null || snapshot.SelectedBounds == null)
                    return;

                foreach (var option in validOptions) {
                    if (option.IsCurved) {
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

                    if (option.IsAngled) {
                        CreateAngledDimensions(snapshot, option);
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

        private static PartBounds GetElementLocalBounds(ElementEndpointsSnapshot element) {
            return GetBoundsFromPoints(element?.Endpoints);
        }

        private static List<TSG.Point> GetAssemblyCornerCandidates(PartBounds bounds) {
            if (bounds == null)
                return new List<TSG.Point>();

            return new List<TSG.Point> {
                new TSG.Point(bounds.MinX, bounds.MinY, 0),
                new TSG.Point(bounds.MinX, bounds.MaxY, 0),
                new TSG.Point(bounds.MaxX, bounds.MinY, 0),
                new TSG.Point(bounds.MaxX, bounds.MaxY, 0)
            };
        }

        private static List<TSG.Point> GetStraightOverallStartAndEndPoints(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            var candidates = GetAssemblyCornerCandidates(snapshot?.AllPartsBounds);
            if (candidates.Count == 0 || options == null)
                return new List<TSG.Point>();

            TSG.Point startPoint;
            TSG.Point endPoint;

            if (options.IsHorizontal) {
                var minX = candidates.Min(point => point.X);
                var maxX = candidates.Max(point => point.X);

                startPoint = SelectBestStraightReferencePoint(
                    candidates.Where(point => Math.Abs(point.X - minX) <= DimensionMergeToleranceMillimeters),
                    options
                );

                endPoint = SelectBestStraightReferencePoint(
                    candidates.Where(point => Math.Abs(point.X - maxX) <= DimensionMergeToleranceMillimeters),
                    options
                );
            }
            else {
                var minY = candidates.Min(point => point.Y);
                var maxY = candidates.Max(point => point.Y);

                startPoint = SelectBestStraightReferencePoint(
                    candidates.Where(point => Math.Abs(point.Y - minY) <= DimensionMergeToleranceMillimeters),
                    options
                );

                endPoint = SelectBestStraightReferencePoint(
                    candidates.Where(point => Math.Abs(point.Y - maxY) <= DimensionMergeToleranceMillimeters),
                    options
                );
            }

            return RemoveDuplicatePointsByDistance(
                new[] { startPoint, endPoint }.Where(point => point != null),
                DimensionMergeToleranceMillimeters
            );
        }

        private static List<TSG.Point> GetAngledOverallStartAndEndPoints(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            if (snapshot == null || options == null || !options.HasUserVector)
                return new List<TSG.Point>();

            var axisVector = GetNormalizedPickedVector(options);
            var sideNormal = GetSelectedAngledSideNormal(options);

            if (axisVector == null || sideNormal == null)
                return new List<TSG.Point>();

            var candidates = snapshot.ElementEndpoints
                .Where(element => element?.Endpoints != null)
                .SelectMany(element => element.Endpoints)
                .Select(point => new TSG.Point(point.X, point.Y, 0))
                .ToList();

            candidates.AddRange(GetAssemblyCornerCandidates(snapshot.AllPartsBounds));
            candidates = RemoveDuplicatePointsByDistance(
                candidates,
                DimensionMergeToleranceMillimeters
            );

            if (candidates.Count == 0)
                return new List<TSG.Point>();

            var projectedCandidates = candidates
                .Select(point => new ProjectedPoint {
                    Point = point,
                    AxisValue = GetProjectedCoordinate(point, options.UserVectorStart, axisVector),
                    SideValue = GetProjectedCoordinate(point, options.UserVectorStart, sideNormal)
                })
                .ToList();

            var minAxis = projectedCandidates.Min(item => item.AxisValue);
            var maxAxis = projectedCandidates.Max(item => item.AxisValue);

            var startPoint = projectedCandidates
                .Where(item => Math.Abs(item.AxisValue - minAxis) <= DimensionMergeToleranceMillimeters)
                .OrderByDescending(item => item.SideValue)
                .Select(item => item.Point)
                .FirstOrDefault();

            var endPoint = projectedCandidates
                .Where(item => Math.Abs(item.AxisValue - maxAxis) <= DimensionMergeToleranceMillimeters)
                .OrderByDescending(item => item.SideValue)
                .Select(item => item.Point)
                .FirstOrDefault();

            return RemoveDuplicatePointsByDistance(
                new[] { startPoint, endPoint }.Where(point => point != null),
                DimensionMergeToleranceMillimeters
            );
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

        private static List<TSG.Point> GetStraightSidePoints(
            IEnumerable<TSG.Point> points,
            PartBounds bounds,
            DimensionOptions options
        ) {
            return FilterPointsBySelectedSide(
                points,
                bounds,
                options,
                options.IsHorizontal
                    ? CurvedSideSelectionMode.Horizontal
                    : CurvedSideSelectionMode.Vertical
            );
        }

        private static TSG.Point SelectBestStraightReferencePoint(
            IEnumerable<TSG.Point> bucket,
            DimensionOptions options
        ) {
            var validPoints = bucket?.Where(point => point != null).ToList() ?? new List<TSG.Point>();
            if (validPoints.Count == 0 || options == null)
                return null;

            if (options.IsHorizontal)
                return options.IsAbove
                    ? validPoints
                        .OrderByDescending(point => point.Y)
                        .ThenBy(point => point.X)
                        .First()
                    : validPoints
                        .OrderBy(point => point.Y)
                        .ThenBy(point => point.X)
                        .First();

            return options.IsRight
                ? validPoints
                    .OrderByDescending(point => point.X)
                    .ThenBy(point => point.Y)
                    .First()
                : validPoints
                    .OrderBy(point => point.X)
                    .ThenBy(point => point.Y)
                    .First();
        }

        private static List<TSG.Point> ReduceStraightReferencePoints(
            IEnumerable<TSG.Point> points,
            DimensionOptions options
        ) {
            var validPoints = points?
                .Where(point => point != null)
                .Select(point => new TSG.Point(point.X, point.Y, 0))
                .ToList() ?? new List<TSG.Point>();

            if (validPoints.Count == 0 || options == null)
                return new List<TSG.Point>();

            Func<TSG.Point, double> axisSelector;
            if (options.IsHorizontal)
                axisSelector = point => point.X;
            else
                axisSelector = point => point.Y;

            validPoints = validPoints
                .OrderBy(axisSelector)
                .ToList();

            var reducedPoints = new List<TSG.Point>();
            var bucket = new List<TSG.Point> { validPoints[0] };
            var bucketStart = axisSelector(validPoints[0]);

            for (var index = 1; index < validPoints.Count; index++) {
                var currentAxisValue = axisSelector(validPoints[index]);

                if (currentAxisValue - bucketStart <= DimensionMergeToleranceMillimeters) {
                    bucket.Add(validPoints[index]);
                    continue;
                }

                var selectedPoint = SelectBestStraightReferencePoint(bucket, options);
                if (selectedPoint != null)
                    reducedPoints.Add(new TSG.Point(selectedPoint.X, selectedPoint.Y, 0));

                bucket = new List<TSG.Point> { validPoints[index] };
                bucketStart = currentAxisValue;
            }

            var lastSelectedPoint = SelectBestStraightReferencePoint(bucket, options);
            if (lastSelectedPoint != null)
                reducedPoints.Add(new TSG.Point(lastSelectedPoint.X, lastSelectedPoint.Y, 0));

            reducedPoints = RemoveDuplicatePointsByDistance(
                reducedPoints,
                DimensionMergeToleranceMillimeters
            );

            return reducedPoints
                .OrderBy(axisSelector)
                .ToList();
        }

        private static List<TSG.Point> GetStraightDimensionReferencePoints(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options,
            ElementEndpointsSnapshot element = null
        ) {
            if (snapshot == null || options == null)
                return new List<TSG.Point>();

            var points = new List<TSG.Point>();

            if (element != null) {
                var localBounds = GetElementLocalBounds(element);
                points.AddRange(GetStraightSidePoints(element.Endpoints, localBounds, options));
            }
            else {
                foreach (var currentElement in snapshot.ElementEndpoints.Where(currentElement =>
                             currentElement?.Endpoints != null)) {
                    var currentBounds = GetElementLocalBounds(currentElement);
                    points.AddRange(GetStraightSidePoints(currentElement.Endpoints, currentBounds, options));
                }

                if (options.IsOverall)
                    points.AddRange(GetStraightOverallStartAndEndPoints(snapshot, options));
            }

            return ReduceStraightReferencePoints(points, options);
        }

        private static double GetStraightDistanceFromReferencePoints(
            IList<TSG.Point> referencePoints,
            DimensionOptions options,
            double baseOffsetMillimeters
        ) {
            if (referencePoints == null || referencePoints.Count == 0 || options == null)
                return baseOffsetMillimeters;

            if (options.IsHorizontal) {
                var firstPointY = referencePoints[0].Y;
                var extremeY = options.IsAbove
                    ? referencePoints.Max(point => point.Y)
                    : referencePoints.Min(point => point.Y);

                return baseOffsetMillimeters + Math.Abs(extremeY - firstPointY);
            }

            var firstPointX = referencePoints[0].X;
            var extremeX = options.IsRight
                ? referencePoints.Max(point => point.X)
                : referencePoints.Min(point => point.X);

            return baseOffsetMillimeters + Math.Abs(extremeX - firstPointX);
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

            if (points.Count == 0)
                return new List<TSG.Point>();

            if (points.Count == 1)
                return new List<TSG.Point> {
                    new TSG.Point(points[0].X, points[0].Y, 0)
                };

            return new List<TSG.Point> {
                new TSG.Point(points[0].X, points[0].Y, 0),
                new TSG.Point(points[points.Count - 1].X, points[points.Count - 1].Y, 0)
            };
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

            result.AddRange(GetAssemblyStartAndEndPoints(allFilteredPoints, options));

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

        private static List<TSG.Point> GetAngledSidePoints(
            IEnumerable<TSG.Point> points,
            DimensionOptions options
        ) {
            var validPoints = points?.Where(point => point != null)
                .Select(point => new TSG.Point(point.X, point.Y, 0))
                .ToList() ?? new List<TSG.Point>();

            if (validPoints.Count == 0 || options == null || !options.HasUserVector)
                return new List<TSG.Point>();

            validPoints = RemoveDuplicatePointsByDistance(validPoints, DuplicateToleranceMillimeters);

            if (validPoints.Count > 1 && ArePointsClose2D(validPoints[0], validPoints[validPoints.Count - 1]))
                validPoints.RemoveAt(validPoints.Count - 1);

            if (validPoints.Count < 2)
                return validPoints;

            var axisVector = GetNormalizedPickedVector(options);
            var sideNormal = GetSelectedAngledSideNormal(options);

            if (axisVector == null || sideNormal == null)
                return new List<TSG.Point>();

            var projected = validPoints
                .Select((point, index) => new ProjectedPoint {
                    Index = index,
                    Point = point,
                    AxisValue = GetProjectedCoordinate(point, options.UserVectorStart, axisVector),
                    SideValue = GetProjectedCoordinate(point, options.UserVectorStart, sideNormal)
                })
                .ToList();

            var minAxis = projected.Min(item => item.AxisValue);
            var maxAxis = projected.Max(item => item.AxisValue);

            var startCandidate = projected
                .Where(item => Math.Abs(item.AxisValue - minAxis) <= DimensionMergeToleranceMillimeters)
                .OrderByDescending(item => item.SideValue)
                .FirstOrDefault();

            var endCandidate = projected
                .Where(item => Math.Abs(item.AxisValue - maxAxis) <= DimensionMergeToleranceMillimeters)
                .OrderByDescending(item => item.SideValue)
                .FirstOrDefault();

            if (startCandidate == null || endCandidate == null)
                return validPoints;

            if (startCandidate.Index == endCandidate.Index)
                return validPoints;

            var forwardPath = CollectPath(validPoints, startCandidate.Index, endCandidate.Index, true);
            var backwardPath = CollectPath(validPoints, startCandidate.Index, endCandidate.Index, false);

            var selectedPath = GetAverageSideValue(forwardPath, options, sideNormal) >=
                               GetAverageSideValue(backwardPath, options, sideNormal)
                ? forwardPath
                : backwardPath;

            return RemoveDuplicatePointsByDistance(selectedPath, DuplicateToleranceMillimeters);
        }

        private static TSG.Point SelectBestAngledReferencePoint(
            IEnumerable<ProjectedPoint> bucket
        ) {
            return bucket?
                .OrderByDescending(item => item.SideValue)
                .Select(item => item.Point)
                .FirstOrDefault();
        }

        private static List<TSG.Point> ReduceAngledReferencePoints(
            IEnumerable<TSG.Point> points,
            DimensionOptions options
        ) {
            if (points == null || options == null || !options.HasUserVector)
                return new List<TSG.Point>();

            var axisVector = GetNormalizedPickedVector(options);
            var sideNormal = GetSelectedAngledSideNormal(options);

            if (axisVector == null || sideNormal == null)
                return new List<TSG.Point>();

            var projectedPoints = points
                .Where(point => point != null)
                .Select(point => new ProjectedPoint {
                    Point = new TSG.Point(point.X, point.Y, 0),
                    AxisValue = GetProjectedCoordinate(point, options.UserVectorStart, axisVector),
                    SideValue = GetProjectedCoordinate(point, options.UserVectorStart, sideNormal)
                })
                .OrderBy(item => item.AxisValue)
                .ToList();

            if (projectedPoints.Count == 0)
                return new List<TSG.Point>();

            var reducedPoints = new List<TSG.Point>();
            var bucket = new List<ProjectedPoint> { projectedPoints[0] };
            var bucketStart = projectedPoints[0].AxisValue;

            for (var index = 1; index < projectedPoints.Count; index++) {
                if (projectedPoints[index].AxisValue - bucketStart <= DimensionMergeToleranceMillimeters) {
                    bucket.Add(projectedPoints[index]);
                    continue;
                }

                var selectedPoint = SelectBestAngledReferencePoint(bucket);
                if (selectedPoint != null)
                    reducedPoints.Add(new TSG.Point(selectedPoint.X, selectedPoint.Y, 0));

                bucket = new List<ProjectedPoint> { projectedPoints[index] };
                bucketStart = projectedPoints[index].AxisValue;
            }

            var lastSelectedPoint = SelectBestAngledReferencePoint(bucket);
            if (lastSelectedPoint != null)
                reducedPoints.Add(new TSG.Point(lastSelectedPoint.X, lastSelectedPoint.Y, 0));

            reducedPoints = RemoveDuplicatePointsByDistance(
                reducedPoints,
                DimensionMergeToleranceMillimeters
            );

            return reducedPoints
                .OrderBy(point => GetProjectedCoordinate(point, options.UserVectorStart, axisVector))
                .ToList();
        }

        private static List<TSG.Point> GetAngledDimensionReferencePoints(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options,
            ElementEndpointsSnapshot element = null
        ) {
            if (snapshot == null || options == null || !options.HasUserVector)
                return new List<TSG.Point>();

            var points = new List<TSG.Point>();

            if (element != null)
                points.AddRange(GetAngledSidePoints(element.Endpoints, options));
            else {
                foreach (var currentElement in snapshot.ElementEndpoints.Where(currentElement =>
                             currentElement?.Endpoints != null))
                    points.AddRange(GetAngledSidePoints(currentElement.Endpoints, options));

                if (options.IsOverall)
                    points.AddRange(GetAngledOverallStartAndEndPoints(snapshot, options));
            }

            return ReduceAngledReferencePoints(points, options);
        }

        private static double GetAngledDistanceFromReferencePoints(
            IList<TSG.Point> referencePoints,
            DimensionOptions options,
            double baseOffsetMillimeters
        ) {
            if (referencePoints == null || referencePoints.Count == 0 || options == null || !options.HasUserVector)
                return baseOffsetMillimeters;

            var sideNormal = GetSelectedAngledSideNormal(options);
            if (sideNormal == null)
                return baseOffsetMillimeters;

            var firstSideCoordinate = GetProjectedCoordinate(
                referencePoints[0],
                options.UserVectorStart,
                sideNormal
            );

            var extremeSideCoordinate = referencePoints.Max(point =>
                GetProjectedCoordinate(point, options.UserVectorStart, sideNormal));

            return baseOffsetMillimeters + Math.Abs(extremeSideCoordinate - firstSideCoordinate);
        }

        private static TSD.PointList CreateDimensionPointList(
            IEnumerable<TSG.Point> referencePoints
        ) {
            var pointList = new TSD.PointList();
            if (referencePoints == null)
                return pointList;

            foreach (var point in referencePoints.Where(point => point != null))
                pointList.Add(new TSG.Point(point.X, point.Y, 0));

            return pointList;
        }

        private static List<TSG.Point> CollectPath(
            List<TSG.Point> points,
            int startIndex,
            int endIndex,
            bool forward
        ) {
            var result = new List<TSG.Point>();
            if (points == null || points.Count == 0)
                return result;

            var index = startIndex;
            result.Add(new TSG.Point(points[index].X, points[index].Y, 0));

            while (index != endIndex) {
                index = forward
                    ? (index + 1) % points.Count
                    : (index - 1 + points.Count) % points.Count;

                result.Add(new TSG.Point(points[index].X, points[index].Y, 0));

                if (result.Count > points.Count + 1)
                    break;
            }

            return result;
        }

        private static double GetAverageSideValue(
            IEnumerable<TSG.Point> points,
            DimensionOptions options,
            TSG.Vector sideNormal
        ) {
            var validPoints = points?.Where(point => point != null).ToList() ?? new List<TSG.Point>();
            if (validPoints.Count == 0 || options == null || sideNormal == null)
                return double.MinValue;

            return validPoints.Average(point =>
                GetProjectedCoordinate(point, options.UserVectorStart, sideNormal));
        }

        private static double GetProjectedCoordinate(
            TSG.Point point,
            TSG.Point origin,
            TSG.Vector direction
        ) {
            return (point.X - origin.X) * direction.X + (point.Y - origin.Y) * direction.Y;
        }

        private static TSG.Vector GetNormalizedPickedVector(DimensionOptions options) {
            if (options == null || !options.HasUserVector)
                return null;

            var vector = new TSG.Vector(
                options.UserVectorEnd.X - options.UserVectorStart.X,
                options.UserVectorEnd.Y - options.UserVectorStart.Y,
                0
            );

            if (vector.GetLength() < MinimumPickedVectorLengthMillimeters)
                return null;

            vector.Normalize();
            return vector;
        }

        private static TSG.Vector GetSelectedAngledSideNormal(DimensionOptions options) {
            var axisVector = GetNormalizedPickedVector(options);
            if (axisVector == null)
                return null;

            var firstNormal = new TSG.Vector(-axisVector.Y, axisVector.X, 0);
            var secondNormal = new TSG.Vector(axisVector.Y, -axisVector.X, 0);

            switch (options.Placement) {
                case DimensionPlacement.Above:
                    return firstNormal.Y >= secondNormal.Y ? firstNormal : secondNormal;
                case DimensionPlacement.Below:
                    return firstNormal.Y <= secondNormal.Y ? firstNormal : secondNormal;
                case DimensionPlacement.Right:
                    return firstNormal.X >= secondNormal.X ? firstNormal : secondNormal;
                case DimensionPlacement.Left:
                    return firstNormal.X <= secondNormal.X ? firstNormal : secondNormal;
                default:
                    return firstNormal;
            }
        }

        private static TSG.Vector GetAngledDirectionVector(DimensionOptions options) {
            var sideNormal = GetSelectedAngledSideNormal(options);
            if (sideNormal == null)
                return new TSG.Vector(0.0, 0.0, 0.0);

            return new TSG.Vector(
                sideNormal.X * 100.0,
                sideNormal.Y * 100.0,
                0.0
            );
        }

        private static bool ArePointsClose2D(TSG.Point firstPoint, TSG.Point secondPoint) {
            if (firstPoint == null || secondPoint == null)
                return false;

            return Math.Abs(firstPoint.X - secondPoint.X) <= DuplicateToleranceMillimeters &&
                   Math.Abs(firstPoint.Y - secondPoint.Y) <= DuplicateToleranceMillimeters;
        }

        private static double GetDistance2D(TSG.Point firstPoint, TSG.Point secondPoint) {
            if (firstPoint == null || secondPoint == null)
                return 0.0;

            var deltaX = secondPoint.X - firstPoint.X;
            var deltaY = secondPoint.Y - firstPoint.Y;
            return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
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

                var referencePoints = GetStraightDimensionReferencePoints(snapshot, options);
                if (referencePoints.Count < 2) return;

                var dimensionPoints = CreateDimensionPointList(referencePoints);
                var distance = GetStraightDistanceFromReferencePoints(
                    referencePoints,
                    options,
                    GetOffset(options)
                );

                CreateDimensionSet(
                    snapshot.View,
                    dimensionPoints,
                    GetDirectionVector(options),
                    distance,
                    options
                );

                return;
            }

            foreach (var element in snapshot.ElementEndpoints) {
                if (element?.Endpoints == null || element.Endpoints.Count == 0) continue;

                var localBounds = GetElementLocalBounds(element);
                if (localBounds == null) continue;
                if (GetDimensionSpan(localBounds, options) < MinimumDimensionSpanMillimeters) continue;

                var referencePoints = GetStraightDimensionReferencePoints(snapshot, options, element);
                if (referencePoints.Count < 2) continue;

                var dimensionPoints = CreateDimensionPointList(referencePoints);
                var distance = GetStraightDistanceFromReferencePoints(
                    referencePoints,
                    options,
                    GetOffset(options)
                );

                CreateDimensionSet(
                    snapshot.View,
                    dimensionPoints,
                    GetDirectionVector(options),
                    distance,
                    options
                );
            }
        }

        private static void CreateAngledDimensions(
            DimensionGeometrySnapshot snapshot,
            DimensionOptions options
        ) {
            if (snapshot?.View == null || options == null || !options.HasUserVector) return;
            if (snapshot.ElementEndpoints == null || snapshot.ElementEndpoints.Count == 0) return;

            var pickedVector = GetNormalizedPickedVector(options);
            var sideNormal = GetSelectedAngledSideNormal(options);
            if (pickedVector == null || sideNormal == null) return;

            if (options.IsOverall) {
                var referencePoints = GetAngledDimensionReferencePoints(snapshot, options);
                if (referencePoints.Count < 2) return;

                var dimensionPoints = CreateDimensionPointList(referencePoints);
                var distance = GetAngledDistanceFromReferencePoints(
                    referencePoints,
                    options,
                    GetOffset(options)
                );

                CreateDimensionSet(
                    snapshot.View,
                    dimensionPoints,
                    GetAngledDirectionVector(options),
                    distance,
                    options
                );

                return;
            }

            foreach (var element in snapshot.ElementEndpoints) {
                if (element?.Endpoints == null || element.Endpoints.Count == 0) continue;

                var referencePoints = GetAngledDimensionReferencePoints(snapshot, options, element);
                if (referencePoints.Count < 2) continue;

                var dimensionPoints = CreateDimensionPointList(referencePoints);
                var distance = GetAngledDistanceFromReferencePoints(
                    referencePoints,
                    options,
                    GetOffset(options)
                );

                CreateDimensionSet(
                    snapshot.View,
                    dimensionPoints,
                    GetAngledDirectionVector(options),
                    distance,
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

                var pickedResult = picker.PickPoints(3, prompts);
                var pickedPoints = pickedResult?.Item1;
                var points = pickedPoints?.OfType<TSG.Point>().ToList();

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

        private static bool PickAngledVectorPoints(TSD.DrawingHandler drawingHandler, DimensionOptions options) {
            try {
                var picker = drawingHandler.GetPicker();
                var prompts = new TSD.StringList {
                    "Wskaż pierwszy punkt kierunku wymiarowania",
                    "Wskaż drugi punkt kierunku wymiarowania"
                };

                var pickedResult = picker.PickPoints(2, prompts);
                var pickedPoints = pickedResult?.Item1;
                var points = pickedPoints?.OfType<TSG.Point>().ToList();

                if (points == null || points.Count < 2) return false;

                var firstPoint = new TSG.Point(points[0].X, points[0].Y, 0);
                var secondPoint = new TSG.Point(points[1].X, points[1].Y, 0);

                if (GetDistance2D(firstPoint, secondPoint) < MinimumPickedVectorLengthMillimeters) {
                    MessageBox.Show("Wybrane punkty kierunku są zbyt blisko siebie.");
                    return false;
                }

                options.UserVectorStart = firstPoint;
                options.UserVectorEnd = secondPoint;

                return true;
            }
            catch {
                return false;
            }
        }

        private static TSD.StraightDimensionSet.StraightDimensionSetAttributes CreateStraightDimensionAttributes(
            DimensionOptions options
        ) {
            var attributes = new TSD.StraightDimensionSet.StraightDimensionSetAttributes {
                ShortDimension = TSD.DimensionSetBaseAttributes.ShortDimensionTypes.Outside,
                ExtensionLine = options?.UseShortExtensionLine == true
                    ? TSD.DimensionSetBaseAttributes.ExtensionLineTypes.Yes
                    : TSD.DimensionSetBaseAttributes.ExtensionLineTypes.No
            };

            return attributes;
        }

        private static void CreateDimensionSet(
            TSD.View selectedView,
            TSD.PointList dimensionPoints,
            TSG.Vector directionVector,
            double offsetMillimeters,
            DimensionOptions options
        ) {
            if (selectedView == null || dimensionPoints == null || dimensionPoints.Count < 2 || options == null)
                return;

            if (options.IsCurved) {
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

                return;
            }

            var straightHandler = new TSD.StraightDimensionSetHandler();
            var attributes = CreateStraightDimensionAttributes(options);

            straightHandler.CreateDimensionSet(
                selectedView,
                dimensionPoints,
                directionVector,
                offsetMillimeters,
                attributes
            );
        }

        #endregion

        #region Constants

        private const double FarDimensionOffsetMillimeters = 20.0;
        private const double PerElementDimensionOffsetMillimeters = 100.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;
        private const double DimensionMergeToleranceMillimeters = 1.0;
        private const double MinimumPickedVectorLengthMillimeters = 1.0;

        #endregion

        #region Data Structures

        private enum DimensionType {
            Straight,
            Curved,
            Angled
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
            public bool UseShortExtensionLine { get; set; }
            public TSG.Point UserArcStart { get; set; }
            public TSG.Point UserArcMid { get; set; }
            public TSG.Point UserArcEnd { get; set; }
            public TSG.Point UserVectorStart { get; set; }
            public TSG.Point UserVectorEnd { get; set; }

            public bool HasUserArcPoints => UserArcStart != null && UserArcMid != null && UserArcEnd != null;
            public bool HasUserVector => UserVectorStart != null && UserVectorEnd != null;

            public bool IsCurved => DimensionType == DimensionType.Curved;
            public bool IsAngled => DimensionType == DimensionType.Angled;
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

        private sealed class ProjectedPoint {
            public int Index { get; set; }
            public TSG.Point Point { get; set; }
            public double AxisValue { get; set; }
            public double SideValue { get; set; }
        }

        #endregion
    }
}