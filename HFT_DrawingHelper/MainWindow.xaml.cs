using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        public MainWindow() {
            var myModel = new TSM.Model();

            if (myModel.GetConnectionStatus()) {
                InitializeComponent();
                ModelDrawingLabel.Content = myModel.GetInfo().ModelName.Replace(".db1", "");
            }
            else MessageBox.Show("Keine Verbindung zu Tekla Structures");
        }

        #region Variables

        private const string ViewAttributeName = "#HFT_Kant_Section";
        private const string MarkAttributeName = "#HFT_SECTION_V";

        private const double ZStepMillimeters = 0.5;
        private const int MaximumEmptyStepsPerDirection = 80;
        private const double DuplicateToleranceMillimeters = 0.5;
        private const double MinimumLengthMillimeters = 100.0;

        #endregion

        #region Button Clicks

        #region Draw Edges

        private static void DrawEdgesWithNumbers() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var detectionResult = DetectEdgesFromSelectedView(selectedView);

            if (!detectionResult.HasEdges) {
                MessageBox.Show(detectionResult.ErrorMessage);
                return;
            }

            DrawEdges(selectedView, detectionResult.EdgesByNumber);
            drawing.CommitChanges();
        }

        private sealed class EdgeDetectionResult {
            public string SourceName { get; set; }
            public Dictionary<int, Tuple<TSG.Point, TSG.Point>> EdgesByNumber { get; set; }
            public string ErrorMessage { get; set; }

            public bool HasEdges => EdgesByNumber != null && EdgesByNumber.Count > 0;
        }

        private static EdgeDetectionResult DetectEdgesFromSelectedView(TSD.View selectedView) {
            if (selectedView == null)
                return new EdgeDetectionResult {
                    ErrorMessage = "Nie zaznaczono widoku."
                };

            var modelParts = GetModelPartsFromDrawingView(selectedView);
            if (modelParts == null || modelParts.Count == 0)
                return new EdgeDetectionResult {
                    ErrorMessage = "Nie znaleziono żadnych elementów typu Part na tym widoku."
                };

            var contourEdges = GetContourPlateEdges(selectedView, modelParts);
            if (contourEdges != null && contourEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "ContourPlate",
                    EdgesByNumber = contourEdges
                };

            var loftedEdges = GetLoftedPlateEdges(selectedView, modelParts);
            if (loftedEdges != null && loftedEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "LoftedPlate",
                    EdgesByNumber = loftedEdges
                };

            var beamEdges = GetBeamEdges(selectedView, modelParts);
            if (beamEdges != null && beamEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "Beam",
                    EdgesByNumber = beamEdges
                };

            var foundTypes = string.Join(", ", modelParts
                .Select(part => part.GetType().Name)
                .Distinct()
                .OrderBy(typeName => typeName));

            return new EdgeDetectionResult {
                ErrorMessage =
                    "Nie udało się wyznaczyć krawędzi dla żadnego obsługiwanego typu elementu.\n" +
                    "Znalezione typy na widoku: " + foundTypes
            };
        }

        private void DrawEdgesButton_Click(object sender, RoutedEventArgs e) {
            DrawEdgesWithNumbers();
        }

        #endregion

        #region Add sections

        private static void AddSections(string edgeNumbersInput) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var detectionResult = DetectEdgesFromSelectedView(selectedView);

            if (!detectionResult.HasEdges) {
                MessageBox.Show(detectionResult.ErrorMessage);
                return;
            }

            var requestedGroupNumbers = ParseEdgeNumbers(edgeNumbersInput);

            const double joinToleranceMillimeters = 0.5;
            const double nearStraightAngleDegrees = 170.0;

            var groupsByGroupNumber = BuildNumberedEdgeGroups(
                detectionResult.EdgesByNumber,
                joinToleranceMillimeters,
                nearStraightAngleDegrees
            );

            var sectionEdgesByGroupNumber = groupsByGroupNumber
                .Where(pair => pair.Value?.SectionEdge != null)
                .ToDictionary(pair => pair.Key, pair => pair.Value.SectionEdge);

            var filteredSectionEdges =
                FilterEdgesOrShowMessage(sectionEdgesByGroupNumber, requestedGroupNumbers);
            if (filteredSectionEdges == null) return;

            CreateSectionViewsFromEdges(selectedView, filteredSectionEdges);
            drawing.CommitChanges();
        }

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            AddSections(EdgeNumbersTextBox.Text);
        }

        #endregion

        #region Add dimensions

        private void AddDimensions() {
            var elements = ElementsToDimensionTextBox.Text;
            var drawingHandler = new TSD.DrawingHandler();

            if (drawingHandler.GetConnectionStatus()) {
                var drawing = drawingHandler.GetActiveDrawing();
                if (drawing == null) return;
            }
        }

        private void AddDimensionsButton_Click(object sender, RoutedEventArgs e) {
            AddDimensions();
        }

        #endregion

        #endregion

        #region Helpers

        #region Create

        private static void CreateSectionViewsFromEdges(
            TSD.View baseView,
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber
        ) {
            if (baseView == null) return;
            if (edgesByNumber == null || edgesByNumber.Count == 0) return;

            var (viewAttrs, markAttrs) = GetSectionAttributes();

            const double depthUp = 1.0;
            const double depthDown = 1.0;
            const double sectionLineLengthMm = 300.0;
            const double gap = 10.0;

            var baseBox = baseView.GetAxisAlignedBoundingBox();

            var cursorX = baseBox.UpperRight.X + gap;
            var cursorY = baseBox.UpperRight.Y;

            foreach (var pair in edgesByNumber.OrderBy(x => x.Key)) {
                var edge = pair.Value;
                if (edge == null) continue;

                var edgeA = edge.Item1;
                var edgeB = edge.Item2;

                var mid = new TSG.Point(
                    (edgeA.X + edgeB.X) * 0.5,
                    (edgeA.Y + edgeB.Y) * 0.5,
                    (edgeA.Z + edgeB.Z) * 0.5
                );

                var dx = edgeB.X - edgeA.X;
                var dy = edgeB.Y - edgeA.Y;

                var len = Math.Sqrt(dx * dx + dy * dy);
                if (len < 1e-6) continue;

                var nx = -dy / len;
                var ny = dx / len;

                const double half = sectionLineLengthMm * 0.5;

                var startPoint = new TSG.Point(
                    mid.X - nx * half,
                    mid.Y - ny * half,
                    mid.Z
                );

                var endPoint = new TSG.Point(
                    mid.X + nx * half,
                    mid.Y + ny * half,
                    mid.Z
                );

                ForceSectionLookLeftOrUp(ref startPoint, ref endPoint);

                var insertionPoint = new TSG.Point(cursorX, cursorY, baseView.Origin.Z);

                var ok = TSD.View.CreateSectionView(
                    baseView,
                    startPoint,
                    endPoint,
                    insertionPoint,
                    depthUp,
                    depthDown,
                    viewAttrs,
                    markAttrs,
                    out var sectionView,
                    out var sectionMark
                );

                if (!ok || sectionView == null || sectionMark == null) continue;

                sectionView.Modify();
                sectionMark.Modify();

                var secBox = sectionView.GetAxisAlignedBoundingBox();

                var deltaX = cursorX - secBox.LowerLeft.X;
                var deltaY = cursorY - secBox.UpperLeft.Y;

                sectionView.Origin = new TSG.Point(
                    sectionView.Origin.X + deltaX,
                    sectionView.Origin.Y + deltaY,
                    sectionView.Origin.Z
                );

                sectionView.Modify();

                secBox = sectionView.GetAxisAlignedBoundingBox();
                cursorX = secBox.UpperRight.X + gap;
            }
        }

        private static void CreateSingleSectionView(TSD.View baseView) {
            var (viewAttrs, markAttrs) = GetSectionAttributes();

            var startPoint = new TSG.Point(baseView.Origin.X + 100, baseView.Origin.Y + 100, baseView.Origin.Z);
            var endPoint = new TSG.Point(baseView.Origin.X + 100, baseView.Origin.Y + 200, baseView.Origin.Z);
            var insertionPoint = new TSG.Point(baseView.Origin.X, baseView.Origin.Y, baseView.Origin.Z);

            const double depthUp = 100.0;
            const double depthDown = 100.0;

            var ok = TSD.View.CreateSectionView(
                baseView,
                startPoint,
                endPoint,
                insertionPoint,
                depthUp,
                depthDown,
                viewAttrs,
                markAttrs,
                out var sectionView,
                out var sectionMark
            );

            if (!ok || sectionView == null || sectionMark == null) return;

            sectionView.Modify();
            sectionMark.Modify();
        }

        #endregion

        #region Gets

        #region ContourPlate

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetContourPlateEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;

            const double minLengthMm = 100.0;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var contourPlate = FindFirstContourPlate(modelParts);
                if (contourPlate == null) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(contourPlate.GetCoordinateSystem()));

                var edgesInPlatePlane = GetContourPlateEdgesInPlane(contourPlate);

                if (edgesInPlatePlane == null || edgesInPlatePlane.Count == 0) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem));

                var edgesInViewPlane = TransformEdgesBetweenCoordinateSystems(
                    contourPlate.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    edgesInPlatePlane
                );

                var rotated = RotateEdgesAroundCenter2D(edgesInViewPlane, 90.0);
                var shiftedDown = TranslateEdgesDownByHalfHeight(rotated);

                var number = 0;

                for (var i = 0; i < shiftedDown.Count; i++) {
                    var start = shiftedDown[i].A;
                    var end = shiftedDown[i].B;

                    var dx = end.X - start.X;
                    var dy = end.Y - start.Y;
                    var dz = end.Z - start.Z;

                    var length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (length <= minLengthMm) continue;

                    number++;

                    var a = new TSG.Point(start.X, start.Y, 0);
                    var b = new TSG.Point(end.X, end.Y, 0);

                    edgesByNumber[number] = Tuple.Create(a, b);
                }

                return edgesByNumber;
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }


        private static List<(TSG.Point A, TSG.Point B)> GetContourPlateEdgesInPlane(TSM.ContourPlate contourPlate) {
            var edges = new List<(TSG.Point A, TSG.Point B)>();
            if (contourPlate == null) return edges;

            var solid = contourPlate.GetSolid();
            if (solid == null) return edges;

            var points = new List<TSG.Point>();

            var enumerator = solid.GetAllIntersectionPoints(
                new TSG.Point(0, 0, 0),
                new TSG.Point(1, 0, 0),
                new TSG.Point(0, 1, 0)
            );

            while (enumerator.MoveNext()) {
                if (!(enumerator.Current is TSG.Point point)) continue;

                points.Add(new TSG.Point(point.X, point.Y, 0));
            }

            points = RemoveNearDuplicates(points, 0.5);
            points = SortByAngle(points);

            if (points.Count < 3) return edges;

            for (var i = 0; i < points.Count; i++) {
                var a = points[i];
                var b = points[(i + 1) % points.Count];
                edges.Add((a, b));
            }

            return edges;
        }

        #endregion

        #region LoftedPlate

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetLoftedPlateEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;


            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var loftedPlate = FindFirstLoftedPlate(modelParts);
                if (loftedPlate == null) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(loftedPlate.GetCoordinateSystem())
                );

                var sampledPointsInPlatePlane = GetLoftedPlateSamplePointsByAdaptiveZSweep(loftedPlate);

                if (sampledPointsInPlatePlane == null || sampledPointsInPlatePlane.Count == 0) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem)
                );

                var sampledPointsInViewPlane = TransformPointsBetweenCoordinateSystems(
                    loftedPlate.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    sampledPointsInPlatePlane
                );

                var edgesInViewPlane = BuildLoftedPlateEnvelopeEdgesInView(sampledPointsInViewPlane);

                if (edgesInViewPlane == null || edgesInViewPlane.Count == 0) return edgesByNumber;

                var number = 0;

                for (var index = 0; index < edgesInViewPlane.Count; index++) {
                    var startPoint = edgesInViewPlane[index].A;
                    var endPoint = edgesInViewPlane[index].B;

                    var dx = endPoint.X - startPoint.X;
                    var dy = endPoint.Y - startPoint.Y;
                    var dz = endPoint.Z - startPoint.Z;

                    var length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (length <= MinimumLengthMillimeters)
                        continue;

                    number++;
                    edgesByNumber[number] = Tuple.Create(
                        new TSG.Point(startPoint.X, startPoint.Y, 0),
                        new TSG.Point(endPoint.X, endPoint.Y, 0)
                    );
                }

                return edgesByNumber;
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static List<TSG.Point> GetLoftedPlateSamplePointsByAdaptiveZSweep(TSM.LoftedPlate loftedPlate) {
            var points = new List<TSG.Point>();
            if (loftedPlate == null) return points;

            var solid = loftedPlate.GetSolid();
            if (solid == null) return points;

            var minimumPoint = solid.MinimumPoint;
            var maximumPoint = solid.MaximumPoint;

            var centerZ = (minimumPoint.Z + maximumPoint.Z) * 0.5;

            var upwardEmptySteps = 0;
            var downwardEmptySteps = 0;

            var offset = 0.0;
            var firstIteration = true;

            while (upwardEmptySteps < MaximumEmptyStepsPerDirection ||
                   downwardEmptySteps < MaximumEmptyStepsPerDirection) {
                if (firstIteration) {
                    var centerPoints = GetIntersectionPointsAtLocalZ(solid, centerZ);

                    if (centerPoints.Count > 0) {
                        points.AddRange(centerPoints);
                        upwardEmptySteps = 0;
                        downwardEmptySteps = 0;
                    }
                    else {
                        upwardEmptySteps++;
                        downwardEmptySteps++;
                    }

                    firstIteration = false;
                    offset += ZStepMillimeters;
                    continue;
                }

                if (upwardEmptySteps < MaximumEmptyStepsPerDirection) {
                    var upperZ = centerZ + offset;
                    var upperPoints = GetIntersectionPointsAtLocalZ(solid, upperZ);

                    if (upperPoints.Count > 0) {
                        points.AddRange(upperPoints);
                        upwardEmptySteps = 0;
                    }
                    else
                        upwardEmptySteps++;
                }

                if (downwardEmptySteps < MaximumEmptyStepsPerDirection) {
                    var lowerZ = centerZ - offset;
                    var lowerPoints = GetIntersectionPointsAtLocalZ(solid, lowerZ);

                    if (lowerPoints.Count > 0) {
                        points.AddRange(lowerPoints);
                        downwardEmptySteps = 0;
                    }
                    else
                        downwardEmptySteps++;
                }

                offset += ZStepMillimeters;
            }

            return RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
        }

        private static List<TSG.Point> TransformPointsBetweenCoordinateSystems(
            TSG.CoordinateSystem fromCoordinateSystem,
            TSG.CoordinateSystem toCoordinateSystem,
            List<TSG.Point> points
        ) {
            var result = new List<TSG.Point>();
            if (points == null || points.Count == 0) return result;

            var fromToGlobal = TSG.MatrixFactory.FromCoordinateSystem(fromCoordinateSystem);
            var globalToTarget = TSG.MatrixFactory.ToCoordinateSystem(toCoordinateSystem);

            result.AddRange(from t in points
                select fromToGlobal.Transform(t)
                into globalPoint
                select globalToTarget.Transform(globalPoint)
                into targetPoint
                select new TSG.Point(targetPoint.X, targetPoint.Y, 0));

            return result;
        }

        private static List<(TSG.Point A, TSG.Point B)> BuildLoftedPlateEnvelopeEdgesInView(List<TSG.Point> points) {
            var edges = new List<(TSG.Point A, TSG.Point B)>();
            if (points == null || points.Count < 3) return edges;

            const double duplicateToleranceMillimeters = 0.5;
            const double minimumSegmentLengthMillimeters = 5.0;
            const int targetBinCount = 60;
            const int smoothingWindowRadius = 2;
            const int edgeLockBinCount = 3;

            var uniquePoints = RemoveNearDuplicates(points, duplicateToleranceMillimeters);
            if (uniquePoints.Count < 3) return edges;

            var minimumX = uniquePoints.Min(point => point.X);
            var maximumX = uniquePoints.Max(point => point.X);

            if (maximumX - minimumX < 1e-6)
                return edges;

            var binWidth = Math.Max(2.0, (maximumX - minimumX) / targetBinCount);

            var bins = new SortedDictionary<int, List<TSG.Point>>();

            foreach (var point in uniquePoints) {
                var binIndex = (int)Math.Floor((point.X - minimumX) / binWidth);

                if (!bins.ContainsKey(binIndex))
                    bins[binIndex] = new List<TSG.Point>();

                bins[binIndex].Add(point);
            }

            var orderedBins = bins
                .OrderBy(pair => pair.Key)
                .Select(pair => pair.Value.OrderBy(point => point.Y).ThenBy(point => point.X).ToList())
                .Where(bucket => bucket.Count > 0)
                .ToList();

            if (orderedBins.Count < 2)
                return edges;

            var upperChain = new List<TSG.Point>();
            var lowerChain = new List<TSG.Point>();

            foreach (var bucket in orderedBins) {
                if (bucket.Count == 0)
                    continue;

                var lowerPointIndex = (int)Math.Floor((bucket.Count - 1) * 0.15);
                var upperPointIndex = (int)Math.Ceiling((bucket.Count - 1) * 0.85);

                if (lowerPointIndex < 0) lowerPointIndex = 0;
                if (upperPointIndex < 0) upperPointIndex = 0;
                if (lowerPointIndex >= bucket.Count) lowerPointIndex = bucket.Count - 1;
                if (upperPointIndex >= bucket.Count) upperPointIndex = bucket.Count - 1;

                var lowerPoint = bucket[lowerPointIndex];
                var upperPoint = bucket[upperPointIndex];

                lowerChain.Add(new TSG.Point(lowerPoint.X, lowerPoint.Y, 0));
                upperChain.Add(new TSG.Point(upperPoint.X, upperPoint.Y, 0));
            }

            upperChain = SmoothPolylineByMedianY(upperChain, smoothingWindowRadius);
            lowerChain = SmoothPolylineByMedianY(lowerChain, smoothingWindowRadius);

            if (upperChain.Count < 2 || lowerChain.Count < 2)
                return edges;

            var leftLockedX = uniquePoints
                .OrderBy(point => point.X)
                .Take(Math.Min(20, uniquePoints.Count))
                .Average(point => point.X);

            var rightLockedX = uniquePoints
                .OrderByDescending(point => point.X)
                .Take(Math.Min(20, uniquePoints.Count))
                .Average(point => point.X);

            var leftLockCount = Math.Min(edgeLockBinCount, Math.Min(upperChain.Count, lowerChain.Count));
            var rightLockCount = Math.Min(edgeLockBinCount, Math.Min(upperChain.Count, lowerChain.Count));

            for (var index = 0; index < leftLockCount; index++) {
                upperChain[index] = new TSG.Point(leftLockedX, upperChain[index].Y, 0);
                lowerChain[index] = new TSG.Point(leftLockedX, lowerChain[index].Y, 0);
            }

            for (var index = 0; index < rightLockCount; index++) {
                var upperIndex = upperChain.Count - 1 - index;
                var lowerIndex = lowerChain.Count - 1 - index;

                upperChain[upperIndex] = new TSG.Point(rightLockedX, upperChain[upperIndex].Y, 0);
                lowerChain[lowerIndex] = new TSG.Point(rightLockedX, lowerChain[lowerIndex].Y, 0);
            }

            upperChain = RemoveNearDuplicates(upperChain, duplicateToleranceMillimeters);
            lowerChain = RemoveNearDuplicates(lowerChain, duplicateToleranceMillimeters);

            upperChain = RemoveVeryShortSegmentsFromOpenPolyline(upperChain, minimumSegmentLengthMillimeters);
            lowerChain = RemoveVeryShortSegmentsFromOpenPolyline(lowerChain, minimumSegmentLengthMillimeters);

            if (upperChain.Count < 2 || lowerChain.Count < 2)
                return edges;

            var boundaryPoints = upperChain.Select(t => new TSG.Point(t.X, t.Y, 0)).ToList();

            for (var index = lowerChain.Count - 1; index >= 0; index--)
                boundaryPoints.Add(new TSG.Point(lowerChain[index].X, lowerChain[index].Y, 0));

            boundaryPoints = RemoveNearDuplicates(boundaryPoints, duplicateToleranceMillimeters);
            boundaryPoints = RemoveVeryShortSegmentsFromClosedPolyline(boundaryPoints, minimumSegmentLengthMillimeters);
            boundaryPoints = RemoveCollinearPointsFromClosedPolyline(boundaryPoints, duplicateToleranceMillimeters);

            if (boundaryPoints.Count < 3)
                return edges;

            for (var index = 0; index < boundaryPoints.Count; index++) {
                var startPoint = boundaryPoints[index];
                var endPoint = boundaryPoints[(index + 1) % boundaryPoints.Count];

                if (ComputeDistance2D(startPoint, endPoint) <= duplicateToleranceMillimeters)
                    continue;

                edges.Add((
                    new TSG.Point(startPoint.X, startPoint.Y, 0),
                    new TSG.Point(endPoint.X, endPoint.Y, 0)
                ));
            }

            return edges;
        }

        private static List<TSG.Point> SmoothPolylineByMedianY(List<TSG.Point> points, int windowRadius) {
            var result = new List<TSG.Point>();
            if (points == null || points.Count == 0) return result;

            for (var index = 0; index < points.Count; index++) {
                var fromIndex = Math.Max(0, index - windowRadius);
                var toIndex = Math.Min(points.Count - 1, index + windowRadius);

                var yValues = new List<double>();

                for (var innerIndex = fromIndex; innerIndex <= toIndex; innerIndex++)
                    yValues.Add(points[innerIndex].Y);

                yValues.Sort();

                var medianY = yValues[yValues.Count / 2];
                result.Add(new TSG.Point(points[index].X, medianY, 0));
            }

            return result;
        }

        private static List<TSG.Point> RemoveVeryShortSegmentsFromOpenPolyline(
            List<TSG.Point> points,
            double minimumSegmentLengthMillimeters
        ) {
            if (points == null || points.Count < 2)
                return points ?? new List<TSG.Point>();

            var result = new List<TSG.Point>(points);
            var changed = true;

            while (changed && result.Count >= 2) {
                changed = false;

                for (var index = 0; index < result.Count - 1; index++) {
                    if (ComputeDistance2D(result[index], result[index + 1]) >= minimumSegmentLengthMillimeters)
                        continue;

                    result.RemoveAt(index + 1);
                    changed = true;
                    break;
                }
            }

            return result;
        }

        private static List<TSG.Point> RemoveVeryShortSegmentsFromClosedPolyline(
            List<TSG.Point> points,
            double minimumSegmentLengthMillimeters
        ) {
            if (points == null || points.Count < 3)
                return points ?? new List<TSG.Point>();

            var result = new List<TSG.Point>(points);
            var changed = true;

            while (changed && result.Count >= 3) {
                changed = false;

                for (var index = 0; index < result.Count; index++) {
                    var nextIndex = (index + 1) % result.Count;

                    if (ComputeDistance2D(result[index], result[nextIndex]) >= minimumSegmentLengthMillimeters)
                        continue;

                    result.RemoveAt(nextIndex);
                    changed = true;
                    break;
                }
            }

            return result;
        }

        private static List<TSG.Point> RemoveCollinearPointsFromClosedPolyline(
            List<TSG.Point> points,
            double toleranceMillimeters
        ) {
            if (points == null || points.Count < 3)
                return points ?? new List<TSG.Point>();

            var result = new List<TSG.Point>(points);
            var changed = true;

            while (changed && result.Count >= 3) {
                changed = false;

                for (var index = 0; index < result.Count; index++) {
                    var previousPoint = result[(index - 1 + result.Count) % result.Count];
                    var currentPoint = result[index];
                    var nextPoint = result[(index + 1) % result.Count];

                    var cross = Math.Abs(
                        (currentPoint.X - previousPoint.X) * (nextPoint.Y - previousPoint.Y) -
                        (currentPoint.Y - previousPoint.Y) * (nextPoint.X - previousPoint.X)
                    );

                    if (cross > toleranceMillimeters)
                        continue;

                    result.RemoveAt(index);
                    changed = true;
                    break;
                }
            }

            return result;
        }

        #endregion

        #region Beam

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetBeamEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var beam = FindFirstBeam(modelParts);
                if (beam == null) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(beam.GetCoordinateSystem())
                );

                var sampledPointsInBeamPlane = GetBeamSamplePointsByAdaptiveZSweep(beam);

                if (sampledPointsInBeamPlane == null || sampledPointsInBeamPlane.Count == 0) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem)
                );

                var sampledPointsInViewPlane = TransformPointsBetweenCoordinateSystems(
                    beam.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    sampledPointsInBeamPlane
                );
                var edgesInViewPlane = BuildBeamEnvelopeEdgesInView(sampledPointsInViewPlane);

                if (edgesInViewPlane == null || edgesInViewPlane.Count == 0) return edgesByNumber;

                var number = 0;

                for (var index = 0; index < edgesInViewPlane.Count; index++) {
                    var startPoint = edgesInViewPlane[index].A;
                    var endPoint = edgesInViewPlane[index].B;

                    var dx = endPoint.X - startPoint.X;
                    var dy = endPoint.Y - startPoint.Y;
                    var dz = endPoint.Z - startPoint.Z;

                    var length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (length <= MinimumLengthMillimeters)
                        continue;

                    number++;
                    edgesByNumber[number] = Tuple.Create(
                        new TSG.Point(startPoint.X, startPoint.Y, 0),
                        new TSG.Point(endPoint.X, endPoint.Y, 0)
                    );
                }

                return edgesByNumber;
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static List<TSG.Point> GetBeamSamplePointsByAdaptiveZSweep(TSM.Beam beam) {
            var points = new List<TSG.Point>();
            if (beam == null) return points;

            var solid = beam.GetSolid();
            if (solid == null) return points;

            var minimumPoint = solid.MinimumPoint;
            var maximumPoint = solid.MaximumPoint;

            var centerZ = (minimumPoint.Z + maximumPoint.Z) * 0.5;

            var upwardEmptySteps = 0;
            var downwardEmptySteps = 0;

            var offset = 0.0;
            var firstIteration = true;

            while (upwardEmptySteps < MaximumEmptyStepsPerDirection ||
                   downwardEmptySteps < MaximumEmptyStepsPerDirection) {
                if (firstIteration) {
                    var centerPoints = GetIntersectionPointsAtLocalZ(solid, centerZ);

                    if (centerPoints.Count > 0) {
                        points.AddRange(centerPoints);
                        upwardEmptySteps = 0;
                        downwardEmptySteps = 0;
                    }
                    else {
                        upwardEmptySteps++;
                        downwardEmptySteps++;
                    }

                    firstIteration = false;
                    offset += ZStepMillimeters;
                    continue;
                }

                if (upwardEmptySteps < MaximumEmptyStepsPerDirection) {
                    var upperZ = centerZ + offset;
                    var upperPoints = GetIntersectionPointsAtLocalZ(solid, upperZ);

                    if (upperPoints.Count > 0) {
                        points.AddRange(upperPoints);
                        upwardEmptySteps = 0;
                    }
                    else
                        upwardEmptySteps++;
                }

                if (downwardEmptySteps < MaximumEmptyStepsPerDirection) {
                    var lowerZ = centerZ - offset;
                    var lowerPoints = GetIntersectionPointsAtLocalZ(solid, lowerZ);

                    if (lowerPoints.Count > 0) {
                        points.AddRange(lowerPoints);
                        downwardEmptySteps = 0;
                    }
                    else
                        downwardEmptySteps++;
                }

                offset += ZStepMillimeters;
            }

            return RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
        }

        private static List<(TSG.Point A, TSG.Point B)> BuildBeamEnvelopeEdgesInView(List<TSG.Point> points) {
            var edges = new List<(TSG.Point A, TSG.Point B)>();
            if (points == null || points.Count < 3) return edges;

            var uniquePoints = RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
            if (uniquePoints.Count < 3) return edges;

            var hull = BuildConvexHull2D(uniquePoints);
            if (hull.Count < 3) return edges;

            for (var index = 0; index < hull.Count; index++) {
                var startPoint = hull[index];
                var endPoint = hull[(index + 1) % hull.Count];

                if (ComputeDistance2D(startPoint, endPoint) <= DuplicateToleranceMillimeters)
                    continue;

                edges.Add((
                    new TSG.Point(startPoint.X, startPoint.Y, 0),
                    new TSG.Point(endPoint.X, endPoint.Y, 0)
                ));
            }

            return edges;
        }

        private static List<TSG.Point> BuildConvexHull2D(List<TSG.Point> points) {
            if (points == null || points.Count < 3)
                return points ?? new List<TSG.Point>();

            var sortedPoints = points
                .OrderBy(point => point.X)
                .ThenBy(point => point.Y)
                .ToList();

            var lowerHull = new List<TSG.Point>();
            foreach (var point in sortedPoints) {
                while (lowerHull.Count >= 2 &&
                       Cross(lowerHull[lowerHull.Count - 2], lowerHull[lowerHull.Count - 1], point) <= 0)
                    lowerHull.RemoveAt(lowerHull.Count - 1);

                lowerHull.Add(point);
            }

            var upperHull = new List<TSG.Point>();
            for (var index = sortedPoints.Count - 1; index >= 0; index--) {
                var point = sortedPoints[index];

                while (upperHull.Count >= 2 &&
                       Cross(upperHull[upperHull.Count - 2], upperHull[upperHull.Count - 1], point) <= 0)
                    upperHull.RemoveAt(upperHull.Count - 1);

                upperHull.Add(point);
            }

            lowerHull.RemoveAt(lowerHull.Count - 1);
            upperHull.RemoveAt(upperHull.Count - 1);

            return lowerHull.Concat(upperHull).ToList();
        }

        private static double Cross(TSG.Point firstPoint, TSG.Point secondPoint, TSG.Point thirdPoint) {
            return (secondPoint.X - firstPoint.X) * (thirdPoint.Y - firstPoint.Y) -
                   (secondPoint.Y - firstPoint.Y) * (thirdPoint.X - firstPoint.X);
        }

        #endregion

        /// Pobiera zaznaczony widok rysunkowy. Jeśli nie ma zaznaczenia lub zaznaczenie nie jest widokiem, pokazuje komunikat i zwraca null.
        private static TSD.View GetSelectedViewOrShowMessage(TSD.DrawingHandler drawingHandler) {
            var selector = drawingHandler.GetDrawingObjectSelector();
            var selected = selector.GetSelected();

            if (selected == null) {
                MessageBox.Show("Nie zaznaczono żadnego obiektu.");
                return null;
            }

            TSD.View selectedView = null;

            selected.SelectInstances = false;
            while (selected.MoveNext()) {
                selectedView = selected.Current as TSD.View;
                if (selectedView != null) break;
            }

            if (selectedView == null) {
                MessageBox.Show("Zaznacz widok na rysunku, a potem uruchom funkcję.");
                return null;
            }

            return selectedView;
        }

        /// Pobiera unikalne elementy typu Part powiązane z obiektami rysunkowymi na danym widoku.
        private static List<TSM.Part> GetModelPartsFromDrawingView(TSD.View drawingView) {
            var modelParts = new List<TSM.Part>();
            var addedModelIdentifiers = new HashSet<string>();

            var model = new TSM.Model();

            var drawingObjectEnumerator = drawingView.GetAllObjects();
            if (drawingObjectEnumerator == null) return modelParts;

            drawingObjectEnumerator.SelectInstances = true;

            while (drawingObjectEnumerator.MoveNext()) {
                if (!(drawingObjectEnumerator.Current is TSD.DrawingObject drawingObject)) continue;

                object modelIdentifierObject;

                try {
                    modelIdentifierObject = ((dynamic)drawingObject).ModelIdentifier;
                }
                catch {
                    continue;
                }

                if (modelIdentifierObject == null) continue;

                var identifierString = modelIdentifierObject.ToString();
                if (addedModelIdentifiers.Contains(identifierString)) continue;

                try {
                    var modelObject = model.SelectModelObject((TS.Identifier)modelIdentifierObject);
                    if (!(modelObject is TSM.Part modelPart)) continue;

                    addedModelIdentifiers.Add(identifierString);
                    modelParts.Add(modelPart);
                }
                catch {
                    // ignored
                }
            }

            return modelParts;
        }

        private static List<TSG.Point> GetIntersectionPointsAtLocalZ(TSM.Solid solid, double zValue) {
            var points = new List<TSG.Point>();
            if (solid == null) return points;

            var enumerator = solid.GetAllIntersectionPoints(
                new TSG.Point(0, 0, zValue),
                new TSG.Point(1, 0, zValue),
                new TSG.Point(0, 1, zValue)
            );

            while (enumerator.MoveNext()) {
                if (!(enumerator.Current is TSG.Point point))
                    continue;

                points.Add(new TSG.Point(point.X, point.Y, point.Z));
            }

            return points;
        }

        /// Pobiera atrybuty dla widoku i marku przekroju zdefiniowane w Tekla Structures. Zakłada, że atrybuty o podanych nazwach istnieją.
        private static (TSD.View.ViewAttributes, TSD.SectionMarkBase.SectionMarkAttributes) GetSectionAttributes() {
            var view = new TSD.View.ViewAttributes(ViewAttributeName);
            var mark = new TSD.SectionMarkBase.SectionMarkAttributes();
            mark.LoadAttributes(MarkAttributeName);

            return (view, mark);
        }

        #endregion

        #region Other

        #region Draw Elements

        /// Rysuje linie krawędzi i numery obok nich na widoku. Zakłada, że krawędzie są już w układzie współrzędnych widoku (Z=0).
        private static void DrawEdges(TSD.View view, Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber) {
            if (view == null) return;
            if (edgesByNumber == null || edgesByNumber.Count == 0) return;

            const double nearStraightAngleDegrees = 170.0;
            const double joinToleranceMillimeters = 0.5;

            const double numberOffsetMillimeters = 20.0;

            var lineAttributes = new TSD.Line.LineAttributes {
                Line = new TSD.LineTypeAttributes(TSD.LineTypes.SolidLine, TSD.DrawingColors.Red)
            };

            var numberedGroups =
                BuildNumberedEdgeGroups(edgesByNumber, joinToleranceMillimeters, nearStraightAngleDegrees);

            foreach (var pair in numberedGroups.OrderBy(p => p.Key)) {
                var group = pair.Value;
                if (group == null) continue;

                if (!group.IsPolyline) {
                    var singleSegment = group.EdgeSegments[0];

                    new TSD.Line(view, singleSegment.StartPoint, singleSegment.EndPoint, lineAttributes).Insert();

                    var numberPoint = ComputeTextInsertionPointForSegment(
                        singleSegment.StartPoint,
                        singleSegment.EndPoint,
                        numberOffsetMillimeters
                    );

                    new TSD.Text(view, numberPoint, group.GroupNumber.ToString()).Insert();
                    continue;
                }

                var polylinePointList = new TSD.PointList();
                foreach (var polylinePoint in group.PolylinePoints)
                    polylinePointList.Add(new TSG.Point(polylinePoint.X, polylinePoint.Y, 0));

                var polyline = new TSD.Polyline(view, polylinePointList);
                polyline.Insert();

                var polylineNumberPoint =
                    ComputeTextInsertionPointForPolyline(group.PolylinePoints, numberOffsetMillimeters);
                new TSD.Text(view, polylineNumberPoint, group.GroupNumber.ToString()).Insert();
            }
        }

        #endregion

        #region Find Elements

        /// Szuka pierwszego elementu typu ContourPlate w liście. Jeśli nie znajdzie, zwraca null.
        private static TSM.ContourPlate FindFirstContourPlate(List<TSM.Part> parts) {
            if (parts == null || parts.Count == 0) return null;

            foreach (var part in parts)
                if (part is TSM.ContourPlate contourPlate)
                    return contourPlate;

            return null;
        }

        /// Szuka pierwszego elementu typu LoftedPlate w liście. Jeśli nie znajdzie, zwraca null.
        private static TSM.LoftedPlate FindFirstLoftedPlate(List<TSM.Part> parts) {
            if (parts == null || parts.Count == 0) return null;

            foreach (var part in parts)
                if (part is TSM.LoftedPlate loftedPlate)
                    return loftedPlate;

            return null;
        }

        /// Szuka pierwszego elementu typu Beam w liście. Jeśli nie znajdzie, zwraca null.
        private static TSM.Beam FindFirstBeam(List<TSM.Part> parts) {
            if (parts == null || parts.Count == 0) return null;

            foreach (var part in parts)
                if (part is TSM.Beam beam)
                    return beam;

            return null;
        }

        #endregion

        #region Move Elemets

        /// Rotuje krawędzie o zadany kąt wokół ich środka w 2D. Po rotacji normalizuje pozycję, aby minimalne X i Y były takie same jak przed rotacją.
        private static List<(TSG.Point A, TSG.Point B)> RotateEdgesAroundCenter2D(
            List<(TSG.Point A, TSG.Point B)> edges,
            double angleDegrees
        ) {
            if (edges == null || edges.Count == 0)
                return new List<(TSG.Point A, TSG.Point B)>();

            var beforePts = new List<TSG.Point>(edges.Count * 2);
            for (var i = 0; i < edges.Count; i++) {
                beforePts.Add(edges[i].A);
                beforePts.Add(edges[i].B);
            }

            var beforeMinX = beforePts.Min(p => p.X);
            var beforeMinY = beforePts.Min(p => p.Y);

            var center = ComputeEdgesCenter2D(edges);

            var rad = angleDegrees * Math.PI / 180.0;
            var cos = Math.Cos(rad);
            var sin = Math.Sin(rad);

            var rotated = new List<(TSG.Point A, TSG.Point B)>(edges.Count);
            for (var i = 0; i < edges.Count; i++)
                rotated.Add((
                    RotatePointAroundCenter2D(edges[i].A, center, cos, sin),
                    RotatePointAroundCenter2D(edges[i].B, center, cos, sin)
                ));

            var afterPts = new List<TSG.Point>(rotated.Count * 2);
            for (var i = 0; i < rotated.Count; i++) {
                afterPts.Add(rotated[i].A);
                afterPts.Add(rotated[i].B);
            }

            var afterMinX = afterPts.Min(p => p.X);
            var afterMinY = afterPts.Min(p => p.Y);

            var normalizedAngle = (angleDegrees % 360.0 + 360.0) % 360.0;

            double deltaX;
            double deltaY;

            if (Math.Abs(normalizedAngle - 90.0) < 1e-6 || Math.Abs(normalizedAngle - 270.0) < 1e-6) {
                deltaX = beforeMinY - afterMinX;
                deltaY = beforeMinX - afterMinY;
            }
            else {
                deltaX = beforeMinX - afterMinX;
                deltaY = beforeMinY - afterMinY;
            }

            for (var i = 0; i < rotated.Count; i++) {
                var a = rotated[i].A;
                var b = rotated[i].B;

                rotated[i] = (
                    new TSG.Point(a.X + deltaX, a.Y + deltaY, 0),
                    new TSG.Point(b.X + deltaX, b.Y + deltaY, 0)
                );
            }

            return rotated;
        }

        /// Przesuwa krawędzie w dół o połowę wysokości, aby były bardziej "w środku" względem pierwotnej pozycji.
        private static List<(TSG.Point A, TSG.Point B)> TranslateEdgesDownByHalfHeight(
            List<(TSG.Point A, TSG.Point B)> edges
        ) {
            if (edges == null || edges.Count == 0) return edges;

            var pts = new List<TSG.Point>(edges.Count * 2);
            for (var i = 0; i < edges.Count; i++) {
                pts.Add(edges[i].A);
                pts.Add(edges[i].B);
            }

            var minY = pts.Min(p => p.Y);
            var maxY = pts.Max(p => p.Y);
            var halfHeight = (maxY - minY) * 0.5;

            var result = new List<(TSG.Point A, TSG.Point B)>(edges.Count);
            for (var i = 0; i < edges.Count; i++) {
                var a = edges[i].A;
                var b = edges[i].B;

                result.Add((
                    new TSG.Point(a.X, a.Y - halfHeight, 0),
                    new TSG.Point(b.X, b.Y - halfHeight, 0)
                ));
            }

            return result;
        }

        /// Rotuje punkt p wokół punktu c o kąt określony przez cos i sin. Zakłada, że punkty są w 2D (Z=0) i wynik również ma Z=0.
        private static TSG.Point RotatePointAroundCenter2D(
            TSG.Point p,
            TSG.Point c,
            double cos,
            double sin
        ) {
            var dx = p.X - c.X;
            var dy = p.Y - c.Y;

            var rx = c.X + dx * cos - dy * sin;
            var ry = c.Y + dx * sin + dy * cos;

            return new TSG.Point(rx, ry, 0);
        }

        #endregion

        #region Numer Edges

        /// Parsuje ciąg znaków, który może zawierać pojedyncze numery (np. "3"), zakresy (np. "5-7") oraz różne separatory (przecinki, średniki, spacje).
        /// Zwraca zbiór unikalnych numerów krawędzi.
        private static HashSet<int> ParseEdgeNumbers(string input) {
            var result = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(input)) return result;

            var tokens = input
                .Replace(";", ",")
                .Replace(" ", ",")
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => t.Length > 0);

            foreach (var token in tokens)
                if (token.Contains("-")) {
                    var parts = token.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(p => p.Trim())
                        .ToArray();

                    if (parts.Length != 2) continue;

                    if (!int.TryParse(parts[0], out var start)) continue;
                    if (!int.TryParse(parts[1], out var end)) continue;

                    if (start > end) (start, end) = (end, start);

                    for (var i = start; i <= end; i++) result.Add(i);
                }
                else {
                    if (int.TryParse(token, out var n)) result.Add(n);
                }

            return result;
        }

        #endregion

        private sealed class EdgeSegment {
            public int EdgeNumber { get; set; }
            public TSG.Point StartPoint { get; set; }
            public TSG.Point EndPoint { get; set; }
        }

        private sealed class PolylineGroup {
            public List<EdgeSegment> EdgeSegments { get; } = new List<EdgeSegment>();
            public List<TSG.Point> PolylinePoints { get; } = new List<TSG.Point>();
        }

        private sealed class NumberedEdgeGroup {
            public int GroupNumber { get; set; }
            public bool IsPolyline { get; set; }
            public List<EdgeSegment> EdgeSegments { get; } = new List<EdgeSegment>();
            public List<TSG.Point> PolylinePoints { get; } = new List<TSG.Point>();
            public Tuple<TSG.Point, TSG.Point> SectionEdge { get; set; }
        }

        private static Dictionary<int, NumberedEdgeGroup> BuildNumberedEdgeGroups(
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber,
            double joinToleranceMillimeters,
            double nearStraightAngleDegrees
        ) {
            var result = new Dictionary<int, NumberedEdgeGroup>();
            if (edgesByNumber == null || edgesByNumber.Count == 0) return result;

            var orderedEdges = edgesByNumber
                .OrderBy(edgePair => edgePair.Key)
                .Select(edgePair => new EdgeSegment {
                    EdgeNumber = edgePair.Key,
                    StartPoint = new TSG.Point(edgePair.Value.Item1.X, edgePair.Value.Item1.Y, 0),
                    EndPoint = new TSG.Point(edgePair.Value.Item2.X, edgePair.Value.Item2.Y, 0)
                })
                .ToList();

            var orderedConnectedEdges = BuildOrderedConnectedEdges(orderedEdges, joinToleranceMillimeters);
            var polylineGroups = SplitEdgesIntoPolylineGroupsByAngle(orderedConnectedEdges, nearStraightAngleDegrees);

            var groupNumber = 0;

            foreach (var polylineGroup in polylineGroups) {
                if (polylineGroup == null) continue;
                if (polylineGroup.EdgeSegments == null || polylineGroup.EdgeSegments.Count == 0) continue;

                groupNumber++;

                var numberedGroup = new NumberedEdgeGroup {
                    GroupNumber = groupNumber,
                    IsPolyline = polylineGroup.EdgeSegments.Count > 1
                };

                foreach (var edgeSegment in polylineGroup.EdgeSegments) numberedGroup.EdgeSegments.Add(edgeSegment);

                foreach (var polylinePoint in polylineGroup.PolylinePoints)
                    numberedGroup.PolylinePoints.Add(polylinePoint);

                if (!numberedGroup.IsPolyline) {
                    var singleSegment = numberedGroup.EdgeSegments[0];
                    numberedGroup.SectionEdge = Tuple.Create(singleSegment.StartPoint, singleSegment.EndPoint);
                }
                else
                    numberedGroup.SectionEdge = ComputeSectionEdgeForPolylineMiddle(numberedGroup.PolylinePoints);

                result[numberedGroup.GroupNumber] = numberedGroup;
            }

            return result;
        }

        private static Tuple<TSG.Point, TSG.Point> ComputeSectionEdgeForPolylineMiddle(List<TSG.Point> polylinePoints) {
            if (polylinePoints == null || polylinePoints.Count < 2)
                return Tuple.Create(new TSG.Point(0, 0, 0), new TSG.Point(1, 0, 0));

            var totalLength = 0.0;

            for (var i = 0; i < polylinePoints.Count - 1; i++)
                totalLength += ComputeDistance2D(polylinePoints[i], polylinePoints[i + 1]);

            if (totalLength < 1e-9)
                return Tuple.Create(
                    new TSG.Point(polylinePoints[0].X, polylinePoints[0].Y, 0),
                    new TSG.Point(polylinePoints[0].X + 1.0, polylinePoints[0].Y, 0)
                );

            var halfLength = totalLength * 0.5;
            var walkedLength = 0.0;

            for (var i = 0; i < polylinePoints.Count - 1; i++) {
                var segmentStartPoint = polylinePoints[i];
                var segmentEndPoint = polylinePoints[i + 1];

                var segmentLength = ComputeDistance2D(segmentStartPoint, segmentEndPoint);
                if (segmentLength < 1e-9) continue;

                if (walkedLength + segmentLength >= halfLength) {
                    var remaining = halfLength - walkedLength;
                    var t = remaining / segmentLength;

                    var middlePoint = new TSG.Point(
                        segmentStartPoint.X + (segmentEndPoint.X - segmentStartPoint.X) * t,
                        segmentStartPoint.Y + (segmentEndPoint.Y - segmentStartPoint.Y) * t,
                        0
                    );

                    var directionX = segmentEndPoint.X - segmentStartPoint.X;
                    var directionY = segmentEndPoint.Y - segmentStartPoint.Y;

                    var directionLength = Math.Sqrt(directionX * directionX + directionY * directionY);
                    if (directionLength < 1e-9)
                        return Tuple.Create(
                            new TSG.Point(middlePoint.X - 1.0, middlePoint.Y, 0),
                            new TSG.Point(middlePoint.X + 1.0, middlePoint.Y, 0)
                        );

                    var unitDirectionX = directionX / directionLength;
                    var unitDirectionY = directionY / directionLength;

                    const double tangentSegmentHalfLengthMillimeters = 50.0;

                    var tangentStartPoint = new TSG.Point(
                        middlePoint.X - unitDirectionX * tangentSegmentHalfLengthMillimeters,
                        middlePoint.Y - unitDirectionY * tangentSegmentHalfLengthMillimeters,
                        0
                    );

                    var tangentEndPoint = new TSG.Point(
                        middlePoint.X + unitDirectionX * tangentSegmentHalfLengthMillimeters,
                        middlePoint.Y + unitDirectionY * tangentSegmentHalfLengthMillimeters,
                        0
                    );

                    return Tuple.Create(tangentStartPoint, tangentEndPoint);
                }

                walkedLength += segmentLength;
            }

            var lastStartPoint = polylinePoints[polylinePoints.Count - 2];
            var lastEndPoint = polylinePoints[polylinePoints.Count - 1];

            return Tuple.Create(
                new TSG.Point(lastStartPoint.X, lastStartPoint.Y, 0),
                new TSG.Point(lastEndPoint.X, lastEndPoint.Y, 0)
            );
        }

        private static List<EdgeSegment> BuildOrderedConnectedEdges(List<EdgeSegment> orderedEdges,
            double joinToleranceMillimeters) {
            if (orderedEdges == null || orderedEdges.Count == 0) return new List<EdgeSegment>();

            var connectedEdges = new List<EdgeSegment>();
            connectedEdges.Add(new EdgeSegment {
                EdgeNumber = orderedEdges[0].EdgeNumber,
                StartPoint = orderedEdges[0].StartPoint,
                EndPoint = orderedEdges[0].EndPoint
            });

            for (var edgeIndex = 1; edgeIndex < orderedEdges.Count; edgeIndex++) {
                var lastConnectedEdge = connectedEdges[connectedEdges.Count - 1];
                var lastPointInChain = lastConnectedEdge.EndPoint;

                var currentEdge = orderedEdges[edgeIndex];
                var currentStartPoint = currentEdge.StartPoint;
                var currentEndPoint = currentEdge.EndPoint;

                var distanceToStartPoint = ComputeDistance2D(lastPointInChain, currentStartPoint);
                var distanceToEndPoint = ComputeDistance2D(lastPointInChain, currentEndPoint);

                if (distanceToStartPoint <= joinToleranceMillimeters) {
                    connectedEdges.Add(new EdgeSegment {
                        EdgeNumber = currentEdge.EdgeNumber,
                        StartPoint = currentStartPoint,
                        EndPoint = currentEndPoint
                    });
                    continue;
                }

                if (distanceToEndPoint <= joinToleranceMillimeters) {
                    connectedEdges.Add(new EdgeSegment {
                        EdgeNumber = currentEdge.EdgeNumber,
                        StartPoint = currentEndPoint,
                        EndPoint = currentStartPoint
                    });
                    continue;
                }

                // Fallback: wybieramy orientację z bliższym końcem (żeby zachować ciągłość wizualnie)
                if (distanceToStartPoint <= distanceToEndPoint)
                    connectedEdges.Add(new EdgeSegment {
                        EdgeNumber = currentEdge.EdgeNumber,
                        StartPoint = currentStartPoint,
                        EndPoint = currentEndPoint
                    });
                else
                    connectedEdges.Add(new EdgeSegment {
                        EdgeNumber = currentEdge.EdgeNumber,
                        StartPoint = currentEndPoint,
                        EndPoint = currentStartPoint
                    });
            }

            return connectedEdges;
        }

        private static List<PolylineGroup> SplitEdgesIntoPolylineGroupsByAngle(List<EdgeSegment> connectedEdges,
            double nearStraightAngleDegrees) {
            var polylineGroups = new List<PolylineGroup>();
            if (connectedEdges == null || connectedEdges.Count == 0) return polylineGroups;

            var currentGroup = new PolylineGroup();
            currentGroup.EdgeSegments.Add(connectedEdges[0]);
            currentGroup.PolylinePoints.Add(connectedEdges[0].StartPoint);
            currentGroup.PolylinePoints.Add(connectedEdges[0].EndPoint);

            for (var edgeIndex = 1; edgeIndex < connectedEdges.Count; edgeIndex++) {
                var previousEdge = connectedEdges[edgeIndex - 1];
                var currentEdge = connectedEdges[edgeIndex];

                var vertexPoint = previousEdge.EndPoint;
                var previousPoint = previousEdge.StartPoint;
                var nextPoint = currentEdge.EndPoint;

                var angleInDegrees = ComputeAngleInDegrees(previousPoint, vertexPoint, nextPoint, 1e-6);

                if (angleInDegrees >= nearStraightAngleDegrees) {
                    currentGroup.EdgeSegments.Add(currentEdge);
                    currentGroup.PolylinePoints.Add(currentEdge.EndPoint);
                }
                else {
                    polylineGroups.Add(currentGroup);

                    currentGroup = new PolylineGroup();
                    currentGroup.EdgeSegments.Add(currentEdge);
                    currentGroup.PolylinePoints.Add(currentEdge.StartPoint);
                    currentGroup.PolylinePoints.Add(currentEdge.EndPoint);
                }
            }

            polylineGroups.Add(currentGroup);
            return polylineGroups;
        }

        private static TSG.Point ComputeTextInsertionPointForSegment(TSG.Point startPoint, TSG.Point endPoint,
            double offsetMillimeters) {
            var midpoint = new TSG.Point(
                (startPoint.X + endPoint.X) * 0.5,
                (startPoint.Y + endPoint.Y) * 0.5,
                0
            );

            var directionX = endPoint.X - startPoint.X;
            var directionY = endPoint.Y - startPoint.Y;

            var directionLength = Math.Sqrt(directionX * directionX + directionY * directionY);
            if (directionLength < 1e-9) return new TSG.Point(midpoint.X, midpoint.Y + offsetMillimeters, 0);

            var normalX = -directionY / directionLength;
            var normalY = directionX / directionLength;

            return new TSG.Point(
                midpoint.X + normalX * offsetMillimeters,
                midpoint.Y + normalY * offsetMillimeters,
                0
            );
        }

        private static TSG.Point ComputeTextInsertionPointForPolyline(List<TSG.Point> polylinePoints,
            double offsetMillimeters) {
            if (polylinePoints == null || polylinePoints.Count < 2) return new TSG.Point(0, 0, 0);

            var sumX = 0.0;
            var sumY = 0.0;

            foreach (var polylinePoint in polylinePoints) {
                sumX += polylinePoint.X;
                sumY += polylinePoint.Y;
            }

            var averageX = sumX / polylinePoints.Count;
            var averageY = sumY / polylinePoints.Count;

            var firstPoint = polylinePoints[0];
            var lastPoint = polylinePoints[polylinePoints.Count - 1];

            var directionX = lastPoint.X - firstPoint.X;
            var directionY = lastPoint.Y - firstPoint.Y;

            var directionLength = Math.Sqrt(directionX * directionX + directionY * directionY);
            if (directionLength < 1e-9) return new TSG.Point(averageX, averageY + offsetMillimeters, 0);

            var normalX = -directionY / directionLength;
            var normalY = directionX / directionLength;

            return new TSG.Point(
                averageX + normalX * offsetMillimeters,
                averageY + normalY * offsetMillimeters,
                0
            );
        }

        private static double ComputeDistance2D(TSG.Point firstPoint, TSG.Point secondPoint) {
            var differenceX = firstPoint.X - secondPoint.X;
            var differenceY = firstPoint.Y - secondPoint.Y;
            return Math.Sqrt(differenceX * differenceX + differenceY * differenceY);
        }

        private static double ComputeAngleInDegrees(TSG.Point firstPoint, TSG.Point vertexPoint, TSG.Point secondPoint,
            double minimumSegmentLength) {
            var firstVectorX = firstPoint.X - vertexPoint.X;
            var firstVectorY = firstPoint.Y - vertexPoint.Y;

            var secondVectorX = secondPoint.X - vertexPoint.X;
            var secondVectorY = secondPoint.Y - vertexPoint.Y;

            var firstVectorLength = Math.Sqrt(firstVectorX * firstVectorX + firstVectorY * firstVectorY);
            var secondVectorLength = Math.Sqrt(secondVectorX * secondVectorX + secondVectorY * secondVectorY);

            if (firstVectorLength < minimumSegmentLength || secondVectorLength < minimumSegmentLength) return 180.0;

            var dotProduct = firstVectorX * secondVectorX + firstVectorY * secondVectorY;
            var cosineValue = dotProduct / (firstVectorLength * secondVectorLength);

            if (cosineValue > 1.0) cosineValue = 1.0;
            if (cosineValue < -1.0) cosineValue = -1.0;

            var angleInRadians = Math.Acos(cosineValue);
            return angleInRadians * 180.0 / Math.PI;
        }

        /// Filtruje słownik krawędzi, pozostawiając tylko te, których numery są w zbiorze "requested". Jeśli "requested" jest pusty, zwraca wszystkie krawędzie. Jeśli po filtracji nie zostanie żadna krawędź, pokazuje komunikat. Dodatkowo informuje o brakujących numerach.
        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> FilterEdgesOrShowMessage(
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edges,
            HashSet<int> requested
        ) {
            if (edges == null || edges.Count == 0) return null;
            if (requested == null || requested.Count == 0) return edges;

            var filtered = edges
                .Where(kvp => requested.Contains(kvp.Key))
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            if (filtered.Count == 0) {
                MessageBox.Show("Nie znaleziono krawędzi o podanych numerach.");
                return null;
            }

            var missing = requested.Where(n => !edges.ContainsKey(n)).OrderBy(n => n).ToList();
            if (missing.Count > 0)
                MessageBox.Show("Brak krawędzi o numerach: " + string.Join(", ", missing));

            return filtered;
        }

        /// Zapewnia, że linia będzie skierowana w lewo lub w górę (w zależności od orientacji), co jest ważne dla poprawnego działania funkcji tworzenia przekroju.
        private static void ForceSectionLookLeftOrUp(ref TSG.Point startPoint, ref TSG.Point endPoint) {
            var dx = endPoint.X - startPoint.X;
            var dy = endPoint.Y - startPoint.Y;

            var len = Math.Sqrt(dx * dx + dy * dy);
            if (len < 1e-9) return;

            var nx = -dy / len;
            var ny = dx / len;

            if (Math.Abs(ny) >= Math.Abs(nx)) {
                if (ny < 0) (startPoint, endPoint) = (endPoint, startPoint);
            }
            else {
                if (nx > 0) (startPoint, endPoint) = (endPoint, startPoint);
            }
        }

        /// Transformuje listę krawędzi (punktów A i B) z jednej CS do drugiej, zachowując tylko X i Y (ustawiając Z=0).
        private static List<(TSG.Point A, TSG.Point B)> TransformEdgesBetweenCoordinateSystems(
            TSG.CoordinateSystem fromCs,
            TSG.CoordinateSystem toCs,
            List<(TSG.Point A, TSG.Point B)> edges
        ) {
            if (edges == null || edges.Count == 0)
                return new List<(TSG.Point A, TSG.Point B)>();

            var fromToGlobal = TSG.MatrixFactory.FromCoordinateSystem(fromCs);
            var globalToTo = TSG.MatrixFactory.ToCoordinateSystem(toCs);

            var result = new List<(TSG.Point A, TSG.Point B)>(edges.Count);

            for (var i = 0; i < edges.Count; i++) {
                var aGlobal = fromToGlobal.Transform(edges[i].A);
                var bGlobal = fromToGlobal.Transform(edges[i].B);

                var aTo = globalToTo.Transform(aGlobal);
                var bTo = globalToTo.Transform(bGlobal);

                result.Add((
                    new TSG.Point(aTo.X, aTo.Y, 0),
                    new TSG.Point(bTo.X, bTo.Y, 0)
                ));
            }

            return result;
        }

        /// Oblicza środek krawędzi w 2D (ignorując Z) jako środek prostokąta otaczającego wszystkie punkty A i B.
        private static TSG.Point ComputeEdgesCenter2D(List<(TSG.Point A, TSG.Point B)> edges) {
            var pts = new List<TSG.Point>(edges.Count * 2);
            for (var i = 0; i < edges.Count; i++) {
                pts.Add(edges[i].A);
                pts.Add(edges[i].B);
            }

            var minX = pts.Min(p => p.X);
            var maxX = pts.Max(p => p.X);
            var minY = pts.Min(p => p.Y);
            var maxY = pts.Max(p => p.Y);

            return new TSG.Point((minX + maxX) * 0.5, (minY + maxY) * 0.5, 0);
        }

        /// Usuwa punkty, które są blisko siebie (w odległości epsilon)
        private static List<TSG.Point> RemoveNearDuplicates(List<TSG.Point> points, double epsilon) {
            if (points == null || points.Count == 0) return new List<TSG.Point>();

            var result = new List<TSG.Point>();
            foreach (var p in from p in points
                     let p1 = p
                     let exists = result.Any(r => Math.Abs(r.X - p1.X) <= epsilon && Math.Abs(r.Y - p1.Y) <= epsilon)
                     where !exists
                     select p) result.Add(p);

            return result;
        }

        /// Sortuje punkty według kąta względem środka (średniej) wszystkich punktów. Zakłada, że punkty są w jednej płaszczyźnie.
        private static List<TSG.Point> SortByAngle(List<TSG.Point> points) {
            if (points == null || points.Count < 3) return points ?? new List<TSG.Point>();

            var cx = points.Average(p => p.X);
            var cy = points.Average(p => p.Y);

            return points
                .OrderBy(p => Math.Atan2(p.Y - cy, p.X - cx))
                .ToList();
        }

        #endregion

        #endregion
    }
}