using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region View And Model Helpers

        private static TPart FindFirstPart<TPart>(List<TSM.Part> parts) where TPart : TSM.Part {
            if (parts == null || parts.Count == 0) return null;

            foreach (var part in parts)
                if (part is TPart typedPart)
                    return typedPart;

            return null;
        }

        private static bool HasExactlyOnePart(TSD.View selectedView) {
            var modelParts = GetModelPartsFromDrawingView(selectedView);
            return modelParts != null && modelParts.Count == 1;
        }

        #endregion

        #region Constants

        private const double ZStepMillimeters = 0.5;
        private const int MaximumEmptyStepsPerDirection = 80;
        private const double DuplicateToleranceMillimeters = 0.5;
        private const double MinimumLengthMillimeters = 100.0;
        private const double MinimumContourPlateLengthMillimeters = 100.0;
        private const double NearStraightAngleDegrees = 170.0;
        private const double JoinToleranceMillimeters = 0.5;
        private const double NumberOffsetMillimeters = 20.0;

        private const int TargetBinCount = 60;
        private const int SmoothingWindowRadius = 2;
        private const int EdgeLockBinCount = 3;
        private const int MaximumInlineEdgeCount = 10;

        #endregion

        #region Main Edge Detection Flow

        private static string DrawEdgesWithNumbers() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return null;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return null;

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return null;

            if (!HasExactlyOnePart(selectedView)) {
                MessageBox.Show("Widok musi zawierać dokładnie jeden element typu Part.");
                return null;
            }

            var detectionResult = DetectEdgesFromSelectedView(selectedView);

            if (!detectionResult.HasEdges) {
                MessageBox.Show(detectionResult.ErrorMessage);
                return null;
            }

            var numberedGroups = BuildNumberedEdgeGroups(
                detectionResult.EdgesByNumber,
                JoinToleranceMillimeters,
                NearStraightAngleDegrees
            );

            DrawEdges(selectedView, detectionResult.EdgesByNumber);
            drawing.CommitChanges();

            return FormatEdgeNumbersForTextBox(numberedGroups.Keys.OrderBy(number => number).ToList());
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
            if (contourEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "ContourPlate",
                    EdgesByNumber = contourEdges
                };

            var loftedEdges = GetLoftedPlateEdges(selectedView, modelParts);
            if (loftedEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "LoftedPlate",
                    EdgesByNumber = loftedEdges
                };

            var beamEdges = GetBeamEdges(selectedView, modelParts);
            if (beamEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "Beam",
                    EdgesByNumber = beamEdges
                };

            var polyBeamEdges = GetPolyBeamEdges(selectedView, modelParts);
            if (polyBeamEdges.Count > 0)
                return new EdgeDetectionResult {
                    SourceName = "PolyBeam",
                    EdgesByNumber = polyBeamEdges
                };

            var foundTypes = string.Join(
                ", ",
                modelParts
                    .Select(part => part.GetType().Name)
                    .Distinct()
                    .OrderBy(typeName => typeName)
            );

            return new EdgeDetectionResult {
                ErrorMessage =
                    "Nie udało się wyznaczyć krawędzi dla żadnego obsługiwanego typu elementu.\n" +
                    "Znalezione typy na widoku: " + foundTypes
            };
        }

        #endregion

        #region Edge Detection By Part Type

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetContourPlateEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;
            if (modelParts == null || modelParts.Count == 0) return edgesByNumber;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var contourPlate = FindFirstPart<TSM.ContourPlate>(modelParts);
                if (contourPlate == null) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(contourPlate.GetCoordinateSystem())
                );

                var edgesInPlatePlane = GetContourPlateEdgesInPlane(contourPlate);
                if (edgesInPlatePlane == null || edgesInPlatePlane.Count == 0) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem)
                );

                var edgesInViewPlane = TransformEdgesBetweenCoordinateSystems(
                    contourPlate.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    edgesInPlatePlane
                );

                var rotatedEdges = RotateEdgesAroundCenter2D(edgesInViewPlane, 90.0);
                var shiftedEdges = TranslateEdgesDownByHalfHeight(rotatedEdges);

                var number = 0;

                for (var index = 0; index < shiftedEdges.Count; index++) {
                    var startPoint = shiftedEdges[index].A;
                    var endPoint = shiftedEdges[index].B;

                    var dx = endPoint.X - startPoint.X;
                    var dy = endPoint.Y - startPoint.Y;
                    var dz = endPoint.Z - startPoint.Z;

                    var length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (length <= MinimumContourPlateLengthMillimeters) continue;

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

            for (var index = 0; index < points.Count; index++) {
                var startPoint = points[index];
                var endPoint = points[(index + 1) % points.Count];
                edges.Add((startPoint, endPoint));
            }

            return edges;
        }

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetLoftedPlateEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            return GetSolidBasedPartEdges<TSM.LoftedPlate>(
                selectedView,
                modelParts,
                BuildLoftedPlateEnvelopeEdgesInView
            );
        }

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetBeamEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            return GetSolidBasedPartEdges<TSM.Beam>(
                selectedView,
                modelParts,
                BuildBeamEnvelopeEdgesInView
            );
        }

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetPolyBeamEdges(
            TSD.View selectedView,
            List<TSM.Part> modelParts
        ) {
            return GetSolidBasedPartEdges<TSM.PolyBeam>(
                selectedView,
                modelParts,
                BuildBeamEnvelopeEdgesInView
            );
        }

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetSolidBasedPartEdges<TPart>(
            TSD.View selectedView,
            List<TSM.Part> modelParts,
            Func<List<TSG.Point>, List<(TSG.Point A, TSG.Point B)>> buildEnvelopeEdges
        ) where TPart : TSM.Part {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;
            if (modelParts == null || modelParts.Count == 0) return edgesByNumber;

            var part = FindFirstPart<TPart>(modelParts);
            if (part == null) return edgesByNumber;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(part.GetCoordinateSystem())
                );

                var solid = part.GetSolid();
                if (solid == null) return edgesByNumber;

                var sampledPointsInPartPlane = GetSamplePointsByAdaptiveZSweep(solid);
                if (sampledPointsInPartPlane.Count == 0) return edgesByNumber;

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem)
                );

                var sampledPointsInViewPlane = TransformPointsBetweenCoordinateSystems(
                    part.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    sampledPointsInPartPlane
                );

                if (sampledPointsInViewPlane.Count == 0) return edgesByNumber;

                var edgesInViewPlane = buildEnvelopeEdges(sampledPointsInViewPlane);
                if (edgesInViewPlane == null || edgesInViewPlane.Count == 0) return edgesByNumber;

                var number = 0;

                for (var index = 0; index < edgesInViewPlane.Count; index++) {
                    var startPoint = edgesInViewPlane[index].A;
                    var endPoint = edgesInViewPlane[index].B;

                    var dx = endPoint.X - startPoint.X;
                    var dy = endPoint.Y - startPoint.Y;
                    var dz = endPoint.Z - startPoint.Z;

                    var length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (length <= MinimumLengthMillimeters) continue;

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

        #endregion

        #region Sampling

        private static List<TSG.Point> GetSamplePointsByAdaptiveZSweep(TSM.Solid solid) {
            var points = new List<TSG.Point>();
            if (solid == null) return points;

            var centerZ = (solid.MinimumPoint.Z + solid.MaximumPoint.Z) * 0.5;

            var upwardEmptySteps = 0;
            var downwardEmptySteps = 0;

            var centerPoints = GetIntersectionPointsAtLocalZ(solid, centerZ);
            if (centerPoints.Count > 0)
                points.AddRange(centerPoints);
            else {
                upwardEmptySteps++;
                downwardEmptySteps++;
            }

            for (var offset = ZStepMillimeters;
                 upwardEmptySteps < MaximumEmptyStepsPerDirection ||
                 downwardEmptySteps < MaximumEmptyStepsPerDirection;
                 offset += ZStepMillimeters) {
                if (upwardEmptySteps < MaximumEmptyStepsPerDirection) {
                    var upperPoints = GetIntersectionPointsAtLocalZ(solid, centerZ + offset);

                    if (upperPoints.Count > 0) {
                        points.AddRange(upperPoints);
                        upwardEmptySteps = 0;
                    }
                    else
                        upwardEmptySteps++;
                }

                if (downwardEmptySteps < MaximumEmptyStepsPerDirection) {
                    var lowerPoints = GetIntersectionPointsAtLocalZ(solid, centerZ - offset);

                    if (lowerPoints.Count > 0) {
                        points.AddRange(lowerPoints);
                        downwardEmptySteps = 0;
                    }
                    else
                        downwardEmptySteps++;
                }
            }

            return RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
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
                if (!(enumerator.Current is TSG.Point point)) continue;
                points.Add(new TSG.Point(point.X, point.Y, point.Z));
            }

            return points;
        }

        #endregion

        #region Envelope Building

        private static List<(TSG.Point A, TSG.Point B)> BuildLoftedPlateEnvelopeEdgesInView(List<TSG.Point> points) {
            var edges = new List<(TSG.Point A, TSG.Point B)>();
            if (points == null || points.Count < 3) return edges;

            var uniquePoints = RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
            if (uniquePoints.Count < 3) return edges;

            var minimumX = uniquePoints.Min(point => point.X);
            var maximumX = uniquePoints.Max(point => point.X);

            if (maximumX - minimumX < 1e-6) return edges;

            var binWidth = Math.Max(2.0, (maximumX - minimumX) / TargetBinCount);
            var bins = new SortedDictionary<int, List<TSG.Point>>();

            foreach (var point in uniquePoints) {
                var binIndex = (int)Math.Floor((point.X - minimumX) / binWidth);

                if (!bins.ContainsKey(binIndex)) bins[binIndex] = new List<TSG.Point>();

                bins[binIndex].Add(point);
            }

            var orderedBins = bins
                .OrderBy(pair => pair.Key)
                .Select(pair => pair.Value.OrderBy(point => point.Y).ThenBy(point => point.X).ToList())
                .Where(bucket => bucket.Count > 0)
                .ToList();

            if (orderedBins.Count < 2) return edges;

            var upperChain = new List<TSG.Point>();
            var lowerChain = new List<TSG.Point>();

            foreach (var bucket in orderedBins) {
                var lowerPointIndex = (int)Math.Floor((bucket.Count - 1) * 0.15);
                var upperPointIndex = (int)Math.Ceiling((bucket.Count - 1) * 0.85);

                lowerChain.Add(new TSG.Point(bucket[lowerPointIndex].X, bucket[lowerPointIndex].Y, 0));
                upperChain.Add(new TSG.Point(bucket[upperPointIndex].X, bucket[upperPointIndex].Y, 0));
            }

            upperChain = SmoothPolylineByMedianY(upperChain, SmoothingWindowRadius);
            lowerChain = SmoothPolylineByMedianY(lowerChain, SmoothingWindowRadius);

            if (upperChain.Count < 2 || lowerChain.Count < 2) return edges;

            var leftLockedX = uniquePoints
                .OrderBy(point => point.X)
                .Take(Math.Min(20, uniquePoints.Count))
                .Average(point => point.X);

            var rightLockedX = uniquePoints
                .OrderByDescending(point => point.X)
                .Take(Math.Min(20, uniquePoints.Count))
                .Average(point => point.X);

            var leftLockCount = Math.Min(EdgeLockBinCount, Math.Min(upperChain.Count, lowerChain.Count));
            var rightLockCount = Math.Min(EdgeLockBinCount, Math.Min(upperChain.Count, lowerChain.Count));

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

            upperChain = RemoveNearDuplicates(upperChain, DuplicateToleranceMillimeters);
            lowerChain = RemoveNearDuplicates(lowerChain, DuplicateToleranceMillimeters);

            upperChain = RemoveVeryShortSegmentsFromOpenPolyline(upperChain, MinimumLengthMillimeters);
            lowerChain = RemoveVeryShortSegmentsFromOpenPolyline(lowerChain, MinimumLengthMillimeters);

            if (upperChain.Count < 2 || lowerChain.Count < 2) return edges;

            var boundaryPoints = upperChain.Select(point => new TSG.Point(point.X, point.Y, 0)).ToList();

            for (var index = lowerChain.Count - 1; index >= 0; index--)
                boundaryPoints.Add(new TSG.Point(lowerChain[index].X, lowerChain[index].Y, 0));

            boundaryPoints = RemoveNearDuplicates(boundaryPoints, DuplicateToleranceMillimeters);
            boundaryPoints = RemoveVeryShortSegmentsFromClosedPolyline(boundaryPoints, MinimumLengthMillimeters);
            boundaryPoints = RemoveCollinearPointsFromClosedPolyline(boundaryPoints, DuplicateToleranceMillimeters);

            if (boundaryPoints.Count < 3) return edges;

            for (var index = 0; index < boundaryPoints.Count; index++) {
                var startPoint = boundaryPoints[index];
                var endPoint = boundaryPoints[(index + 1) % boundaryPoints.Count];

                if (ComputeDistance2D(startPoint, endPoint) <= DuplicateToleranceMillimeters) continue;

                edges.Add((
                    new TSG.Point(startPoint.X, startPoint.Y, 0),
                    new TSG.Point(endPoint.X, endPoint.Y, 0)
                ));
            }

            return edges;
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

                if (ComputeDistance2D(startPoint, endPoint) <= DuplicateToleranceMillimeters) continue;

                edges.Add((
                    new TSG.Point(startPoint.X, startPoint.Y, 0),
                    new TSG.Point(endPoint.X, endPoint.Y, 0)
                ));
            }

            return edges;
        }

        #endregion

        #region Drawing And Numbering

        private static void DrawEdges(TSD.View view, Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber) {
            if (view == null) return;
            if (edgesByNumber == null || edgesByNumber.Count == 0) return;

            var lineAttributes = new TSD.Line.LineAttributes {
                Line = new TSD.LineTypeAttributes(TSD.LineTypes.SolidLine, TSD.DrawingColors.Red)
            };

            var numberedGroups = BuildNumberedEdgeGroups(
                edgesByNumber,
                JoinToleranceMillimeters,
                NearStraightAngleDegrees
            );

            foreach (var pair in numberedGroups.OrderBy(p => p.Key)) {
                var group = pair.Value;
                if (group == null) continue;

                if (!group.IsPolyline) {
                    var singleSegment = group.EdgeSegments[0];

                    new TSD.Line(view, singleSegment.StartPoint, singleSegment.EndPoint, lineAttributes).Insert();

                    var numberPoint = ComputeTextInsertionPointForSegment(
                        singleSegment.StartPoint,
                        singleSegment.EndPoint,
                        NumberOffsetMillimeters
                    );

                    new TSD.Text(view, numberPoint, group.GroupNumber.ToString()).Insert();
                    continue;
                }

                var polylinePointList = new TSD.PointList();
                foreach (var polylinePoint in group.PolylinePoints)
                    polylinePointList.Add(new TSG.Point(polylinePoint.X, polylinePoint.Y, 0));

                var polyline = new TSD.Polyline(view, polylinePointList);
                polyline.Insert();

                var polylineNumberPoint = ComputeTextInsertionPointForPolyline(
                    group.PolylinePoints,
                    NumberOffsetMillimeters
                );

                new TSD.Text(view, polylineNumberPoint, group.GroupNumber.ToString()).Insert();
            }
        }

        private static HashSet<int> ParseEdgeNumbers(string input) {
            var result = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(input)) return result;

            var tokens = input
                .Replace(";", ",")
                .Replace(" ", ",")
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim())
                .Where(token => token.Length > 0);

            foreach (var token in tokens)
                if (token.Contains("-")) {
                    var parts = token.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(part => part.Trim())
                        .ToArray();

                    if (parts.Length != 2) continue;
                    if (!int.TryParse(parts[0], out var start)) continue;
                    if (!int.TryParse(parts[1], out var end)) continue;

                    if (start > end) (start, end) = (end, start);

                    for (var index = start; index <= end; index++) result.Add(index);
                }
                else {
                    if (int.TryParse(token, out var number)) result.Add(number);
                }

            return result;
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

            foreach (var polylineGroup in polylineGroups.Where(polylineGroup =>
                         polylineGroup?.EdgeSegments != null && polylineGroup.EdgeSegments.Count != 0)) {
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

            for (var index = 0; index < polylinePoints.Count - 1; index++)
                totalLength += ComputeDistance2D(polylinePoints[index], polylinePoints[index + 1]);

            if (totalLength < 1e-9)
                return Tuple.Create(
                    new TSG.Point(polylinePoints[0].X, polylinePoints[0].Y, 0),
                    new TSG.Point(polylinePoints[0].X + 1.0, polylinePoints[0].Y, 0)
                );

            var halfLength = totalLength * 0.5;
            var walkedLength = 0.0;

            for (var index = 0; index < polylinePoints.Count - 1; index++) {
                var segmentStartPoint = polylinePoints[index];
                var segmentEndPoint = polylinePoints[index + 1];

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

        private static List<EdgeSegment> BuildOrderedConnectedEdges(
            List<EdgeSegment> orderedEdges,
            double joinToleranceMillimeters
        ) {
            if (orderedEdges == null || orderedEdges.Count == 0) return new List<EdgeSegment>();

            var connectedEdges = new List<EdgeSegment> {
                new EdgeSegment {
                    EdgeNumber = orderedEdges[0].EdgeNumber,
                    StartPoint = orderedEdges[0].StartPoint,
                    EndPoint = orderedEdges[0].EndPoint
                }
            };

            for (var index = 1; index < orderedEdges.Count; index++) {
                var lastConnectedEdge = connectedEdges[connectedEdges.Count - 1];
                var lastPointInChain = lastConnectedEdge.EndPoint;

                var currentEdge = orderedEdges[index];
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

        private static List<PolylineGroup> SplitEdgesIntoPolylineGroupsByAngle(
            List<EdgeSegment> connectedEdges,
            double nearStraightAngleDegrees
        ) {
            var polylineGroups = new List<PolylineGroup>();
            if (connectedEdges == null || connectedEdges.Count == 0) return polylineGroups;

            var currentGroup = new PolylineGroup();
            currentGroup.EdgeSegments.Add(connectedEdges[0]);
            currentGroup.PolylinePoints.Add(connectedEdges[0].StartPoint);
            currentGroup.PolylinePoints.Add(connectedEdges[0].EndPoint);

            for (var index = 1; index < connectedEdges.Count; index++) {
                var previousEdge = connectedEdges[index - 1];
                var currentEdge = connectedEdges[index];

                var vertexPoint = previousEdge.EndPoint;
                var previousPoint = previousEdge.StartPoint;
                var nextPoint = currentEdge.EndPoint;

                var angleInDegrees = ComputeAngleInDegrees(previousPoint, vertexPoint, nextPoint, 1e-6);

                if (angleInDegrees >= nearStraightAngleDegrees)
                    currentGroup.EdgeSegments.Add(currentEdge);
                else {
                    polylineGroups.Add(currentGroup);

                    currentGroup = new PolylineGroup();
                    currentGroup.EdgeSegments.Add(currentEdge);
                    currentGroup.PolylinePoints.Add(currentEdge.StartPoint);
                }

                currentGroup.PolylinePoints.Add(currentEdge.EndPoint);
            }

            polylineGroups.Add(currentGroup);
            return polylineGroups;
        }

        private static string FormatEdgeNumbersForTextBox(List<int> numbers) {
            if (numbers == null || numbers.Count == 0) return string.Empty;

            var orderedNumbers = numbers
                .Distinct()
                .OrderBy(number => number)
                .ToList();

            if (orderedNumbers.Count <= MaximumInlineEdgeCount)
                return string.Join(",", orderedNumbers);

            return orderedNumbers.First() + "-" + orderedNumbers.Last();
        }

        private static TSG.Point ComputeTextInsertionPointForSegment(
            TSG.Point startPoint,
            TSG.Point endPoint,
            double offsetMillimeters
        ) {
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

        private static TSG.Point ComputeTextInsertionPointForPolyline(
            List<TSG.Point> polylinePoints,
            double offsetMillimeters
        ) {
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

        #endregion

        #region Coordinate Systems And Transformations

        private static List<(TSG.Point A, TSG.Point B)> TransformEdgesBetweenCoordinateSystems(
            TSG.CoordinateSystem fromCoordinateSystem,
            TSG.CoordinateSystem toCoordinateSystem,
            List<(TSG.Point A, TSG.Point B)> edges
        ) {
            if (edges == null || edges.Count == 0) return new List<(TSG.Point A, TSG.Point B)>();

            var fromToGlobal = TSG.MatrixFactory.FromCoordinateSystem(fromCoordinateSystem);
            var globalToTarget = TSG.MatrixFactory.ToCoordinateSystem(toCoordinateSystem);

            var result = new List<(TSG.Point A, TSG.Point B)>(edges.Count);

            for (var index = 0; index < edges.Count; index++) {
                var aGlobal = fromToGlobal.Transform(edges[index].A);
                var bGlobal = fromToGlobal.Transform(edges[index].B);

                var aTarget = globalToTarget.Transform(aGlobal);
                var bTarget = globalToTarget.Transform(bGlobal);

                result.Add((
                    new TSG.Point(aTarget.X, aTarget.Y, 0),
                    new TSG.Point(bTarget.X, bTarget.Y, 0)
                ));
            }

            return result;
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

            result.AddRange(from point in points
                select fromToGlobal.Transform(point)
                into globalPoint
                select globalToTarget.Transform(globalPoint)
                into targetPoint
                select new TSG.Point(targetPoint.X, targetPoint.Y, 0));

            return result;
        }

        private static TSG.Point ComputeEdgesCenter2D(List<(TSG.Point A, TSG.Point B)> edges) {
            var points = new List<TSG.Point>(edges.Count * 2);
            for (var index = 0; index < edges.Count; index++) {
                points.Add(edges[index].A);
                points.Add(edges[index].B);
            }

            var minX = points.Min(point => point.X);
            var maxX = points.Max(point => point.X);
            var minY = points.Min(point => point.Y);
            var maxY = points.Max(point => point.Y);

            return new TSG.Point((minX + maxX) * 0.5, (minY + maxY) * 0.5, 0);
        }

        private static List<(TSG.Point A, TSG.Point B)> RotateEdgesAroundCenter2D(
            List<(TSG.Point A, TSG.Point B)> edges,
            double angleDegrees
        ) {
            if (edges == null || edges.Count == 0) return new List<(TSG.Point A, TSG.Point B)>();

            var beforePoints = new List<TSG.Point>(edges.Count * 2);
            for (var index = 0; index < edges.Count; index++) {
                beforePoints.Add(edges[index].A);
                beforePoints.Add(edges[index].B);
            }

            var beforeMinX = beforePoints.Min(point => point.X);
            var beforeMinY = beforePoints.Min(point => point.Y);

            var center = ComputeEdgesCenter2D(edges);
            var radians = angleDegrees * Math.PI / 180.0;
            var cos = Math.Cos(radians);
            var sin = Math.Sin(radians);

            var rotatedEdges = new List<(TSG.Point A, TSG.Point B)>(edges.Count);
            for (var index = 0; index < edges.Count; index++)
                rotatedEdges.Add((
                    RotatePointAroundCenter2D(edges[index].A, center, cos, sin),
                    RotatePointAroundCenter2D(edges[index].B, center, cos, sin)
                ));

            var afterPoints = new List<TSG.Point>(rotatedEdges.Count * 2);
            for (var index = 0; index < rotatedEdges.Count; index++) {
                afterPoints.Add(rotatedEdges[index].A);
                afterPoints.Add(rotatedEdges[index].B);
            }

            var afterMinX = afterPoints.Min(point => point.X);
            var afterMinY = afterPoints.Min(point => point.Y);

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

            for (var index = 0; index < rotatedEdges.Count; index++) {
                var a = rotatedEdges[index].A;
                var b = rotatedEdges[index].B;

                rotatedEdges[index] = (
                    new TSG.Point(a.X + deltaX, a.Y + deltaY, 0),
                    new TSG.Point(b.X + deltaX, b.Y + deltaY, 0)
                );
            }

            return rotatedEdges;
        }

        private static List<(TSG.Point A, TSG.Point B)> TranslateEdgesDownByHalfHeight(
            List<(TSG.Point A, TSG.Point B)> edges
        ) {
            if (edges == null || edges.Count == 0) return edges;

            var points = new List<TSG.Point>(edges.Count * 2);
            for (var index = 0; index < edges.Count; index++) {
                points.Add(edges[index].A);
                points.Add(edges[index].B);
            }

            var minY = points.Min(point => point.Y);
            var maxY = points.Max(point => point.Y);
            var halfHeight = (maxY - minY) * 0.5;

            var result = new List<(TSG.Point A, TSG.Point B)>(edges.Count);
            for (var index = 0; index < edges.Count; index++) {
                var a = edges[index].A;
                var b = edges[index].B;

                result.Add((
                    new TSG.Point(a.X, a.Y - halfHeight, 0),
                    new TSG.Point(b.X, b.Y - halfHeight, 0)
                ));
            }

            return result;
        }

        private static TSG.Point RotatePointAroundCenter2D(
            TSG.Point point,
            TSG.Point center,
            double cos,
            double sin
        ) {
            var dx = point.X - center.X;
            var dy = point.Y - center.Y;

            return new TSG.Point(
                center.X + dx * cos - dy * sin,
                center.Y + dx * sin + dy * cos,
                0
            );
        }

        #endregion

        #region Geometry Helpers

        private static double ComputeDistance2D(TSG.Point firstPoint, TSG.Point secondPoint) {
            var differenceX = firstPoint.X - secondPoint.X;
            var differenceY = firstPoint.Y - secondPoint.Y;
            return Math.Sqrt(differenceX * differenceX + differenceY * differenceY);
        }

        private static double ComputeAngleInDegrees(
            TSG.Point firstPoint,
            TSG.Point vertexPoint,
            TSG.Point secondPoint,
            double minimumSegmentLength
        ) {
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

            return Math.Acos(cosineValue) * 180.0 / Math.PI;
        }

        private static List<TSG.Point> RemoveNearDuplicates(List<TSG.Point> points, double epsilon) {
            if (points == null || points.Count == 0) return new List<TSG.Point>();

            var seen = new HashSet<(long, long)>();
            var result = new List<TSG.Point>();

            foreach (var point in points) {
                var key = (
                    (long)Math.Round(point.X / epsilon),
                    (long)Math.Round(point.Y / epsilon)
                );

                if (seen.Add(key))
                    result.Add(point);
            }

            return result;
        }

        private static List<TSG.Point> SortByAngle(List<TSG.Point> points) {
            if (points == null || points.Count < 3) return points ?? new List<TSG.Point>();

            var centerX = points.Average(point => point.X);
            var centerY = points.Average(point => point.Y);

            return points
                .OrderBy(point => Math.Atan2(point.Y - centerY, point.X - centerX))
                .ToList();
        }

        private static List<TSG.Point> BuildConvexHull2D(List<TSG.Point> points) {
            if (points == null || points.Count < 3) return points ?? new List<TSG.Point>();

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
                result.Add(new TSG.Point(points[index].X, yValues[yValues.Count / 2], 0));
            }

            return result;
        }

        private static List<TSG.Point> RemoveVeryShortSegmentsFromOpenPolyline(
            List<TSG.Point> points,
            double minimumSegmentLengthMillimeters
        ) {
            if (points == null || points.Count < 2) return points ?? new List<TSG.Point>();

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
            if (points == null || points.Count < 3) return points ?? new List<TSG.Point>();

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
            if (points == null || points.Count < 3) return points ?? new List<TSG.Point>();

            var result = new List<TSG.Point>(points);
            var changed = true;

            while (changed && result.Count >= 3) {
                changed = false;

                for (var index = 0; index < result.Count; index++) {
                    var previousPoint = result[(index - 1 + result.Count) % result.Count];
                    var currentPoint = result[index];
                    var nextPoint = result[(index + 1) % result.Count];

                    var crossValue = Math.Abs(
                        (currentPoint.X - previousPoint.X) * (nextPoint.Y - previousPoint.Y) -
                        (currentPoint.Y - previousPoint.Y) * (nextPoint.X - previousPoint.X)
                    );

                    if (crossValue > toleranceMillimeters) continue;

                    result.RemoveAt(index);
                    changed = true;
                    break;
                }
            }

            return result;
        }

        #endregion

        #region Nested Types

        private sealed class EdgeDetectionResult {
            public string SourceName { get; set; }
            public Dictionary<int, Tuple<TSG.Point, TSG.Point>> EdgesByNumber { get; set; }
            public string ErrorMessage { get; set; }

            public bool HasEdges => EdgesByNumber != null && EdgesByNumber.Count > 0;
        }

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

        #endregion
    }
}