using System;
using System.Collections.Generic;
using System.Linq;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Shared Constants

        private const double DuplicateToleranceMillimeters = 0.5;
        private const double SweepStepMillimeters = 1.0;
        private const double SignificantCornerAngleDegrees = 15.0;
        private const double NearStraightAngleDegrees = 170.0;
        private const double JoinToleranceMillimeters = 0.5;
        private const double MinimumNumberedEdgeLengthMillimeters = 100.0;
        private const double MinimumPartSizeForDimensionMillimeters = 10.0;
        private const int MaximumInlineEdgeCount = 10;
        private const double NumberOffsetMillimeters = 5.0;

        #endregion

        #region Shared Data Structures

        internal sealed class DrawingPartWithBounds {
            public TSD.Part DrawingPart { get; set; }
        }

        private sealed class PartBounds {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
        }

        private sealed class OutlineSegmentGroup {
            public bool IsPolyline { get; set; }
            public List<TSG.Point> Points { get; } = new List<TSG.Point>();
        }

        private sealed class PartOutlineSnapshot {
            public PartBounds Bounds { get; set; }
            public List<TSG.Point> Vertices { get; set; }
            public List<OutlineSegmentGroup> SegmentGroups { get; } = new List<OutlineSegmentGroup>();
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

        #region Shared View And Selection Helpers

        private static bool HasExactlyOnePart(TSD.View selectedView) {
            var modelParts = GetModelPartsFromDrawingView(selectedView);
            return modelParts != null && modelParts.Count == 1;
        }

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
                var firstCoordinateSystem = firstView.DisplayCoordinateSystem;
                var secondCoordinateSystem = secondView.DisplayCoordinateSystem;

                if (firstCoordinateSystem == null || secondCoordinateSystem == null) return false;

                return
                    Math.Abs(firstCoordinateSystem.Origin.X - secondCoordinateSystem.Origin.X) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.Origin.Y - secondCoordinateSystem.Origin.Y) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.Origin.Z - secondCoordinateSystem.Origin.Z) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisX.X - secondCoordinateSystem.AxisX.X) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisX.Y - secondCoordinateSystem.AxisX.Y) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisX.Z - secondCoordinateSystem.AxisX.Z) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisY.X - secondCoordinateSystem.AxisY.X) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisY.Y - secondCoordinateSystem.AxisY.Y) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisY.Z - secondCoordinateSystem.AxisY.Z) < 0.001;
            }
            catch {
                return false;
            }
        }

        private static DrawingPartWithBounds GetSingleDrawingPartFromPickedView(TSD.View pickedView) {
            if (pickedView == null) return null;

            var iterator = pickedView.GetAllObjects(typeof(TSD.ModelObject));
            if (iterator == null) return null;

            DrawingPartWithBounds found = null;
            var foundModelIds = new HashSet<int>();

            while (iterator.MoveNext()) {
                if (!(iterator.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                if (!(MyModel.SelectModelObject(drawingPart.ModelIdentifier) is TSM.Part modelPart)) continue;
                if (!foundModelIds.Add(modelPart.Identifier.ID)) continue;

                if (found != null)
                    return null;

                found = new DrawingPartWithBounds {
                    DrawingPart = drawingPart
                };
            }

            return found;
        }

        #endregion

        #region Shared Bounds Helpers

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

        private static PartBounds GetAssemblyBounds(List<PartBounds> parts) {
            var valid = parts?.Where(part => part != null).ToList();
            if (valid == null || valid.Count == 0) return null;

            return new PartBounds {
                MinX = valid.Min(part => part.MinX),
                MaxX = valid.Max(part => part.MaxX),
                MinY = valid.Min(part => part.MinY),
                MaxY = valid.Max(part => part.MaxY)
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

        #region Shared Outline Extraction

        private static List<PartOutlineSnapshot> GetPartOutlineSnapshots(
            TSD.View view,
            List<DrawingPartWithBounds> drawingParts
        ) {
            var result = new List<PartOutlineSnapshot>();
            if (view == null || drawingParts == null) return result;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(view.DisplayCoordinateSystem)
                );

                foreach (var dp in drawingParts.Where(part => part?.DrawingPart != null)) {
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

                    var vertices = GetOpenOutlineVertices(outline);
                    if (vertices == null || vertices.Count < 2) continue;

                    var snapshot = new PartOutlineSnapshot {
                        Bounds = bounds,
                        Vertices = vertices.Select(point => new TSG.Point(point.X, point.Y, 0)).ToList()
                    };

                    var segmentGroups = BuildOutlineSegmentGroups(vertices);
                    foreach (var group in segmentGroups)
                        snapshot.SegmentGroups.Add(group);

                    result.Add(snapshot);
                }
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }

            return result;
        }

        private static List<OutlineSegmentGroup> BuildOutlineSegmentGroups(List<TSG.Point> vertices) {
            var result = new List<OutlineSegmentGroup>();
            if (vertices == null || vertices.Count < 2) return result;

            if (vertices.Count == 2) {
                var singleGroup = new OutlineSegmentGroup {
                    IsPolyline = false
                };

                singleGroup.Points.Add(new TSG.Point(vertices[0].X, vertices[0].Y, 0));
                singleGroup.Points.Add(new TSG.Point(vertices[1].X, vertices[1].Y, 0));
                result.Add(singleGroup);

                return result;
            }

            var significantCornerIndices = GetSignificantCornerIndices(vertices);

            if (significantCornerIndices.Count <= 1) {
                var polylineGroup = new OutlineSegmentGroup {
                    IsPolyline = true
                };

                foreach (var point in vertices)
                    polylineGroup.Points.Add(new TSG.Point(point.X, point.Y, 0));

                result.Add(polylineGroup);
                return result;
            }

            for (var index = 0; index < significantCornerIndices.Count; index++) {
                var startIndex = significantCornerIndices[index];
                var endIndex = significantCornerIndices[(index + 1) % significantCornerIndices.Count];

                var path = CollectPathBetweenCorners(vertices, startIndex, endIndex);
                path = PrepareSegmentPath(path);

                if (path == null || path.Count < 2) continue;

                var group = new OutlineSegmentGroup {
                    IsPolyline = path.Count > 2
                };

                foreach (var point in path)
                    group.Points.Add(new TSG.Point(point.X, point.Y, 0));

                result.Add(group);
            }

            return result;
        }

        #endregion

        #region Shared Edge Numbering

        private static double ComputeGroupLength(NumberedEdgeGroup group) {
            if (group == null) return 0.0;

            if (group.IsPolyline) {
                if (group.PolylinePoints == null || group.PolylinePoints.Count < 2) return 0.0;

                var length = 0.0;
                for (var index = 0; index < group.PolylinePoints.Count - 1; index++)
                    length += ComputeDistance2D(group.PolylinePoints[index], group.PolylinePoints[index + 1]);

                return length;
            }

            if (group.EdgeSegments == null || group.EdgeSegments.Count == 0) return 0.0;

            var segment = group.EdgeSegments[0];
            return ComputeDistance2D(segment.StartPoint, segment.EndPoint);
        }

        private static Dictionary<int, NumberedEdgeGroup> FilterShortNumberedEdgeGroups(
            Dictionary<int, NumberedEdgeGroup> groupsByNumber,
            double minimumLengthMillimeters
        ) {
            var result = new Dictionary<int, NumberedEdgeGroup>();
            if (groupsByNumber == null || groupsByNumber.Count == 0) return result;

            var newGroupNumber = 0;

            foreach (var pair in groupsByNumber.OrderBy(item => item.Key)) {
                var group = pair.Value;
                if (group == null) continue;

                var length = ComputeGroupLength(group);
                if (length < minimumLengthMillimeters) continue;

                newGroupNumber++;

                var filteredGroup = new NumberedEdgeGroup {
                    GroupNumber = newGroupNumber,
                    IsPolyline = group.IsPolyline,
                    SectionEdge = group.SectionEdge == null
                        ? null
                        : Tuple.Create(
                            new TSG.Point(group.SectionEdge.Item1.X, group.SectionEdge.Item1.Y, 0),
                            new TSG.Point(group.SectionEdge.Item2.X, group.SectionEdge.Item2.Y, 0)
                        )
                };

                foreach (var edgeSegment in group.EdgeSegments)
                    filteredGroup.EdgeSegments.Add(new EdgeSegment {
                        EdgeNumber = edgeSegment.EdgeNumber,
                        StartPoint = new TSG.Point(edgeSegment.StartPoint.X, edgeSegment.StartPoint.Y, 0),
                        EndPoint = new TSG.Point(edgeSegment.EndPoint.X, edgeSegment.EndPoint.Y, 0)
                    });

                foreach (var polylinePoint in group.PolylinePoints)
                    filteredGroup.PolylinePoints.Add(new TSG.Point(polylinePoint.X, polylinePoint.Y, 0));

                result[newGroupNumber] = filteredGroup;
            }

            return result;
        }

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> BuildEdgesByNumberFromOutlines(
            List<PartOutlineSnapshot> outlineSnapshots
        ) {
            var result = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (outlineSnapshots == null || outlineSnapshots.Count == 0) return result;

            var number = 0;

            foreach (var group in outlineSnapshots
                         .Where(snapshot => snapshot?.SegmentGroups != null && snapshot.SegmentGroups.Count != 0)
                         .SelectMany(snapshot =>
                             snapshot.SegmentGroups.Where(currentGroup =>
                                 currentGroup?.Points != null && currentGroup.Points.Count >= 2))) {
                if (!group.IsPolyline || group.Points.Count == 2) {
                    number++;
                    result[number] = Tuple.Create(
                        new TSG.Point(group.Points[0].X, group.Points[0].Y, 0),
                        new TSG.Point(group.Points[1].X, group.Points[1].Y, 0)
                    );

                    continue;
                }

                for (var index = 0; index < group.Points.Count - 1; index++) {
                    number++;
                    result[number] = Tuple.Create(
                        new TSG.Point(group.Points[index].X, group.Points[index].Y, 0),
                        new TSG.Point(group.Points[index + 1].X, group.Points[index + 1].Y, 0)
                    );
                }
            }

            return result;
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

                    for (var index = start; index <= end; index++)
                        result.Add(index);
                }
                else {
                    if (int.TryParse(token, out var number))
                        result.Add(number);
                }

            return result;
        }

        private static HashSet<int> ParseGroupNumbers(string input) {
            return ParseEdgeNumbers(input);
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

            foreach (var polylineGroup in polylineGroups.Where(group =>
                         group?.EdgeSegments != null && group.EdgeSegments.Count != 0)) {
                groupNumber++;

                var numberedGroup = new NumberedEdgeGroup {
                    GroupNumber = groupNumber,
                    IsPolyline = polylineGroup.EdgeSegments.Count > 1
                };

                foreach (var edgeSegment in polylineGroup.EdgeSegments)
                    numberedGroup.EdgeSegments.Add(edgeSegment);

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

        #endregion

        #region Shared Geometry Helpers

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

            if (firstVectorLength < minimumSegmentLength || secondVectorLength < minimumSegmentLength)
                return 180.0;

            var dotProduct = firstVectorX * secondVectorX + firstVectorY * secondVectorY;
            var cosineValue = dotProduct / (firstVectorLength * secondVectorLength);

            if (cosineValue > 1.0) cosineValue = 1.0;
            if (cosineValue < -1.0) cosineValue = -1.0;

            return Math.Acos(cosineValue) * 180.0 / Math.PI;
        }

        private static List<TSG.Point> RemoveNearDuplicates(List<TSG.Point> points, double epsilon) {
            if (points == null || points.Count == 0) return new List<TSG.Point>();

            var seen = new HashSet<(long, long)>();

            return (from point in points
                let key = ((long)Math.Round(point.X / epsilon), (long)Math.Round(point.Y / epsilon))
                where seen.Add(key)
                select point).ToList();
        }

        private static List<TSG.Point> RemoveDuplicatePointsByDistance(
            IEnumerable<TSG.Point> points,
            double tolerance
        ) {
            var result = new List<TSG.Point>();
            if (points == null) return result;

            foreach (var point in points.Where(currentPoint => currentPoint != null)) {
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

        private static List<TSG.Point> GetLineAndPolylineEndpoints(TSD.View view) {
            var result = new List<TSG.Point>();
            if (view == null) return result;

            var iterator = view.GetAllObjects(typeof(TSD.Line));
            while (iterator?.MoveNext() == true)
                if (iterator.Current is TSD.Line line) {
                    result.Add(new TSG.Point(line.StartPoint.X, line.StartPoint.Y, 0));
                    result.Add(new TSG.Point(line.EndPoint.X, line.EndPoint.Y, 0));
                }

            iterator = view.GetAllObjects(typeof(TSD.Polyline));
            while (iterator?.MoveNext() == true) {
                if (!(iterator.Current is TSD.Polyline polyline)) continue;

                var points = new List<TSG.Point>();
                foreach (var item in polyline.Points)
                    if (item is TSG.Point point)
                        points.Add(point);

                if (points.Count == 0) continue;

                result.Add(new TSG.Point(points[0].X, points[0].Y, 0));
                if (points.Count > 1)
                    result.Add(new TSG.Point(points[points.Count - 1].X, points[points.Count - 1].Y, 0));
            }

            return result;
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

            for (var index = 0; index < vertices.Count; index++) {
                var previousIndex = index == 0 ? vertices.Count - 1 : index - 1;
                var nextIndex = index == vertices.Count - 1 ? 0 : index + 1;

                var angleDegrees = GetCornerAngleDegrees(
                    vertices[previousIndex],
                    vertices[index],
                    vertices[nextIndex]
                );

                if (angleDegrees >= SignificantCornerAngleDegrees)
                    result.Add(index);
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

        private static List<TSG.Point> PrepareSegmentPath(List<TSG.Point> points) {
            if (points == null) return null;

            var result = RemoveNearDuplicates(new List<TSG.Point>(points), DuplicateToleranceMillimeters);
            result = SimplifyPolyline(result);
            result = RemoveNearDuplicates(result, DuplicateToleranceMillimeters);

            return result;
        }

        private static List<TSG.Point> BuildUpperLowerChainOutline(List<TSG.Point> points) {
            if (points == null || points.Count < 2) return null;

            var byX = points
                .GroupBy(point => point.X)
                .OrderBy(group => group.Key)
                .ToList();

            if (byX.Count < 2) return null;

            var upper = byX.Select(group => new TSG.Point(group.Key, group.Max(point => point.Y), 0)).ToList();
            var lower = byX.Select(group => new TSG.Point(group.Key, group.Min(point => point.Y), 0)).ToList();

            var outline = new List<TSG.Point>();
            outline.AddRange(upper);
            outline.AddRange(Enumerable.Reverse(lower));

            return outline;
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
                        for (var innerIndex = index; innerIndex < current.Count; innerIndex++)
                            if (next.Count == 0 || !ArePointsEqual(next[next.Count - 1], current[innerIndex]))
                                next.Add(current[innerIndex]);

                        break;
                    }
                }

                current = RemoveNearDuplicates(next, DuplicateToleranceMillimeters);
            }

            return current;
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
                    if (enumerator.Current is TSG.Point point)
                        points.Add(new TSG.Point(xValue, point.Y, 0));
            }

            points = RemoveNearDuplicates(points, DuplicateToleranceMillimeters);
            return points;
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
    }
}