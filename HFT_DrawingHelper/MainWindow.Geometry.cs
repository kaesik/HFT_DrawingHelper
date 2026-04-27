using System;
using System.Collections.Generic;
using System.Linq;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TSS = Tekla.Structures.Solid;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Shared Constants

        private const double DuplicateToleranceMillimeters = 0.5;
        private const double SweepStepMillimeters = 1.0;
        private const double SignificantCornerAngleDegrees = 15.0;
        private const double NearStraightAngleDegrees = 170.0;
        private const double JoinToleranceMillimeters = 0.5;
        private const double BoundarySideOffsetMillimeters = 0.75;
        private const double MinimumNumberedEdgeLengthMillimeters = 100.0;
        private const double MinimumPartSizeForDimensionMillimeters = 10.0;
        private const int MaximumInlineEdgeCount = 10;
        private const double NumberOffsetMillimeters = 5.0;
        private const double SmallGapJoinToleranceMillimeters = 10.0;

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

        private sealed class SegmentEndpointReference {
            public OutlineSegmentGroup SegmentGroup { get; set; }
            public bool IsStartPoint { get; set; }

            public TSG.Point GetPoint() {
                if (SegmentGroup?.Points == null || SegmentGroup.Points.Count == 0)
                    return null;

                return IsStartPoint
                    ? SegmentGroup.Points[0]
                    : SegmentGroup.Points[SegmentGroup.Points.Count - 1];
            }

            public void SetPoint(TSG.Point point) {
                if (SegmentGroup?.Points == null || SegmentGroup.Points.Count == 0 || point == null)
                    return;

                var index = IsStartPoint ? 0 : SegmentGroup.Points.Count - 1;
                SegmentGroup.Points[index] = new TSG.Point(point.X, point.Y, 0);
            }
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

        private sealed class ProjectedFaceLoop {
            public List<TSG.Point> Points { get; } = new List<TSG.Point>();
        }

        private sealed class ProjectedSegment2D {
            public TSG.Point StartPoint { get; set; }
            public TSG.Point EndPoint { get; set; }
        }

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

                    var boundaryPaths = GetProjectedBoundaryPathsFromSolid(solid);
                    if (boundaryPaths == null || boundaryPaths.Count == 0)
                        boundaryPaths = GetLegacyBoundaryPaths(modelPart, solid);

                    if (boundaryPaths == null || boundaryPaths.Count == 0) continue;

                    var snapshot = new PartOutlineSnapshot {
                        Bounds = bounds,
                        Vertices = new List<TSG.Point>()
                    };

                    var primaryBoundaryPath = SelectPrimaryOuterBoundaryPath(boundaryPaths);
                    if (primaryBoundaryPath == null || primaryBoundaryPath.Count < 2)
                        continue;

                    var vertices = primaryBoundaryPath
                        .Select(point => new TSG.Point(point.X, point.Y, 0))
                        .ToList();

                    snapshot.Vertices = vertices
                        .Select(point => new TSG.Point(point.X, point.Y, 0))
                        .ToList();

                    var segmentGroups = BuildOutlineSegmentGroups(vertices);
                    foreach (var group in segmentGroups)
                        snapshot.SegmentGroups.Add(group);

                    CloseSmallGapsInSegmentGroups(
                        snapshot.SegmentGroups,
                        SmallGapJoinToleranceMillimeters
                    );

                    if (snapshot.SegmentGroups.Count == 0) continue;
                    if (snapshot.Vertices == null || snapshot.Vertices.Count == 0)
                        snapshot.Vertices = snapshot.SegmentGroups
                            .SelectMany(group => group.Points)
                            .Select(point => new TSG.Point(point.X, point.Y, 0))
                            .ToList();

                    result.Add(snapshot);
                }
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }

            return result;
        }

        private static void CloseSmallGapsInSegmentGroups(
            List<OutlineSegmentGroup> segmentGroups,
            double toleranceMillimeters
        ) {
            if (segmentGroups == null || segmentGroups.Count == 0)
                return;

            SnapNearbySegmentEndpoints(segmentGroups, toleranceMillimeters);
            SnapSegmentEndpointsToNearestSegments(segmentGroups, toleranceMillimeters);
            NormalizeSegmentGroups(segmentGroups);
            MergeSegmentGroupsByInsignificantAngle(
                segmentGroups,
                toleranceMillimeters,
                NearStraightAngleDegrees
            );
            NormalizeSegmentGroups(segmentGroups);
        }

        private static void SnapNearbySegmentEndpoints(
            List<OutlineSegmentGroup> segmentGroups,
            double toleranceMillimeters
        ) {
            var endpointReferences = GetSegmentEndpointReferences(segmentGroups);
            if (endpointReferences.Count < 2)
                return;

            var visited = new bool[endpointReferences.Count];

            for (var index = 0; index < endpointReferences.Count; index++) {
                if (visited[index])
                    continue;

                var clusterIndices = new List<int>();
                var queue = new Queue<int>();

                queue.Enqueue(index);
                visited[index] = true;

                while (queue.Count > 0) {
                    var currentIndex = queue.Dequeue();
                    clusterIndices.Add(currentIndex);

                    var currentPoint = endpointReferences[currentIndex].GetPoint();
                    if (currentPoint == null)
                        continue;

                    for (var candidateIndex = 0; candidateIndex < endpointReferences.Count; candidateIndex++) {
                        if (visited[candidateIndex])
                            continue;

                        var candidatePoint = endpointReferences[candidateIndex].GetPoint();
                        if (candidatePoint == null)
                            continue;

                        if (ComputeDistance2D(currentPoint, candidatePoint) > toleranceMillimeters)
                            continue;

                        visited[candidateIndex] = true;
                        queue.Enqueue(candidateIndex);
                    }
                }

                if (clusterIndices.Count < 2)
                    continue;

                var averageX = 0.0;
                var averageY = 0.0;
                var validPointsCount = 0;

                foreach (var clusterIndex in clusterIndices) {
                    var point = endpointReferences[clusterIndex].GetPoint();
                    if (point == null)
                        continue;

                    averageX += point.X;
                    averageY += point.Y;
                    validPointsCount++;
                }

                if (validPointsCount < 2)
                    continue;

                var snappedPoint = new TSG.Point(
                    averageX / validPointsCount,
                    averageY / validPointsCount,
                    0
                );

                foreach (var clusterIndex in clusterIndices)
                    endpointReferences[clusterIndex].SetPoint(snappedPoint);
            }
        }

        private static void SnapSegmentEndpointsToNearestSegments(
            List<OutlineSegmentGroup> segmentGroups,
            double toleranceMillimeters
        ) {
            var endpointReferences = GetSegmentEndpointReferences(segmentGroups);
            if (endpointReferences.Count == 0)
                return;

            foreach (var endpointReference in endpointReferences) {
                var currentPoint = endpointReference.GetPoint();
                if (currentPoint == null)
                    continue;

                var bestDistance = double.MaxValue;
                TSG.Point bestPoint = null;

                foreach (var segmentGroup in segmentGroups.Where(group =>
                             group?.Points != null && group.Points.Count >= 2))
                    for (var index = 0; index < segmentGroup.Points.Count - 1; index++) {
                        var segmentStartPoint = segmentGroup.Points[index];
                        var segmentEndPoint = segmentGroup.Points[index + 1];

                        if (segmentStartPoint == null || segmentEndPoint == null)
                            continue;

                        if (ReferenceEquals(segmentGroup, endpointReference.SegmentGroup)) {
                            var isOwnStartSegment = endpointReference.IsStartPoint && index == 0;
                            var isOwnEndSegment = !endpointReference.IsStartPoint &&
                                                  index == segmentGroup.Points.Count - 2;

                            if (isOwnStartSegment || isOwnEndSegment)
                                continue;
                        }

                        var closestPoint = GetClosestPointOnSegment2D(
                            currentPoint,
                            segmentStartPoint,
                            segmentEndPoint
                        );

                        if (closestPoint == null)
                            continue;

                        var distance = ComputeDistance2D(currentPoint, closestPoint);
                        if (distance <= DuplicateToleranceMillimeters ||
                            distance > toleranceMillimeters ||
                            distance >= bestDistance)
                            continue;

                        bestDistance = distance;
                        bestPoint = closestPoint;
                    }

                if (bestPoint != null)
                    endpointReference.SetPoint(bestPoint);
            }
        }

        private static void NormalizeSegmentGroups(List<OutlineSegmentGroup> segmentGroups) {
            if (segmentGroups == null || segmentGroups.Count == 0)
                return;

            var groupsToRemove = new List<OutlineSegmentGroup>();

            foreach (var segmentGroup in segmentGroups) {
                if (segmentGroup?.Points == null)
                    continue;

                var normalizedPoints = RemoveSequentialDuplicatePoints(segmentGroup.Points);
                normalizedPoints = RemoveNearDuplicates(normalizedPoints, DuplicateToleranceMillimeters);

                if (normalizedPoints == null || normalizedPoints.Count < 2) {
                    groupsToRemove.Add(segmentGroup);
                    continue;
                }

                segmentGroup.Points.Clear();
                foreach (var normalizedPoint in normalizedPoints)
                    segmentGroup.Points.Add(new TSG.Point(normalizedPoint.X, normalizedPoint.Y, 0));
            }

            foreach (var groupToRemove in groupsToRemove)
                segmentGroups.Remove(groupToRemove);
        }

        private static void MergeSegmentGroupsByInsignificantAngle(
            List<OutlineSegmentGroup> segmentGroups,
            double joinToleranceMillimeters,
            double nearStraightAngleDegrees
        ) {
            if (segmentGroups == null || segmentGroups.Count < 2)
                return;

            var merged = true;

            while (merged) {
                merged = false;

                for (var firstIndex = 0; firstIndex < segmentGroups.Count; firstIndex++) {
                    var firstGroup = segmentGroups[firstIndex];
                    if (firstGroup?.Points == null || firstGroup.Points.Count < 2)
                        continue;

                    for (var secondIndex = firstIndex + 1; secondIndex < segmentGroups.Count; secondIndex++) {
                        var secondGroup = segmentGroups[secondIndex];
                        if (secondGroup?.Points == null || secondGroup.Points.Count < 2)
                            continue;

                        if (!TryMergeSegmentGroups(
                                firstGroup,
                                secondGroup,
                                joinToleranceMillimeters,
                                nearStraightAngleDegrees,
                                out var mergedGroup))
                            continue;

                        segmentGroups.RemoveAt(secondIndex);
                        segmentGroups.RemoveAt(firstIndex);
                        segmentGroups.Add(mergedGroup);

                        merged = true;
                        break;
                    }

                    if (merged)
                        break;
                }
            }
        }

        private static bool TryMergeSegmentGroups(
            OutlineSegmentGroup firstGroup,
            OutlineSegmentGroup secondGroup,
            double joinToleranceMillimeters,
            double nearStraightAngleDegrees,
            out OutlineSegmentGroup mergedGroup
        ) {
            mergedGroup = null;

            var firstPoints = ClonePoints(firstGroup?.Points);
            var secondPoints = ClonePoints(secondGroup?.Points);

            if (firstPoints == null || secondPoints == null || firstPoints.Count < 2 || secondPoints.Count < 2)
                return false;

            if (TryMergeSegmentGroupOrientation(
                    firstPoints,
                    secondPoints,
                    joinToleranceMillimeters,
                    nearStraightAngleDegrees,
                    out mergedGroup))
                return true;

            if (TryMergeSegmentGroupOrientation(
                    firstPoints,
                    ReversePoints(secondPoints),
                    joinToleranceMillimeters,
                    nearStraightAngleDegrees,
                    out mergedGroup))
                return true;

            if (TryMergeSegmentGroupOrientation(
                    ReversePoints(firstPoints),
                    secondPoints,
                    joinToleranceMillimeters,
                    nearStraightAngleDegrees,
                    out mergedGroup))
                return true;

            return TryMergeSegmentGroupOrientation(
                secondPoints,
                firstPoints,
                joinToleranceMillimeters,
                nearStraightAngleDegrees,
                out mergedGroup
            );
        }

        private static bool TryMergeSegmentGroupOrientation(
            List<TSG.Point> firstPath,
            List<TSG.Point> secondPath,
            double joinToleranceMillimeters,
            double nearStraightAngleDegrees,
            out OutlineSegmentGroup mergedGroup
        ) {
            mergedGroup = null;

            if (firstPath == null || secondPath == null || firstPath.Count < 2 || secondPath.Count < 2)
                return false;

            var firstEndPoint = firstPath[firstPath.Count - 1];
            var secondStartPoint = secondPath[0];

            if (ComputeDistance2D(firstEndPoint, secondStartPoint) > joinToleranceMillimeters)
                return false;

            if (!ShouldMergeInsignificantCorner(
                    firstPath,
                    secondPath,
                    joinToleranceMillimeters,
                    nearStraightAngleDegrees))
                return false;

            var mergedPoints = BuildMergedPath(firstPath, secondPath);
            mergedPoints = PrepareBoundaryPath(mergedPoints);

            if (mergedPoints == null || mergedPoints.Count < 2)
                return false;

            mergedGroup = new OutlineSegmentGroup {
                IsPolyline = mergedPoints.Count > 2
            };

            foreach (var mergedPoint in mergedPoints)
                mergedGroup.Points.Add(new TSG.Point(mergedPoint.X, mergedPoint.Y, 0));

            return true;
        }

        private static bool ShouldMergeInsignificantCorner(
            List<TSG.Point> firstPath,
            List<TSG.Point> secondPath,
            double joinToleranceMillimeters,
            double nearStraightAngleDegrees
        ) {
            if (firstPath == null || secondPath == null || firstPath.Count < 2 || secondPath.Count < 2)
                return false;

            var previousPoint = firstPath[firstPath.Count - 2];
            var firstEndPoint = firstPath[firstPath.Count - 1];
            var secondStartPoint = secondPath[0];
            var nextPoint = secondPath[1];

            if (ComputeDistance2D(firstEndPoint, secondStartPoint) > joinToleranceMillimeters)
                return false;

            var firstLength = ComputeDistance2D(previousPoint, firstEndPoint);
            var secondLength = ComputeDistance2D(secondStartPoint, nextPoint);

            if (firstLength <= joinToleranceMillimeters || secondLength <= joinToleranceMillimeters)
                return true;

            var connectionPoint = new TSG.Point(
                (firstEndPoint.X + secondStartPoint.X) * 0.5,
                (firstEndPoint.Y + secondStartPoint.Y) * 0.5,
                0
            );

            var angleInDegrees = ComputeAngleInDegrees(
                previousPoint,
                connectionPoint,
                nextPoint,
                1e-6
            );

            return angleInDegrees >= nearStraightAngleDegrees;
        }

        private static List<TSG.Point> BuildMergedPath(
            List<TSG.Point> firstPath,
            List<TSG.Point> secondPath
        ) {
            var result = new List<TSG.Point>();
            if (firstPath == null || secondPath == null || firstPath.Count == 0 || secondPath.Count == 0)
                return result;

            for (var index = 0; index < firstPath.Count - 1; index++)
                result.Add(new TSG.Point(firstPath[index].X, firstPath[index].Y, 0));

            var mergedConnectionPoint = new TSG.Point(
                (firstPath[firstPath.Count - 1].X + secondPath[0].X) * 0.5,
                (firstPath[firstPath.Count - 1].Y + secondPath[0].Y) * 0.5,
                0
            );

            result.Add(mergedConnectionPoint);

            for (var index = 1; index < secondPath.Count; index++)
                result.Add(new TSG.Point(secondPath[index].X, secondPath[index].Y, 0));

            return result;
        }

        private static List<TSG.Point> ClonePoints(IEnumerable<TSG.Point> points) {
            var result = new List<TSG.Point>();
            if (points == null)
                return result;

            foreach (var point in points) {
                if (point == null)
                    continue;

                result.Add(new TSG.Point(point.X, point.Y, 0));
            }

            return result;
        }

        private static List<TSG.Point> ReversePoints(IEnumerable<TSG.Point> points) {
            var result = ClonePoints(points);
            result.Reverse();
            return result;
        }

        private static List<SegmentEndpointReference> GetSegmentEndpointReferences(
            List<OutlineSegmentGroup> segmentGroups
        ) {
            var result = new List<SegmentEndpointReference>();
            if (segmentGroups == null)
                return result;

            foreach (var segmentGroup in segmentGroups.Where(group =>
                         group?.Points != null && group.Points.Count >= 2)) {
                result.Add(new SegmentEndpointReference {
                    SegmentGroup = segmentGroup,
                    IsStartPoint = true
                });

                result.Add(new SegmentEndpointReference {
                    SegmentGroup = segmentGroup,
                    IsStartPoint = false
                });
            }

            return result;
        }

        private static TSG.Point GetClosestPointOnSegment2D(
            TSG.Point point,
            TSG.Point segmentStartPoint,
            TSG.Point segmentEndPoint
        ) {
            if (point == null || segmentStartPoint == null || segmentEndPoint == null)
                return null;

            var segmentX = segmentEndPoint.X - segmentStartPoint.X;
            var segmentY = segmentEndPoint.Y - segmentStartPoint.Y;
            var segmentLengthSquared = segmentX * segmentX + segmentY * segmentY;

            if (segmentLengthSquared < 1e-12)
                return new TSG.Point(segmentStartPoint.X, segmentStartPoint.Y, 0);

            var t =
                ((point.X - segmentStartPoint.X) * segmentX +
                 (point.Y - segmentStartPoint.Y) * segmentY) / segmentLengthSquared;

            if (t < 0.0) t = 0.0;
            if (t > 1.0) t = 1.0;

            return new TSG.Point(
                segmentStartPoint.X + t * segmentX,
                segmentStartPoint.Y + t * segmentY,
                0
            );
        }

        private static List<TSG.Point> BuildDimensionReferencePoints(PartOutlineSnapshot outlineSnapshot) {
            var result = new List<TSG.Point>();
            if (outlineSnapshot == null)
                return result;

            foreach (var segmentGroup in outlineSnapshot.SegmentGroups.Where(group =>
                         group?.Points != null && group.Points.Count >= 2)) {
                var groupPoints = RemoveSequentialDuplicatePoints(segmentGroup.Points);
                groupPoints = RemoveNearDuplicates(groupPoints, DuplicateToleranceMillimeters);

                if (groupPoints == null || groupPoints.Count == 0)
                    continue;

                result.Add(new TSG.Point(groupPoints[0].X, groupPoints[0].Y, 0));

                if (groupPoints.Count > 1)
                    result.Add(new TSG.Point(
                        groupPoints[groupPoints.Count - 1].X,
                        groupPoints[groupPoints.Count - 1].Y,
                        0
                    ));

                if (groupPoints.Count > 2)
                    result.AddRange(GetSignificantOutlineReferencePoints(groupPoints));
                else
                    foreach (var groupPoint in groupPoints)
                        result.Add(new TSG.Point(groupPoint.X, groupPoint.Y, 0));
            }

            if (result.Count == 0 && outlineSnapshot.Vertices != null && outlineSnapshot.Vertices.Count > 0)
                result.AddRange(GetSignificantOutlineReferencePoints(outlineSnapshot.Vertices));

            return RemoveDuplicatePointsByDistance(result, DimensionMergeToleranceMillimeters);
        }


        private static List<List<TSG.Point>> GetProjectedBoundaryPathsFromSolid(TSM.Solid solid) {
            var result = new List<List<TSG.Point>>();
            if (solid == null) return result;

            var faceLoops = GetProjectedFaceLoopsFromSolid(solid);
            if (faceLoops.Count == 0) return result;

            var projectedSegments = GetProjectedSegmentsFromSolid(solid);
            if (projectedSegments.Count == 0) return result;

            var splitSegments = SplitProjectedSegmentsAtIntersections(projectedSegments);
            if (splitSegments.Count == 0) return result;

            var boundarySegments = FilterBoundarySegments(splitSegments, faceLoops);
            if (boundarySegments.Count == 0) return result;

            var boundaryPaths = BuildBoundaryPathsFromSegments(boundarySegments);
            if (boundaryPaths.Count == 0) return result;

            boundaryPaths = JoinBoundaryPathsAcrossSmallGaps(boundaryPaths);
            if (boundaryPaths.Count == 0) return result;

            foreach (var boundaryPath in boundaryPaths) {
                var preparedPath = PrepareBoundaryPath(boundaryPath);
                if (preparedPath != null && preparedPath.Count >= 2)
                    result.Add(preparedPath);
            }

            result = JoinBoundaryPathsAcrossSmallGaps(result)
                .Select(PrepareBoundaryPath)
                .Where(path => path != null && path.Count >= 2)
                .ToList();

            return RemoveInteriorBoundaryPaths(result);
        }

        private static List<List<TSG.Point>> GetLegacyBoundaryPaths(TSM.Part modelPart, TSM.Solid solid) {
            var result = new List<List<TSG.Point>>();
            if (modelPart == null || solid == null) return result;

            List<TSG.Point> outline = null;

            if (ShouldUseSweepOutline(modelPart)) {
                var points = GetSweptPlateOutlinePoints(solid);
                if (points.Count >= 2)
                    outline = BuildUpperLowerChainOutline(points);
            }
            else {
                var backZ = solid.MinimumPoint.Z;
                const double inward = 1.0;

                if (solid.MaximumPoint.Z - solid.MinimumPoint.Z > inward * 2)
                    backZ += inward;
                else
                    backZ = (solid.MinimumPoint.Z + solid.MaximumPoint.Z) * 0.5;

                var backPoints = GetIntersectionPointsAtLocalZ(solid, backZ);
                if (backPoints != null && backPoints.Count >= 3) {
                    backPoints = RemoveNearDuplicates(backPoints, DuplicateToleranceMillimeters);
                    outline = BuildConvexHull2D(backPoints);
                }
            }

            if (outline == null || outline.Count < 3) return result;

            outline = PrepareBoundaryPath(outline);
            outline = EnsureClosedOutline(outline);

            var vertices = GetOpenOutlineVertices(outline);
            if (vertices != null && vertices.Count >= 2)
                result.Add(vertices);

            return result;
        }

        private static List<ProjectedFaceLoop> GetProjectedFaceLoopsFromSolid(TSM.Solid solid) {
            var result = new List<ProjectedFaceLoop>();
            if (solid == null) return result;

            var faceEnumerator = solid.GetFaceEnumerator();
            while (faceEnumerator.MoveNext()) {
                if (!(faceEnumerator.Current is TSS.Face face)) continue;

                var loopEnumerator = face.GetLoopEnumerator();
                while (loopEnumerator.MoveNext()) {
                    if (!(loopEnumerator.Current is TSS.Loop loop)) continue;

                    var points = new List<TSG.Point>();
                    var vertexEnumerator = loop.GetVertexEnumerator();

                    while (vertexEnumerator.MoveNext())
                        if (vertexEnumerator.Current is TSG.Point vertex)
                            points.Add(new TSG.Point(vertex.X, vertex.Y, 0));

                    points = RemoveSequentialDuplicatePoints(points);
                    if (points.Count < 3) continue;

                    if (!ArePointsEqual(points[0], points[points.Count - 1]))
                        points.Add(new TSG.Point(points[0].X, points[0].Y, 0));

                    if (Math.Abs(ComputeSignedArea(points)) <=
                        DuplicateToleranceMillimeters * DuplicateToleranceMillimeters)
                        continue;

                    var faceLoop = new ProjectedFaceLoop();
                    foreach (var point in points)
                        faceLoop.Points.Add(new TSG.Point(point.X, point.Y, 0));

                    result.Add(faceLoop);
                }
            }

            return result;
        }

        private static List<ProjectedSegment2D> GetProjectedSegmentsFromSolid(TSM.Solid solid) {
            var result = new List<ProjectedSegment2D>();
            if (solid == null) return result;

            var seenKeys = new HashSet<string>();

            var faceEnumerator = solid.GetFaceEnumerator();
            while (faceEnumerator.MoveNext()) {
                if (!(faceEnumerator.Current is TSS.Face face)) continue;

                var loopEnumerator = face.GetLoopEnumerator();
                while (loopEnumerator.MoveNext()) {
                    if (!(loopEnumerator.Current is TSS.Loop loop)) continue;

                    var vertices = new List<TSG.Point>();
                    var vertexEnumerator = loop.GetVertexEnumerator();

                    while (vertexEnumerator.MoveNext())
                        if (vertexEnumerator.Current is TSG.Point vertex)
                            vertices.Add(new TSG.Point(vertex.X, vertex.Y, 0));

                    vertices = RemoveSequentialDuplicatePoints(vertices);
                    if (vertices.Count < 2) continue;

                    for (var index = 0; index < vertices.Count; index++) {
                        var startPoint = vertices[index];
                        var endPoint = vertices[(index + 1) % vertices.Count];

                        if (ComputeDistance2D(startPoint, endPoint) <= DuplicateToleranceMillimeters)
                            continue;

                        var normalizedStartPoint = startPoint;
                        var normalizedEndPoint = endPoint;

                        if (ComparePointsLexicographically(normalizedStartPoint, normalizedEndPoint) > 0) {
                            normalizedStartPoint = endPoint;
                            normalizedEndPoint = startPoint;
                        }

                        var key = BuildSegmentKey(normalizedStartPoint, normalizedEndPoint);
                        if (!seenKeys.Add(key)) continue;

                        result.Add(new ProjectedSegment2D {
                            StartPoint = new TSG.Point(normalizedStartPoint.X, normalizedStartPoint.Y, 0),
                            EndPoint = new TSG.Point(normalizedEndPoint.X, normalizedEndPoint.Y, 0)
                        });
                    }
                }
            }

            return result;
        }

        private static List<ProjectedSegment2D> SplitProjectedSegmentsAtIntersections(
            List<ProjectedSegment2D> projectedSegments
        ) {
            var result = new List<ProjectedSegment2D>();
            if (projectedSegments == null || projectedSegments.Count == 0) return result;

            var uniqueKeys = new HashSet<string>();

            for (var index = 0; index < projectedSegments.Count; index++) {
                var currentSegment = projectedSegments[index];
                var splitPoints = new List<TSG.Point> {
                    currentSegment.StartPoint,
                    currentSegment.EndPoint
                };

                for (var otherIndex = 0; otherIndex < projectedSegments.Count; otherIndex++) {
                    if (index == otherIndex) continue;

                    var otherSegment = projectedSegments[otherIndex];
                    AddSegmentIntersections(currentSegment, otherSegment, splitPoints);
                }

                splitPoints = RemoveNearDuplicates(splitPoints, DuplicateToleranceMillimeters);
                splitPoints = splitPoints
                    .OrderBy(point =>
                        GetPointParameterOnSegment(point, currentSegment.StartPoint, currentSegment.EndPoint))
                    .ToList();

                for (var pointIndex = 0; pointIndex < splitPoints.Count - 1; pointIndex++) {
                    var startPoint = splitPoints[pointIndex];
                    var endPoint = splitPoints[pointIndex + 1];

                    if (ComputeDistance2D(startPoint, endPoint) <= DuplicateToleranceMillimeters)
                        continue;

                    var normalizedStartPoint = startPoint;
                    var normalizedEndPoint = endPoint;

                    if (ComparePointsLexicographically(normalizedStartPoint, normalizedEndPoint) > 0) {
                        normalizedStartPoint = endPoint;
                        normalizedEndPoint = startPoint;
                    }

                    var key = BuildSegmentKey(normalizedStartPoint, normalizedEndPoint);
                    if (!uniqueKeys.Add(key)) continue;

                    result.Add(new ProjectedSegment2D {
                        StartPoint = new TSG.Point(normalizedStartPoint.X, normalizedStartPoint.Y, 0),
                        EndPoint = new TSG.Point(normalizedEndPoint.X, normalizedEndPoint.Y, 0)
                    });
                }
            }

            return result;
        }

        private static void AddSegmentIntersections(
            ProjectedSegment2D firstSegment,
            ProjectedSegment2D secondSegment,
            List<TSG.Point> splitPoints
        ) {
            if (firstSegment == null || secondSegment == null || splitPoints == null) return;

            if (TryGetSegmentIntersectionPoint(firstSegment.StartPoint, firstSegment.EndPoint,
                    secondSegment.StartPoint, secondSegment.EndPoint, out var intersectionPoint)) {
                splitPoints.Add(intersectionPoint);
                return;
            }

            if (!AreSegmentsCollinear(firstSegment.StartPoint, firstSegment.EndPoint,
                    secondSegment.StartPoint, secondSegment.EndPoint))
                return;

            if (IsPointOnSegment(firstSegment.StartPoint, secondSegment.StartPoint, secondSegment.EndPoint))
                splitPoints.Add(firstSegment.StartPoint);
            if (IsPointOnSegment(firstSegment.EndPoint, secondSegment.StartPoint, secondSegment.EndPoint))
                splitPoints.Add(firstSegment.EndPoint);
            if (IsPointOnSegment(secondSegment.StartPoint, firstSegment.StartPoint, firstSegment.EndPoint))
                splitPoints.Add(secondSegment.StartPoint);
            if (IsPointOnSegment(secondSegment.EndPoint, firstSegment.StartPoint, firstSegment.EndPoint))
                splitPoints.Add(secondSegment.EndPoint);
        }

        private static List<ProjectedSegment2D> FilterBoundarySegments(
            List<ProjectedSegment2D> splitSegments,
            List<ProjectedFaceLoop> faceLoops
        ) {
            var result = new List<ProjectedSegment2D>();
            if (splitSegments == null || faceLoops == null) return result;

            foreach (var splitSegment in splitSegments) {
                var segmentLength = ComputeDistance2D(splitSegment.StartPoint, splitSegment.EndPoint);
                if (segmentLength <= DuplicateToleranceMillimeters) continue;

                if (!IsProjectedSegmentOnOutlineBoundary(splitSegment, faceLoops))
                    continue;

                result.Add(new ProjectedSegment2D {
                    StartPoint = new TSG.Point(splitSegment.StartPoint.X, splitSegment.StartPoint.Y, 0),
                    EndPoint = new TSG.Point(splitSegment.EndPoint.X, splitSegment.EndPoint.Y, 0)
                });
            }

            return result;
        }

        private static bool IsProjectedSegmentOnOutlineBoundary(
            ProjectedSegment2D splitSegment,
            List<ProjectedFaceLoop> faceLoops
        ) {
            if (splitSegment == null || splitSegment.StartPoint == null || splitSegment.EndPoint == null)
                return false;

            var segmentLength = ComputeDistance2D(splitSegment.StartPoint, splitSegment.EndPoint);
            if (segmentLength <= DuplicateToleranceMillimeters)
                return false;

            var directionX = splitSegment.EndPoint.X - splitSegment.StartPoint.X;
            var directionY = splitSegment.EndPoint.Y - splitSegment.StartPoint.Y;
            var directionLength = Math.Sqrt(directionX * directionX + directionY * directionY);

            if (directionLength < 1e-9)
                return false;

            var normalX = -directionY / directionLength;
            var normalY = directionX / directionLength;
            var offset = Math.Max(0.15, Math.Min(BoundarySideOffsetMillimeters, segmentLength * 0.05));

            var separatingSampleCount = 0;
            var validSampleCount = 0;

            foreach (var fraction in GetProjectedSegmentSampleFractions(segmentLength)) {
                var samplePoint = new TSG.Point(
                    splitSegment.StartPoint.X + (splitSegment.EndPoint.X - splitSegment.StartPoint.X) * fraction,
                    splitSegment.StartPoint.Y + (splitSegment.EndPoint.Y - splitSegment.StartPoint.Y) * fraction,
                    0
                );

                var pointOnLeftSide = new TSG.Point(
                    samplePoint.X + normalX * offset,
                    samplePoint.Y + normalY * offset,
                    0
                );

                var pointOnRightSide = new TSG.Point(
                    samplePoint.X - normalX * offset,
                    samplePoint.Y - normalY * offset,
                    0
                );

                var leftInside = IsPointInsideProjectedSolid(pointOnLeftSide, faceLoops);
                var rightInside = IsPointInsideProjectedSolid(pointOnRightSide, faceLoops);

                if (leftInside == rightInside)
                    continue;

                validSampleCount++;
                separatingSampleCount++;
            }

            if (validSampleCount == 0)
                return false;

            return separatingSampleCount >= Math.Max(1, (int)Math.Ceiling(validSampleCount * 0.8));
        }

        private static List<double> GetProjectedSegmentSampleFractions(double segmentLength) {
            if (segmentLength <= MinimumPartSizeForDimensionMillimeters * 0.5)
                return new List<double> { 0.5 };

            return new List<double> { 0.1, 0.2, 0.35, 0.5, 0.65, 0.8, 0.9 };
        }

        private static bool IsPointInsideProjectedSolid(TSG.Point point, List<ProjectedFaceLoop> faceLoops) {
            if (point == null || faceLoops == null || faceLoops.Count == 0) return false;

            foreach (var faceLoop in faceLoops) {
                if (faceLoop?.Points == null || faceLoop.Points.Count < 3) continue;
                if (IsPointInsidePolygon(point, faceLoop.Points)) return true;
            }

            return false;
        }


        private static List<List<TSG.Point>> JoinBoundaryPathsAcrossSmallGaps(List<List<TSG.Point>> boundaryPaths) {
            var result = boundaryPaths?
                .Where(path => path != null && path.Count >= 2)
                .Select(path => path.Select(point => new TSG.Point(point.X, point.Y, 0)).ToList())
                .ToList() ?? new List<List<TSG.Point>>();

            if (result.Count <= 1)
                return result;

            var changed = true;
            while (changed) {
                changed = false;

                for (var firstIndex = 0; firstIndex < result.Count && !changed; firstIndex++) {
                    var firstPath = PrepareBoundaryPath(result[firstIndex]);
                    if (firstPath == null || firstPath.Count < 2)
                        continue;

                    for (var secondIndex = firstIndex + 1; secondIndex < result.Count && !changed; secondIndex++) {
                        var secondPath = PrepareBoundaryPath(result[secondIndex]);
                        if (secondPath == null || secondPath.Count < 2)
                            continue;

                        if (!TryMergeBoundaryPaths(firstPath, secondPath, out var mergedPath))
                            continue;

                        result[firstIndex] = mergedPath;
                        result.RemoveAt(secondIndex);
                        changed = true;
                    }
                }
            }

            return result
                .Select(PrepareBoundaryPath)
                .Where(path => path != null && path.Count >= 2)
                .ToList();
        }

        private static bool TryMergeBoundaryPaths(
            List<TSG.Point> firstPath,
            List<TSG.Point> secondPath,
            out List<TSG.Point> mergedPath
        ) {
            mergedPath = null;
            if (firstPath == null || secondPath == null || firstPath.Count < 2 || secondPath.Count < 2)
                return false;

            var candidates = new List<(double Distance, int CaseId)> {
                (ComputeDistance2D(firstPath[firstPath.Count - 1], secondPath[0]), 0),
                (ComputeDistance2D(firstPath[firstPath.Count - 1], secondPath[secondPath.Count - 1]), 1),
                (ComputeDistance2D(firstPath[0], secondPath[0]), 2),
                (ComputeDistance2D(firstPath[0], secondPath[secondPath.Count - 1]), 3)
            };

            var bestCandidate = candidates
                .OrderBy(candidate => candidate.Distance)
                .First();

            if (bestCandidate.Distance > SmallGapJoinToleranceMillimeters)
                return false;

            List<TSG.Point> normalizedFirstPath;
            List<TSG.Point> normalizedSecondPath;

            switch (bestCandidate.CaseId) {
                case 0:
                    normalizedFirstPath = ClonePoints(firstPath);
                    normalizedSecondPath = ClonePoints(secondPath);
                    break;
                case 1:
                    normalizedFirstPath = ClonePoints(firstPath);
                    normalizedSecondPath = ClonePoints(secondPath.AsEnumerable().Reverse().ToList());
                    break;
                case 2:
                    normalizedFirstPath = ClonePoints(firstPath.AsEnumerable().Reverse().ToList());
                    normalizedSecondPath = ClonePoints(secondPath);
                    break;
                default:
                    normalizedFirstPath = ClonePoints(firstPath.AsEnumerable().Reverse().ToList());
                    normalizedSecondPath = ClonePoints(secondPath.AsEnumerable().Reverse().ToList());
                    break;
            }

            if (!ShouldMergeInsignificantCorner(
                    normalizedFirstPath,
                    normalizedSecondPath,
                    SmallGapJoinToleranceMillimeters,
                    NearStraightAngleDegrees))
                return false;

            var sharedPoint = new TSG.Point(
                (normalizedFirstPath[normalizedFirstPath.Count - 1].X + normalizedSecondPath[0].X) * 0.5,
                (normalizedFirstPath[normalizedFirstPath.Count - 1].Y + normalizedSecondPath[0].Y) * 0.5,
                0
            );

            normalizedFirstPath[normalizedFirstPath.Count - 1] = sharedPoint;
            normalizedSecondPath[0] = new TSG.Point(sharedPoint.X, sharedPoint.Y, 0);

            var candidatePath = ClonePoints(normalizedFirstPath);
            candidatePath.AddRange(normalizedSecondPath.Skip(1).Select(point => new TSG.Point(point.X, point.Y, 0)));
            candidatePath = PrepareBoundaryPath(candidatePath);

            if (candidatePath == null || candidatePath.Count < 2)
                return false;

            mergedPath = candidatePath;
            return true;
        }

        private static List<List<TSG.Point>> BuildBoundaryPathsFromSegments(List<ProjectedSegment2D> boundarySegments) {
            var result = new List<List<TSG.Point>>();
            if (boundarySegments == null || boundarySegments.Count == 0) return result;

            var usedSegments = new bool[boundarySegments.Count];
            var nodeDegrees = BuildBoundaryNodeDegrees(boundarySegments);

            while (usedSegments.Any(used => !used)) {
                var startSegmentIndex = GetNextUnusedSegmentIndex(usedSegments);
                if (startSegmentIndex < 0) break;

                var currentSegment = boundarySegments[startSegmentIndex];
                var currentPath = new List<TSG.Point> {
                    new TSG.Point(currentSegment.StartPoint.X, currentSegment.StartPoint.Y, 0),
                    new TSG.Point(currentSegment.EndPoint.X, currentSegment.EndPoint.Y, 0)
                };

                usedSegments[startSegmentIndex] = true;

                ExtendBoundaryPath(boundarySegments, usedSegments, currentPath, true, nodeDegrees);
                ExtendBoundaryPath(boundarySegments, usedSegments, currentPath, false, nodeDegrees);

                currentPath = RemoveSequentialDuplicatePoints(currentPath);
                if (currentPath.Count < 2) continue;

                if (ArePointsEqual(currentPath[0], currentPath[currentPath.Count - 1]))
                    currentPath.RemoveAt(currentPath.Count - 1);

                if (currentPath.Count >= 2)
                    result.Add(currentPath);
            }

            return result;
        }

        private static List<List<TSG.Point>> RemoveInteriorBoundaryPaths(List<List<TSG.Point>> boundaryPaths) {
            if (boundaryPaths == null || boundaryPaths.Count == 0)
                return new List<List<TSG.Point>>();

            var preparedPaths = boundaryPaths
                .Select(PrepareBoundaryPath)
                .Where(path => path != null && path.Count >= 2)
                .ToList();

            if (preparedPaths.Count <= 1)
                return preparedPaths;

            var dominantBoundary = GetDominantBoundaryPolygon(preparedPaths);
            if (dominantBoundary == null || dominantBoundary.Count < 4)
                return preparedPaths;

            var keptPaths = new List<List<TSG.Point>>();
            foreach (var path in preparedPaths)
                if (!IsBoundaryPathInsidePolygon(path, dominantBoundary))
                    keptPaths.Add(path);

            return keptPaths.Count > 0 ? keptPaths : preparedPaths;
        }

        private static List<TSG.Point> SelectPrimaryOuterBoundaryPath(List<List<TSG.Point>> boundaryPaths) {
            if (boundaryPaths == null || boundaryPaths.Count == 0)
                return null;

            var preparedPaths = boundaryPaths
                .Select(PrepareBoundaryPath)
                .Where(path => path != null && path.Count >= 2)
                .ToList();

            if (preparedPaths.Count == 0)
                return null;

            var dominantBoundary = GetDominantBoundaryPolygon(preparedPaths);
            if (dominantBoundary != null && dominantBoundary.Count >= 4)
                return GetOpenOutlineVertices(dominantBoundary);

            return preparedPaths
                .OrderByDescending(ComputePolylineLength)
                .FirstOrDefault();
        }

        private static List<TSG.Point> GetDominantBoundaryPolygon(List<List<TSG.Point>> boundaryPaths) {
            if (boundaryPaths == null || boundaryPaths.Count == 0)
                return null;

            return boundaryPaths
                .Select(EnsureClosedOutlineCopy)
                .Where(path => path != null && path.Count >= 4)
                .OrderByDescending(path => Math.Abs(ComputeSignedArea(path)))
                .ThenByDescending(ComputePolylineLength)
                .FirstOrDefault();
        }

        private static bool IsBoundaryPathInsidePolygon(List<TSG.Point> candidatePath, List<TSG.Point> polygon) {
            if (candidatePath == null || candidatePath.Count < 2 || polygon == null || polygon.Count < 4)
                return false;

            var samplePoints = GetBoundaryPathSamplePoints(candidatePath);
            if (samplePoints.Count == 0)
                return false;

            foreach (var samplePoint in samplePoints)
                if (!IsPointStrictlyInsidePolygon(samplePoint, polygon))
                    return false;

            return true;
        }

        private static List<TSG.Point> GetBoundaryPathSamplePoints(List<TSG.Point> path) {
            var result = new List<TSG.Point>();
            if (path == null || path.Count < 2) return result;

            for (var index = 0; index < path.Count - 1; index++) {
                var startPoint = path[index];
                var endPoint = path[index + 1];

                if (ComputeDistance2D(startPoint, endPoint) <= DuplicateToleranceMillimeters)
                    continue;

                result.Add(new TSG.Point(
                    (startPoint.X + endPoint.X) * 0.5,
                    (startPoint.Y + endPoint.Y) * 0.5,
                    0
                ));
            }

            if (result.Count == 0 && path.Count >= 3) {
                var centroidX = path.Average(point => point.X);
                var centroidY = path.Average(point => point.Y);
                result.Add(new TSG.Point(centroidX, centroidY, 0));
            }

            return result;
        }

        private static bool IsPointStrictlyInsidePolygon(TSG.Point point, IReadOnlyList<TSG.Point> polygonPoints) {
            if (point == null || polygonPoints == null || polygonPoints.Count < 3) return false;

            for (var index = 0; index < polygonPoints.Count - 1; index++)
                if (IsPointOnSegment(point, polygonPoints[index], polygonPoints[index + 1]))
                    return false;

            return IsPointInsidePolygon(point, polygonPoints);
        }

        private static List<TSG.Point> EnsureClosedOutlineCopy(List<TSG.Point> outline) {
            if (outline == null) return new List<TSG.Point>();

            var copy = outline
                .Where(point => point != null)
                .Select(point => new TSG.Point(point.X, point.Y, 0))
                .ToList();

            return EnsureClosedOutline(copy);
        }


        private static void ExtendBoundaryPath(
            List<ProjectedSegment2D> boundarySegments,
            bool[] usedSegments,
            List<TSG.Point> currentPath,
            bool appendToEnd,
            Dictionary<string, int> nodeDegrees
        ) {
            if (boundarySegments == null || usedSegments == null || currentPath == null || currentPath.Count < 2)
                return;

            while (true) {
                var bestSegmentIndex = FindBestConnectedBoundarySegmentIndex(
                    boundarySegments,
                    usedSegments,
                    currentPath,
                    appendToEnd,
                    nodeDegrees,
                    out var nextPoint
                );

                if (bestSegmentIndex < 0 || nextPoint == null)
                    break;

                if (appendToEnd)
                    currentPath.Add(new TSG.Point(nextPoint.X, nextPoint.Y, 0));
                else
                    currentPath.Insert(0, new TSG.Point(nextPoint.X, nextPoint.Y, 0));

                usedSegments[bestSegmentIndex] = true;
            }
        }

        private static int FindBestConnectedBoundarySegmentIndex(
            List<ProjectedSegment2D> boundarySegments,
            bool[] usedSegments,
            List<TSG.Point> currentPath,
            bool appendToEnd,
            Dictionary<string, int> nodeDegrees,
            out TSG.Point bestNextPoint
        ) {
            bestNextPoint = null;
            if (boundarySegments == null || usedSegments == null || currentPath == null || currentPath.Count < 2)
                return -1;

            var connectionPoint = appendToEnd ? currentPath[currentPath.Count - 1] : currentPath[0];
            var previousPoint = appendToEnd ? currentPath[currentPath.Count - 2] : currentPath[1];

            var bestIndex = -1;
            var bestClosesPath = true;
            var bestBranchDegree = int.MaxValue;
            var bestLength = double.PositiveInfinity;
            var bestAngle = double.NegativeInfinity;

            for (var index = 0; index < boundarySegments.Count; index++) {
                if (usedSegments[index]) continue;

                var segment = boundarySegments[index];
                TSG.Point candidateNextPoint = null;

                if (ArePointsEqual(connectionPoint, segment.StartPoint))
                    candidateNextPoint = segment.EndPoint;
                else if (ArePointsEqual(connectionPoint, segment.EndPoint))
                    candidateNextPoint = segment.StartPoint;

                if (candidateNextPoint == null) continue;
                if (ArePointsEqual(candidateNextPoint, previousPoint)) continue;

                var closesPath = appendToEnd
                    ? ArePointsEqual(candidateNextPoint, currentPath[0])
                    : ArePointsEqual(candidateNextPoint, currentPath[currentPath.Count - 1]);

                if (!closesPath && IsPointAlreadyUsedInPath(candidateNextPoint, currentPath))
                    continue;

                if (WouldCandidateSegmentIntersectCurrentPath(currentPath, connectionPoint, candidateNextPoint,
                        appendToEnd))
                    continue;

                var branchDegree = GetBoundaryNodeDegree(nodeDegrees, candidateNextPoint);
                var length = ComputeDistance2D(connectionPoint, candidateNextPoint);
                var angle = ComputeAngleInDegrees(previousPoint, connectionPoint, candidateNextPoint, 1e-6);

                if (bestIndex < 0 ||
                    (bestClosesPath && !closesPath) ||
                    (bestClosesPath == closesPath && branchDegree < bestBranchDegree) ||
                    (bestClosesPath == closesPath && branchDegree == bestBranchDegree &&
                     Math.Abs(length - bestLength) > 0.001 && length < bestLength) ||
                    (bestClosesPath == closesPath && branchDegree == bestBranchDegree &&
                     Math.Abs(length - bestLength) <= 0.001 && Math.Abs(angle - bestAngle) > 0.001 &&
                     angle > bestAngle)) {
                    bestIndex = index;
                    bestNextPoint = candidateNextPoint;
                    bestClosesPath = closesPath;
                    bestBranchDegree = branchDegree;
                    bestLength = length;
                    bestAngle = angle;
                }
            }

            return bestIndex;
        }

        private static Dictionary<string, int> BuildBoundaryNodeDegrees(List<ProjectedSegment2D> boundarySegments) {
            var result = new Dictionary<string, int>();
            if (boundarySegments == null || boundarySegments.Count == 0) return result;

            foreach (var segment in boundarySegments) {
                if (segment?.StartPoint != null) {
                    var startKey = BuildPointKey(segment.StartPoint);
                    if (!result.ContainsKey(startKey))
                        result[startKey] = 0;
                    result[startKey]++;
                }

                if (segment?.EndPoint != null) {
                    var endKey = BuildPointKey(segment.EndPoint);
                    if (!result.ContainsKey(endKey))
                        result[endKey] = 0;
                    result[endKey]++;
                }
            }

            return result;
        }

        private static int GetBoundaryNodeDegree(Dictionary<string, int> nodeDegrees, TSG.Point point) {
            if (nodeDegrees == null || point == null) return int.MaxValue;

            var key = BuildPointKey(point);
            return nodeDegrees.TryGetValue(key, out var degree) ? degree : int.MaxValue;
        }

        private static string BuildPointKey(TSG.Point point) {
            return
                Math.Round(point.X / DuplicateToleranceMillimeters) + ":" +
                Math.Round(point.Y / DuplicateToleranceMillimeters);
        }

        private static bool IsPointAlreadyUsedInPath(TSG.Point point, List<TSG.Point> path) {
            if (point == null || path == null || path.Count == 0) return false;

            foreach (var currentPoint in path)
                if (ArePointsEqual(point, currentPoint))
                    return true;

            return false;
        }

        private static bool WouldCandidateSegmentIntersectCurrentPath(
            List<TSG.Point> currentPath,
            TSG.Point segmentStartPoint,
            TSG.Point segmentEndPoint,
            bool appendToEnd
        ) {
            if (currentPath == null || currentPath.Count < 2 || segmentStartPoint == null || segmentEndPoint == null)
                return false;

            for (var index = 0; index < currentPath.Count - 1; index++) {
                var pathSegmentStartPoint = currentPath[index];
                var pathSegmentEndPoint = currentPath[index + 1];

                if (appendToEnd && index == currentPath.Count - 2)
                    continue;

                if (!appendToEnd && index == 0)
                    continue;

                if (!TryGetSegmentIntersectionPoint(
                        segmentStartPoint,
                        segmentEndPoint,
                        pathSegmentStartPoint,
                        pathSegmentEndPoint,
                        out var intersectionPoint))
                    continue;

                var touchesAtAllowedEndpoint =
                    ArePointsEqual(intersectionPoint, segmentStartPoint) ||
                    ArePointsEqual(intersectionPoint, segmentEndPoint) ||
                    ArePointsEqual(intersectionPoint, pathSegmentStartPoint) ||
                    ArePointsEqual(intersectionPoint, pathSegmentEndPoint);

                if (!touchesAtAllowedEndpoint)
                    return true;

                var sharesConnectionEndpoint = ArePointsEqual(intersectionPoint, segmentStartPoint);
                var closesPathAtOppositeEnd = appendToEnd
                    ? ArePointsEqual(intersectionPoint, currentPath[0])
                    : ArePointsEqual(intersectionPoint, currentPath[currentPath.Count - 1]);

                if (!sharesConnectionEndpoint && !closesPathAtOppositeEnd)
                    return true;
            }

            return false;
        }

        private static int GetNextUnusedSegmentIndex(bool[] usedSegments) {
            if (usedSegments == null) return -1;

            for (var index = 0; index < usedSegments.Length; index++)
                if (!usedSegments[index])
                    return index;

            return -1;
        }

        private static List<TSG.Point> PrepareBoundaryPath(List<TSG.Point> points) {
            if (points == null) return null;

            var result = RemoveSequentialDuplicatePoints(points);
            result = RemoveNearDuplicates(result, DuplicateToleranceMillimeters);
            result = SimplifyBoundaryPath(result);
            result = RemoveNearDuplicates(result, DuplicateToleranceMillimeters);

            return result;
        }

        private static List<TSG.Point> SimplifyBoundaryPath(List<TSG.Point> points) {
            if (points == null || points.Count < 3) return points;

            var result = new List<TSG.Point> {
                new TSG.Point(points[0].X, points[0].Y, 0)
            };

            for (var index = 1; index < points.Count - 1; index++) {
                var previousPoint = result[result.Count - 1];
                var currentPoint = points[index];
                var nextPoint = points[index + 1];

                if (GetDistanceToSegment(currentPoint, previousPoint, nextPoint) <= 0.1)
                    continue;

                result.Add(new TSG.Point(currentPoint.X, currentPoint.Y, 0));
            }

            result.Add(new TSG.Point(points[points.Count - 1].X, points[points.Count - 1].Y, 0));
            return RemoveSequentialDuplicatePoints(result);
        }

        private static List<TSG.Point> RemoveSequentialDuplicatePoints(IEnumerable<TSG.Point> points) {
            var result = new List<TSG.Point>();
            if (points == null) return result;

            foreach (var point in points) {
                if (point == null) continue;
                if (result.Count == 0 || !ArePointsEqual(result[result.Count - 1], point))
                    result.Add(new TSG.Point(point.X, point.Y, 0));
            }

            return result;
        }

        private static double ComputeSignedArea(IReadOnlyList<TSG.Point> points) {
            if (points == null || points.Count < 3) return 0.0;

            var area = 0.0;
            for (var index = 0; index < points.Count - 1; index++)
                area += points[index].X * points[index + 1].Y - points[index + 1].X * points[index].Y;

            return area * 0.5;
        }

        private static int ComparePointsLexicographically(TSG.Point firstPoint, TSG.Point secondPoint) {
            if (Math.Abs(firstPoint.X - secondPoint.X) > DuplicateToleranceMillimeters)
                return firstPoint.X < secondPoint.X ? -1 : 1;

            if (Math.Abs(firstPoint.Y - secondPoint.Y) > DuplicateToleranceMillimeters)
                return firstPoint.Y < secondPoint.Y ? -1 : 1;

            return 0;
        }

        private static string BuildSegmentKey(TSG.Point startPoint, TSG.Point endPoint) {
            return
                Math.Round(startPoint.X / DuplicateToleranceMillimeters) + ":" +
                Math.Round(startPoint.Y / DuplicateToleranceMillimeters) + ":" +
                Math.Round(endPoint.X / DuplicateToleranceMillimeters) + ":" +
                Math.Round(endPoint.Y / DuplicateToleranceMillimeters);
        }

        private static bool TryGetSegmentIntersectionPoint(
            TSG.Point firstStartPoint,
            TSG.Point firstEndPoint,
            TSG.Point secondStartPoint,
            TSG.Point secondEndPoint,
            out TSG.Point intersectionPoint
        ) {
            intersectionPoint = null;

            var denominator =
                (firstStartPoint.X - firstEndPoint.X) * (secondStartPoint.Y - secondEndPoint.Y) -
                (firstStartPoint.Y - firstEndPoint.Y) * (secondStartPoint.X - secondEndPoint.X);

            if (Math.Abs(denominator) <= 1e-9)
                return false;

            var t =
                ((firstStartPoint.X - secondStartPoint.X) * (secondStartPoint.Y - secondEndPoint.Y) -
                 (firstStartPoint.Y - secondStartPoint.Y) * (secondStartPoint.X - secondEndPoint.X)) / denominator;

            var u =
                ((firstStartPoint.X - secondStartPoint.X) * (firstStartPoint.Y - firstEndPoint.Y) -
                 (firstStartPoint.Y - secondStartPoint.Y) * (firstStartPoint.X - firstEndPoint.X)) / denominator;

            if (t < -1e-6 || t > 1.0 + 1e-6) return false;
            if (u < -1e-6 || u > 1.0 + 1e-6) return false;

            intersectionPoint = new TSG.Point(
                firstStartPoint.X + t * (firstEndPoint.X - firstStartPoint.X),
                firstStartPoint.Y + t * (firstEndPoint.Y - firstStartPoint.Y),
                0
            );

            return true;
        }

        private static bool AreSegmentsCollinear(
            TSG.Point firstStartPoint,
            TSG.Point firstEndPoint,
            TSG.Point secondStartPoint,
            TSG.Point secondEndPoint
        ) {
            return
                Math.Abs(CrossProduct(firstStartPoint, firstEndPoint, secondStartPoint)) <=
                DuplicateToleranceMillimeters &&
                Math.Abs(CrossProduct(firstStartPoint, firstEndPoint, secondEndPoint)) <= DuplicateToleranceMillimeters;
        }

        private static double CrossProduct(TSG.Point firstPoint, TSG.Point secondPoint, TSG.Point thirdPoint) {
            return
                (secondPoint.X - firstPoint.X) * (thirdPoint.Y - firstPoint.Y) -
                (secondPoint.Y - firstPoint.Y) * (thirdPoint.X - firstPoint.X);
        }

        private static bool IsPointOnSegment(TSG.Point point, TSG.Point segmentStartPoint, TSG.Point segmentEndPoint) {
            if (point == null || segmentStartPoint == null || segmentEndPoint == null) return false;
            if (Math.Abs(CrossProduct(segmentStartPoint, segmentEndPoint, point)) > DuplicateToleranceMillimeters)
                return false;

            var minX = Math.Min(segmentStartPoint.X, segmentEndPoint.X) - DuplicateToleranceMillimeters;
            var maxX = Math.Max(segmentStartPoint.X, segmentEndPoint.X) + DuplicateToleranceMillimeters;
            var minY = Math.Min(segmentStartPoint.Y, segmentEndPoint.Y) - DuplicateToleranceMillimeters;
            var maxY = Math.Max(segmentStartPoint.Y, segmentEndPoint.Y) + DuplicateToleranceMillimeters;

            return point.X >= minX && point.X <= maxX && point.Y >= minY && point.Y <= maxY;
        }

        private static bool IsPointInsidePolygon(TSG.Point point, IReadOnlyList<TSG.Point> polygonPoints) {
            if (point == null || polygonPoints == null || polygonPoints.Count < 3) return false;

            var isInside = false;
            var lastIndex = polygonPoints.Count - 1;

            for (var index = 0; index < polygonPoints.Count; index++) {
                var currentPoint = polygonPoints[index];
                var previousPoint = polygonPoints[lastIndex];

                if (IsPointOnSegment(point, previousPoint, currentPoint))
                    return true;

                var intersects =
                    currentPoint.Y > point.Y != previousPoint.Y > point.Y &&
                    point.X < (previousPoint.X - currentPoint.X) * (point.Y - currentPoint.Y) /
                    (previousPoint.Y - currentPoint.Y + 1e-12) + currentPoint.X;

                if (intersects)
                    isInside = !isInside;

                lastIndex = index;
            }

            return isInside;
        }

        private static double ComputePolylineLength(IReadOnlyList<TSG.Point> points) {
            if (points == null || points.Count < 2) return 0.0;

            var totalLength = 0.0;
            for (var index = 0; index < points.Count - 1; index++)
                totalLength += ComputeDistance2D(points[index], points[index + 1]);

            return totalLength;
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
                path = PrepareBoundaryPath(path);

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