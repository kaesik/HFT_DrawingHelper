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

        private static bool HasExactlyOnePart(TSD.View selectedView) {
            var modelParts = GetModelPartsFromDrawingView(selectedView);
            return modelParts != null && modelParts.Count == 1;
        }

        #endregion

        #region Main Edge Drawing Flow

        private static string DrawEdgesWithNumbers() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return null;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return null;

            var pickedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (pickedView == null) return null;

            var drawingParts = GetSelectedDrawingParts(drawingHandler);
            TSD.View commonView;

            if (drawingParts.Count == 0) {
                var singleDrawingPart = GetSingleDrawingPartFromPickedView(pickedView);
                if (singleDrawingPart == null) {
                    MessageBox.Show(
                        "Jeśli nie zaznaczasz elementu, wskazany widok musi zawierać dokładnie jeden element typu Part.");
                    return null;
                }

                drawingParts = new List<DrawingPartWithBounds> {
                    singleDrawingPart
                };

                commonView = singleDrawingPart.DrawingPart.GetView() as TSD.View;
                if (commonView == null) {
                    MessageBox.Show("Nie udało się ustalić widoku elementu.");
                    return null;
                }
            }
            else {
                commonView = GetCommonViewFromSelectedParts(drawingParts);
                if (commonView == null) {
                    MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                    return null;
                }
            }

            var outlineSnapshots = GetPartOutlineSnapshots(commonView, drawingParts);
            if (outlineSnapshots == null || outlineSnapshots.Count == 0) {
                MessageBox.Show("Nie udało się wyznaczyć krawędzi dla zaznaczonych elementów.");
                return null;
            }

            DrawOutlineSnapshots(commonView, outlineSnapshots);

            if (!HasExactlyOnePart(commonView)) {
                drawing.CommitChanges();
                return string.Empty;
            }

            var edgesByNumber = BuildEdgesByNumberFromOutlines(outlineSnapshots);
            if (edgesByNumber.Count == 0) {
                drawing.CommitChanges();
                return string.Empty;
            }

            var numberedGroups = BuildNumberedEdgeGroups(
                edgesByNumber,
                JoinToleranceMillimeters,
                NearStraightAngleDegrees
            );

            numberedGroups = FilterShortNumberedEdgeGroups(
                numberedGroups,
                MinimumNumberedEdgeLengthMillimeters
            );

            DrawEdgeNumbers(commonView, numberedGroups);

            drawing.CommitChanges();

            return FormatEdgeNumbersForTextBox(
                numberedGroups.Keys.OrderBy(number => number).ToList()
            );
        }

        private static void DrawOutlineSnapshots(
            TSD.View view,
            List<PartOutlineSnapshot> outlineSnapshots
        ) {
            if (view == null || outlineSnapshots == null || outlineSnapshots.Count == 0) return;

            foreach (var group in outlineSnapshots
                         .Where(snapshot => snapshot?.SegmentGroups != null && snapshot.SegmentGroups.Count != 0)
                         .SelectMany(snapshot =>
                             snapshot.SegmentGroups.Where(currentGroup =>
                                 currentGroup?.Points != null && currentGroup.Points.Count >= 2)))
                if (!group.IsPolyline || group.Points.Count == 2)
                    DrawStraightSegmentPrimitive(view, group.Points[0], group.Points[1], TSD.DrawingColors.Red);
                else
                    DrawPolylinePrimitive(view, group.Points, TSD.DrawingColors.Red);
        }

        #endregion

        #region Constants

        private const double DuplicateToleranceMillimeters = 0.5;
        private const double NearStraightAngleDegrees = 170.0;
        private const double JoinToleranceMillimeters = 0.5;
        private const double NumberOffsetMillimeters = 20.0;
        private const double MinimumNumberedEdgeLengthMillimeters = 100.0;
        private const int MaximumInlineEdgeCount = 10;

        #endregion

        #region Shared Outline Data

        private sealed class OutlineSegmentGroup {
            public bool IsPolyline { get; set; }
            public List<TSG.Point> Points { get; } = new List<TSG.Point>();
        }

        private sealed class PartOutlineSnapshot {
            public PartBounds Bounds { get; set; }
            public List<TSG.Point> Vertices { get; set; }
            public List<OutlineSegmentGroup> SegmentGroups { get; } = new List<OutlineSegmentGroup>();
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

        #region Shared Part Geometry Helpers

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

        private static List<DrawingPartWithBounds> GetDrawingPartsFromView(TSD.View view) {
            var result = new List<DrawingPartWithBounds>();
            if (view == null) return result;

            var iterator = view.GetAllObjects(typeof(TSD.ModelObject));
            if (iterator == null) return result;

            while (iterator.MoveNext()) {
                if (!(iterator.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                result.Add(new DrawingPartWithBounds { DrawingPart = drawingPart });
            }

            return result;
        }

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

        private static List<DrawingPartWithBounds> GetSingleDrawingPartFromView(TSD.View view) {
            var result = new List<DrawingPartWithBounds>();
            if (view == null) return result;

            var addedModelIds = new HashSet<int>();
            var iterator = view.GetAllObjects(typeof(TSD.ModelObject));
            if (iterator == null) return result;

            while (iterator.MoveNext()) {
                if (!(iterator.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                if (!(MyModel.SelectModelObject(drawingPart.ModelIdentifier) is TSM.Part modelPart)) continue;

                if (!addedModelIds.Add(modelPart.Identifier.ID)) continue;

                result.Add(new DrawingPartWithBounds { DrawingPart = drawingPart });
            }

            return result;
        }

        #endregion

        #region Drawing And Numbering

        private static void DrawEdgeNumbers(
            TSD.View view,
            Dictionary<int, NumberedEdgeGroup> numberedGroups
        ) {
            if (view == null || numberedGroups == null || numberedGroups.Count == 0) return;

            foreach (var pair in numberedGroups.OrderBy(group => group.Key)) {
                var group = pair.Value;
                if (group == null) continue;

                TSG.Point numberPoint;

                if (!group.IsPolyline) {
                    var singleSegment = group.EdgeSegments[0];
                    numberPoint = ComputeTextInsertionPointForSegment(
                        singleSegment.StartPoint,
                        singleSegment.EndPoint,
                        NumberOffsetMillimeters
                    );
                }
                else
                    numberPoint = ComputeTextInsertionPointForPolyline(
                        group.PolylinePoints,
                        NumberOffsetMillimeters
                    );

                new TSD.Text(view, numberPoint, group.GroupNumber.ToString()).Insert();
            }
        }

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
                             snapshot.SegmentGroups.Where(group => group?.Points != null && group.Points.Count >= 2))) {
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

        #endregion

        #region Nested Types

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