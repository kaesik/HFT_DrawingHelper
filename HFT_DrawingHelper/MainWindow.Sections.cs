using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSG = Tekla.Structures.Geometry3d;
using TSM = Tekla.Structures.Model;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Section View Creation

        private static void CreateSectionViewsFromEdges(
            TSD.View baseView,
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber
        ) {
            if (baseView == null) return;
            if (edgesByNumber == null || edgesByNumber.Count == 0) return;

            var (viewAttributes, markAttributes) = GetSectionAttributes();

            var baseBox = baseView.GetAxisAlignedBoundingBox();
            var cursorX = baseBox.UpperRight.X + Gap;
            var cursorY = baseBox.UpperRight.Y;

            foreach (var pair in edgesByNumber.OrderBy(item => item.Key)) {
                var edge = pair.Value;
                if (edge == null) continue;

                var edgeStart = edge.Item1;
                var edgeEnd = edge.Item2;

                var middlePoint = new TSG.Point(
                    (edgeStart.X + edgeEnd.X) * 0.5,
                    (edgeStart.Y + edgeEnd.Y) * 0.5,
                    (edgeStart.Z + edgeEnd.Z) * 0.5
                );

                var dx = edgeEnd.X - edgeStart.X;
                var dy = edgeEnd.Y - edgeStart.Y;
                var length = Math.Sqrt(dx * dx + dy * dy);
                if (length < 1e-6) continue;

                var normalX = -dy / length;
                var normalY = dx / length;

                var halfLength = _sectionHeightMillimeters * 0.5;

                var startPoint = new TSG.Point(
                    middlePoint.X - normalX * halfLength,
                    middlePoint.Y - normalY * halfLength,
                    middlePoint.Z
                );

                var endPoint = new TSG.Point(
                    middlePoint.X + normalX * halfLength,
                    middlePoint.Y + normalY * halfLength,
                    middlePoint.Z
                );

                NormalizeSectionDirection(ref startPoint, ref endPoint);

                var insertionPoint = new TSG.Point(cursorX, cursorY, baseView.Origin.Z);

                var ok = TSD.View.CreateSectionView(
                    baseView,
                    startPoint,
                    endPoint,
                    insertionPoint,
                    _sectionDepthMillimeters,
                    _sectionDepthMillimeters,
                    viewAttributes,
                    markAttributes,
                    out var sectionView,
                    out var sectionMark
                );

                if (!ok || sectionView == null || sectionMark == null) continue;

                ApplySectionViewGeometrySettings(sectionView);

                sectionView.Modify();
                sectionMark.Modify();

                var sectionBox = sectionView.GetAxisAlignedBoundingBox();

                var deltaX = cursorX - sectionBox.LowerLeft.X;
                var deltaY = cursorY - sectionBox.UpperLeft.Y;

                sectionView.Origin = new TSG.Point(
                    sectionView.Origin.X + deltaX,
                    sectionView.Origin.Y + deltaY,
                    sectionView.Origin.Z
                );

                sectionView.Modify();

                sectionBox = sectionView.GetAxisAlignedBoundingBox();
                cursorX = sectionBox.UpperRight.X + Gap;
            }
        }

        #endregion

        private static void ApplySectionViewGeometrySettings(TSD.View sectionView) {
            if (sectionView == null) return;

            ExpandSectionViewRestrictionBox(sectionView, _sectionWidthMillimeters);
        }

        private static void ExpandSectionViewRestrictionBox(TSD.View sectionView, double marginMillimeters) {
            if (sectionView == null) return;
            if (marginMillimeters <= 0) return;

            var restrictionBox = sectionView.RestrictionBox;
            if (restrictionBox == null) return;

            var minPoint = restrictionBox.MinPoint;
            var maxPoint = restrictionBox.MaxPoint;
            if (minPoint == null || maxPoint == null) return;

            sectionView.RestrictionBox.MinPoint = new TSG.Point(
                minPoint.X - marginMillimeters,
                minPoint.Y - marginMillimeters,
                minPoint.Z
            );

            sectionView.RestrictionBox.MaxPoint = new TSG.Point(
                maxPoint.X + marginMillimeters,
                maxPoint.Y + marginMillimeters,
                maxPoint.Z
            );
        }

        #region Section Orientation

        private static void NormalizeSectionDirection(ref TSG.Point startPoint, ref TSG.Point endPoint) {
            var dx = endPoint.X - startPoint.X;
            var dy = endPoint.Y - startPoint.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 1e-9) return;

            var nx = -dy / length;
            var ny = dx / length;

            if (Math.Abs(ny) >= Math.Abs(nx)) {
                if (ny < 0) (startPoint, endPoint) = (endPoint, startPoint);
            }
            else {
                if (nx > 0) (startPoint, endPoint) = (endPoint, startPoint);
            }
        }

        #endregion

        #region Edge Filtering

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> FilterEdgesOrShowMessage(
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edges,
            HashSet<int> requested
        ) {
            if (edges == null || edges.Count == 0) return null;
            if (requested == null || requested.Count == 0) return edges;

            var filtered = edges
                .Where(pair => requested.Contains(pair.Key))
                .ToDictionary(pair => pair.Key, pair => pair.Value);

            if (filtered.Count == 0) {
                MessageBox.Show("Nie znaleziono numerowanych krawędzi o podanych numerach.");
                return null;
            }

            var missing = requested
                .Where(number => !edges.ContainsKey(number))
                .OrderBy(number => number)
                .ToList();

            if (missing.Count > 0)
                MessageBox.Show("Brak numerowanych krawędzi o numerach: " + string.Join(", ", missing));

            return filtered;
        }

        #endregion

        #region Main Section Creation Flow

        private static void AddSections(string edgeNumbersInput) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var pickedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (pickedView == null) return;

            var drawingParts = FilterIgnoredSectionDrawingParts(GetSelectedDrawingParts(drawingHandler));
            TSD.View commonView;

            if (drawingParts.Count == 0) {
                var partsFromPickedView = GetDrawingPartsFromView(pickedView);
                if (partsFromPickedView.Count != 1) {
                    MessageBox.Show(
                        "Jeśli nie zaznaczasz elementu, wskazany widok musi zawierać dokładnie jeden element typu Part.");
                    return;
                }

                var singleDrawingPart = partsFromPickedView[0];

                drawingParts = new List<DrawingPartWithBounds> {
                    singleDrawingPart
                };

                commonView = singleDrawingPart.DrawingPart.GetView() as TSD.View;
                if (commonView == null) {
                    MessageBox.Show("Nie udało się ustalić widoku elementu.");
                    return;
                }
            }
            else {
                commonView = GetCommonViewFromSelectedParts(drawingParts);
                if (commonView == null) {
                    MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                    return;
                }
            }

            if (GetDrawingPartsFromView(commonView).Count != 1) {
                MessageBox.Show(
                    "Przekroje można tworzyć tylko dla widoku zawierającego dokładnie jeden element typu Part.");
                return;
            }

            var outlineSnapshots = GetPartOutlineSnapshots(commonView, drawingParts);
            if (outlineSnapshots == null || outlineSnapshots.Count == 0) {
                MessageBox.Show("Nie udało się wyznaczyć krawędzi dla zaznaczonych elementów.");
                return;
            }

            var edgesByNumber = BuildEdgesByNumberFromOutlines(outlineSnapshots);
            if (edgesByNumber == null || edgesByNumber.Count == 0) {
                MessageBox.Show("Nie znaleziono krawędzi do utworzenia przekrojów.");
                return;
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

            if (numberedGroups.Count == 0) {
                MessageBox.Show("Nie znaleziono numerowanych krawędzi do utworzenia przekrojów.");
                return;
            }

            var requestedGroupNumbers = ParseGroupNumbers(edgeNumbersInput);

            var sectionEdgesByGroupNumber = numberedGroups
                .Where(pair => pair.Value?.SectionEdge != null)
                .ToDictionary(pair => pair.Key, pair => pair.Value.SectionEdge);

            var filteredSectionEdges = FilterEdgesOrShowMessage(sectionEdgesByGroupNumber, requestedGroupNumbers);
            if (filteredSectionEdges == null || filteredSectionEdges.Count == 0) return;

            CreateSectionViewsFromEdges(commonView, filteredSectionEdges);
            drawing.CommitChanges();
        }

        private static void AutoAddSectionsAtWeldMarks() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) {
                MessageBox.Show("Brak połączenia z Tekla Structures.");
                return;
            }

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) {
                MessageBox.Show("Brak otwartego rysunku.");
                return;
            }

            var selectedDrawingParts = FilterIgnoredSectionDrawingParts(GetSelectedDrawingParts(drawingHandler));
            if (selectedDrawingParts.Count != 1) {
                MessageBox.Show("Zaznacz dokładnie jeden element w widoku, w którym mają powstać przekroje MS/WS.");
                return;
            }

            var selectedDrawingPart = selectedDrawingParts[0];
            var commonView = selectedDrawingPart.DrawingPart.GetView() as TSD.View;
            if (commonView == null) {
                MessageBox.Show("Nie udało się ustalić widoku zaznaczonego elementu.");
                return;
            }

            var allDrawingPartsInView = GetAllDrawingPartsFromViewIncludingIgnored(commonView);
            var weldDrawingParts = allDrawingPartsInView
                .Where(part => part?.DrawingPart != null)
                .Where(part => !AreSameDrawingPart(part.DrawingPart, selectedDrawingPart.DrawingPart))
                .Where(part => IsWeldMarkSectionDrawingPart(part.DrawingPart))
                .ToList();

            if (weldDrawingParts.Count == 0) {
                MessageBox.Show(
                    "W widoku zaznaczonego elementu nie znaleziono elementów zaczynających się od MS lub WS.");
                return;
            }

            var sectionEdgesByGroupNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            var acceptedSectionEdges = new List<Tuple<TSG.Point, TSG.Point>>();
            var sectionIndex = 1;

            foreach (var sectionEdge in from weldDrawingPart in weldDrawingParts
                     let weldOutlineSnapshots = GetPartOutlineSnapshots(
                         commonView,
                         new List<DrawingPartWithBounds> { weldDrawingPart }
                     )
                     select GetSectionEdgeFromWeldPart(commonView, weldDrawingPart, weldOutlineSnapshots)
                     into sectionEdge
                     where sectionEdge != null
                     where !IsSectionCutTooCloseToAcceptedCuts(
                         sectionEdge,
                         acceptedSectionEdges,
                         MinimumSectionCutDistanceMillimeters
                     )
                     select sectionEdge) {
                acceptedSectionEdges.Add(sectionEdge);
                sectionEdgesByGroupNumber[sectionIndex] = sectionEdge;
                sectionIndex++;
            }

            if (sectionEdgesByGroupNumber.Count == 0) {
                MessageBox.Show("Nie udało się utworzyć przekroju żadnego elementu MS/WS.");
                return;
            }

            CreateSectionViewsFromEdges(commonView, sectionEdgesByGroupNumber);
            drawing.CommitChanges();

            var message = "Automatycznie dodano przekroje MS/WS: " + sectionEdgesByGroupNumber.Count;

            MessageBox.Show(message);
        }

        private static bool IsSectionCutTooCloseToAcceptedCuts(
            Tuple<TSG.Point, TSG.Point> sectionEdge,
            List<Tuple<TSG.Point, TSG.Point>> acceptedSectionEdges,
            double minimumDistanceMillimeters
        ) {
            if (sectionEdge == null) return true;
            if (acceptedSectionEdges == null || acceptedSectionEdges.Count == 0) return false;
            return !(minimumDistanceMillimeters <= 0) &&
                   (from acceptedSectionEdge in acceptedSectionEdges
                       where acceptedSectionEdge != null
                       select GetSectionCutDistanceMillimeters(sectionEdge, acceptedSectionEdge))
                   .Any(cutDistance => cutDistance < minimumDistanceMillimeters);
        }

        private static double GetSectionCutDistanceMillimeters(
            Tuple<TSG.Point, TSG.Point> firstSectionEdge,
            Tuple<TSG.Point, TSG.Point> secondSectionEdge
        ) {
            if (firstSectionEdge == null || secondSectionEdge == null)
                return double.MaxValue;

            var firstCenter = GetSectionEdgeCenterPoint(firstSectionEdge);
            var secondCenter = GetSectionEdgeCenterPoint(secondSectionEdge);
            if (firstCenter == null || secondCenter == null)
                return double.MaxValue;

            var centerDistance = ComputeDistance2D(firstCenter, secondCenter);

            var firstDirection = GetSectionEdgeDirection(firstSectionEdge);
            var secondDirection = GetSectionEdgeDirection(secondSectionEdge);
            if (firstDirection == null || secondDirection == null)
                return centerDistance;

            var directionDot = Math.Abs(firstDirection.Item1 * secondDirection.Item1 +
                                        firstDirection.Item2 * secondDirection.Item2);
            var directionsAreAlmostParallel = directionDot >= Math.Cos(30.0 * Math.PI / 180.0);

            if (!directionsAreAlmostParallel)
                return centerDistance;

            var centerDeltaX = secondCenter.X - firstCenter.X;
            var centerDeltaY = secondCenter.Y - firstCenter.Y;

            var firstStationDistance =
                Math.Abs(centerDeltaX * firstDirection.Item1 + centerDeltaY * firstDirection.Item2);
            var secondStationDistance =
                Math.Abs(centerDeltaX * secondDirection.Item1 + centerDeltaY * secondDirection.Item2);

            return Math.Min(centerDistance, Math.Min(firstStationDistance, secondStationDistance));
        }

        private static TSG.Point GetSectionEdgeCenterPoint(Tuple<TSG.Point, TSG.Point> sectionEdge) {
            if (sectionEdge?.Item1 == null || sectionEdge.Item2 == null)
                return null;

            return new TSG.Point(
                (sectionEdge.Item1.X + sectionEdge.Item2.X) * 0.5,
                (sectionEdge.Item1.Y + sectionEdge.Item2.Y) * 0.5,
                (sectionEdge.Item1.Z + sectionEdge.Item2.Z) * 0.5
            );
        }

        private static Tuple<double, double> GetSectionEdgeDirection(Tuple<TSG.Point, TSG.Point> sectionEdge) {
            if (sectionEdge?.Item1 == null || sectionEdge.Item2 == null)
                return null;

            var dx = sectionEdge.Item2.X - sectionEdge.Item1.X;
            var dy = sectionEdge.Item2.Y - sectionEdge.Item1.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            return length < 1e-6 ? null : Tuple.Create(dx / length, dy / length);
        }

        private static Tuple<TSG.Point, TSG.Point> GetSectionEdgeFromWeldPart(
            TSD.View view,
            DrawingPartWithBounds weldDrawingPart,
            List<PartOutlineSnapshot> weldOutlineSnapshots
        ) {
            var outlineEdge = GetLongestOutlineEdgeFromSnapshots(weldOutlineSnapshots);
            if (outlineEdge != null)
                return outlineEdge;

            var modelEdge = GetLongestDrawingEdgeFromModelPart(view, weldDrawingPart?.DrawingPart);
            return modelEdge ?? GetBoundingBoxSectionEdgeFromSnapshots(weldOutlineSnapshots);
        }

        private static Tuple<TSG.Point, TSG.Point> GetLongestOutlineEdgeFromSnapshots(
            List<PartOutlineSnapshot> snapshots
        ) {
            if (snapshots == null || snapshots.Count == 0)
                return null;

            Tuple<TSG.Point, TSG.Point> bestEdge = null;
            var bestLength = 0.0;

            foreach (var snapshot in snapshots.Where(snapshot => snapshot != null)) {
                if (snapshot.SegmentGroups != null)
                    foreach (var segmentGroup in snapshot.SegmentGroups.Where(segmentGroup =>
                                 segmentGroup?.Points != null && segmentGroup.Points.Count >= 2)) {
                        for (var index = 0; index < segmentGroup.Points.Count - 1; index++)
                            TryUseLongerEdge(segmentGroup.Points[index], segmentGroup.Points[index + 1], ref bestEdge,
                                ref bestLength);

                        if (segmentGroup.IsPolyline && segmentGroup.Points.Count > 2)
                            TryUseLongerEdge(segmentGroup.Points[segmentGroup.Points.Count - 1], segmentGroup.Points[0],
                                ref bestEdge, ref bestLength);
                    }

                if (snapshot.Vertices != null && snapshot.Vertices.Count >= 2) {
                    for (var index = 0; index < snapshot.Vertices.Count - 1; index++)
                        TryUseLongerEdge(snapshot.Vertices[index], snapshot.Vertices[index + 1], ref bestEdge,
                            ref bestLength);

                    if (snapshot.Vertices.Count > 2)
                        TryUseLongerEdge(snapshot.Vertices[snapshot.Vertices.Count - 1], snapshot.Vertices[0],
                            ref bestEdge, ref bestLength);
                }
            }

            return bestLength >= 1.0 ? bestEdge : null;
        }

        private static Tuple<TSG.Point, TSG.Point> GetBoundingBoxSectionEdgeFromSnapshots(
            List<PartOutlineSnapshot> snapshots
        ) {
            if (snapshots == null || snapshots.Count == 0)
                return null;

            var bounds = snapshots
                .Where(snapshot => snapshot?.Bounds != null)
                .Select(snapshot => snapshot.Bounds)
                .ToList();

            if (bounds.Count == 0)
                return null;

            return GetLongestAxisEdgeFromBounds(
                bounds.Min(bound => bound.MinX),
                bounds.Max(bound => bound.MaxX),
                bounds.Min(bound => bound.MinY),
                bounds.Max(bound => bound.MaxY)
            );
        }

        private static Tuple<TSG.Point, TSG.Point> GetLongestDrawingEdgeFromModelPart(
            TSD.View view,
            TSD.Part drawingPart
        ) {
            if (view == null || drawingPart == null)
                return null;

            TSM.Part modelPart;
            try {
                modelPart = MyModel.SelectModelObject(drawingPart.ModelIdentifier) as TSM.Part;
            }
            catch {
                modelPart = null;
            }

            if (modelPart == null)
                return null;

            try {
                var solid = modelPart.GetSolid();
                if (solid == null)
                    return null;

                var minimumPoint = solid.MinimumPoint;
                var maximumPoint = solid.MaximumPoint;

                var modelCorners = new List<TSG.Point> {
                    new TSG.Point(minimumPoint.X, minimumPoint.Y, minimumPoint.Z),
                    new TSG.Point(minimumPoint.X, minimumPoint.Y, maximumPoint.Z),
                    new TSG.Point(minimumPoint.X, maximumPoint.Y, minimumPoint.Z),
                    new TSG.Point(minimumPoint.X, maximumPoint.Y, maximumPoint.Z),
                    new TSG.Point(maximumPoint.X, minimumPoint.Y, minimumPoint.Z),
                    new TSG.Point(maximumPoint.X, minimumPoint.Y, maximumPoint.Z),
                    new TSG.Point(maximumPoint.X, maximumPoint.Y, minimumPoint.Z),
                    new TSG.Point(maximumPoint.X, maximumPoint.Y, maximumPoint.Z)
                };

                var transformationMatrix = TSG.MatrixFactory.ToCoordinateSystem(view.DisplayCoordinateSystem);
                var drawingPoints = modelCorners
                    .Select(point => transformationMatrix.Transform(point))
                    .Select(point => new TSG.Point(point.X, point.Y, 0))
                    .ToList();

                return GetLongestAxisEdgeFromBounds(
                    drawingPoints.Min(point => point.X),
                    drawingPoints.Max(point => point.X),
                    drawingPoints.Min(point => point.Y),
                    drawingPoints.Max(point => point.Y)
                );
            }
            catch {
                return null;
            }
        }

        private static Tuple<TSG.Point, TSG.Point> GetLongestAxisEdgeFromBounds(
            double minX,
            double maxX,
            double minY,
            double maxY
        ) {
            var width = Math.Abs(maxX - minX);
            var height = Math.Abs(maxY - minY);

            if (width < 1.0 && height < 1.0)
                return null;

            if (width >= height) {
                var y = (minY + maxY) * 0.5;
                return Tuple.Create(
                    new TSG.Point(minX, y, 0),
                    new TSG.Point(maxX, y, 0)
                );
            }

            var x = (minX + maxX) * 0.5;
            return Tuple.Create(
                new TSG.Point(x, minY, 0),
                new TSG.Point(x, maxY, 0)
            );
        }

        private static void TryUseLongerEdge(
            TSG.Point startPoint,
            TSG.Point endPoint,
            ref Tuple<TSG.Point, TSG.Point> bestEdge,
            ref double bestLength
        ) {
            if (startPoint == null || endPoint == null)
                return;

            var length = ComputeDistance2D(startPoint, endPoint);
            if (length <= bestLength)
                return;

            bestLength = length;
            bestEdge = Tuple.Create(
                new TSG.Point(startPoint.X, startPoint.Y, 0),
                new TSG.Point(endPoint.X, endPoint.Y, 0)
            );
        }

        private static List<DrawingPartWithBounds> GetAllDrawingPartsFromViewIncludingIgnored(TSD.View view) {
            var result = new List<DrawingPartWithBounds>();
            if (view == null) return result;

            var drawingObjects = view.GetAllObjects(typeof(TSD.Part));
            if (drawingObjects == null) return result;

            while (drawingObjects.MoveNext()) {
                if (!(drawingObjects.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                result.Add(new DrawingPartWithBounds { DrawingPart = drawingPart });
            }

            return result;
        }

        private static bool IsWeldMarkSectionDrawingPart(TSD.Part drawingPart) {
            if (drawingPart == null) return false;

            try {
                if (!(MyModel.SelectModelObject(drawingPart.ModelIdentifier) is TSM.Part modelPart))
                    return false;

                return IsWeldMarkSectionPartName(modelPart.Name);
            }
            catch {
                return false;
            }
        }

        private static bool IsWeldMarkSectionPartName(string name) {
            if (string.IsNullOrWhiteSpace(name)) return false;

            return name.StartsWith("MS", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("WS", StringComparison.OrdinalIgnoreCase);
        }

        private static bool AreSameDrawingPart(TSD.Part firstPart, TSD.Part secondPart) {
            if (firstPart == null || secondPart == null) return false;

            try {
                if (firstPart.ModelIdentifier == null || secondPart.ModelIdentifier == null)
                    return false;

                return firstPart.ModelIdentifier.ID == secondPart.ModelIdentifier.ID;
            }
            catch {
                return ReferenceEquals(firstPart, secondPart);
            }
        }

        #endregion

        #region Section Attributes

        private static (TSD.View.ViewAttributes, TSD.SectionMarkBase.SectionMarkAttributes) GetSectionAttributes() {
            var view = new TSD.View.ViewAttributes(
                GetSectionAttributeNameOrDefault(_viewAttributeName, DefaultViewAttributeName));
            var mark = new TSD.SectionMarkBase.SectionMarkAttributes();
            mark.LoadAttributes(GetSectionAttributeNameOrDefault(_markAttributeName, DefaultMarkAttributeName));

            return (view, mark);
        }

        private static void UpdateSectionAttributeNames(string viewAttributeName, string markAttributeName) {
            _viewAttributeName = GetSectionAttributeNameOrDefault(viewAttributeName, DefaultViewAttributeName);
            _markAttributeName = GetSectionAttributeNameOrDefault(markAttributeName, DefaultMarkAttributeName);
        }

        private static string GetSectionAttributeNameOrDefault(string value, string defaultValue) {
            return string.IsNullOrWhiteSpace(value) ? defaultValue : value.Trim();
        }

        #endregion

        #region Constants

        private const string DefaultViewAttributeName = "standard";
        private const string DefaultMarkAttributeName = "standard";
        private static string _viewAttributeName = DefaultViewAttributeName;
        private static string _markAttributeName = DefaultMarkAttributeName;
        private const double DefaultSectionDepthMillimeters = 1.0;
        private const double DefaultSectionHeightMillimeters = 300.0;
        private const double DefaultSectionWidthMillimeters = 100.0;
        private const double Gap = 10.0;
        private const double MinimumSectionCutDistanceMillimeters = 10.0;
        private static double _sectionDepthMillimeters = DefaultSectionDepthMillimeters;
        private static double _sectionHeightMillimeters = DefaultSectionHeightMillimeters;
        private static double _sectionWidthMillimeters = DefaultSectionWidthMillimeters;

        private static void UpdateSectionGeometrySettings(
            double sectionDepthMillimeters,
            double sectionHeightMillimeters,
            double sectionWidthMillimeters
        ) {
            _sectionDepthMillimeters = sectionDepthMillimeters;
            _sectionHeightMillimeters = sectionHeightMillimeters;
            _sectionWidthMillimeters = sectionWidthMillimeters;
        }

        private static void ResetSectionGeometrySettingsToDefault() {
            UpdateSectionGeometrySettings(
                DefaultSectionDepthMillimeters,
                DefaultSectionHeightMillimeters,
                DefaultSectionWidthMillimeters
            );
        }

        private static double GetSectionDepthMillimeters() {
            return _sectionDepthMillimeters;
        }

        private static double GetSectionHeightMillimeters() {
            return _sectionHeightMillimeters;
        }

        private static double GetSectionWidthMillimeters() {
            return _sectionWidthMillimeters;
        }

        private static double GetDefaultSectionDepthMillimeters() {
            return DefaultSectionDepthMillimeters;
        }

        private static double GetDefaultSectionHeightMillimeters() {
            return DefaultSectionHeightMillimeters;
        }

        private static double GetDefaultSectionWidthMillimeters() {
            return DefaultSectionWidthMillimeters;
        }

        #endregion
    }
}