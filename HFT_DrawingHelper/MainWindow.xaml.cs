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

            var edges = GetContourPlateEdges(selectedView, true);
            if (edges != null && edges.Count > 0) {
                DrawEdges(selectedView, edges);
                drawing.CommitChanges();
                return;
            }

            edges = GetLoftedPlateEdges(selectedView, true);
            if (edges != null && edges.Count > 0) {
                DrawEdges(selectedView, edges);
                drawing.CommitChanges();
                return;
            }

            MessageBox.Show("Nie znaleziono ani ContourPlate ani LoftedPlate na tym widoku.");
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

            var requestedGroupNumbers = ParseEdgeNumbers(edgeNumbersInput);

            const double joinToleranceMillimeters = 0.5;
            const double nearStraightAngleDegrees = 170.0;

            var contourEdgesBySegmentNumber = GetContourPlateEdges(selectedView, false);
            if (contourEdgesBySegmentNumber != null && contourEdgesBySegmentNumber.Count > 0) {
                var contourGroupsByGroupNumber = BuildNumberedEdgeGroups(
                    contourEdgesBySegmentNumber,
                    joinToleranceMillimeters,
                    nearStraightAngleDegrees
                );

                var contourSectionEdgesByGroupNumber = contourGroupsByGroupNumber
                    .Where(pair => pair.Value != null && pair.Value.SectionEdge != null)
                    .ToDictionary(pair => pair.Key, pair => pair.Value.SectionEdge);

                var filteredSectionEdges =
                    FilterEdgesOrShowMessage(contourSectionEdgesByGroupNumber, requestedGroupNumbers);
                if (filteredSectionEdges == null) return;

                CreateSectionViewsFromEdges(selectedView, filteredSectionEdges);
                drawing.CommitChanges();
                return;
            }

            var loftedEdgesBySegmentNumber = GetLoftedPlateEdges(selectedView, false);
            if (loftedEdgesBySegmentNumber != null && loftedEdgesBySegmentNumber.Count > 0) {
                var loftedGroupsByGroupNumber = BuildNumberedEdgeGroups(
                    loftedEdgesBySegmentNumber,
                    joinToleranceMillimeters,
                    nearStraightAngleDegrees
                );

                var loftedSectionEdgesByGroupNumber = loftedGroupsByGroupNumber
                    .Where(pair => pair.Value?.SectionEdge != null)
                    .ToDictionary(pair => pair.Key, pair => pair.Value.SectionEdge);

                var filteredSectionEdges =
                    FilterEdgesOrShowMessage(loftedSectionEdgesByGroupNumber, requestedGroupNumbers);
                if (filteredSectionEdges == null) return;

                CreateSectionViewsFromEdges(selectedView, filteredSectionEdges);
                drawing.CommitChanges();
            }
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

        #region Arrange

        private static void ArrangeDrawingObject(TSD.Drawing drawing) {
            var containerView = drawing.GetSheet();
            var viewEnumerator = containerView.GetViews();

            if (viewEnumerator != null) {
                viewEnumerator.SelectInstances = false;

                while (viewEnumerator.MoveNext())
                    if (viewEnumerator.Current is TSD.View view) {
                        var drawingMarkEnumerator = view.GetAllObjects(typeof(TSD.MarkBase));

                        while (drawingMarkEnumerator.MoveNext())
                            if (drawingMarkEnumerator.Current is TSD.MarkBase mark) {
                                var newPosition = new TSG.Point(
                                    mark.InsertionPoint.X + 10,
                                    mark.InsertionPoint.Y + 10,
                                    mark.InsertionPoint.Z
                                );

                                mark.InsertionPoint = newPosition;
                                mark.Modify();
                            }
                    }
            }
        }

        private static void ArrangeDrawingDim(TSD.Drawing drawing) {
            var containerView = drawing.GetSheet();
            var viewEnumerator = containerView.GetViews();

            if (viewEnumerator != null)
                while (viewEnumerator.MoveNext())
                    if (viewEnumerator.Current is TSD.View view) {
                        var dimensionEnumerator = view.GetAllObjects(typeof(TSD.StraightDimensionSet));

                        while (dimensionEnumerator.MoveNext())
                            if (dimensionEnumerator.Current is TSD.StraightDimensionSet dimensionSet)
                                dimensionSet.Modify();
                    }
        }

        private static void ArrangeDrawingView(TSD.Drawing drawing) {
            var containerView = drawing.GetSheet();
            var viewEnumerator = containerView.GetViews();

            if (viewEnumerator != null)
                while (viewEnumerator.MoveNext()) {
                    var view = viewEnumerator.Current as TSD.View;

                    if (view != null) {
                        var newOrigin = new TSG.Point(view.Origin.X + 50, view.Origin.Y + 50, view.Origin.Z);
                        view.Origin = newOrigin;
                    }

                    view?.Modify();
                }
        }

        #endregion

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
            bool showMessages
        ) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;

            const double minLengthMm = 100.0;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var modelParts = GetModelPartsFromDrawingView(selectedView, showMessages);

                var contourPlate = FindFirstContourPlate(modelParts);
                if (contourPlate == null) {
                    if (showMessages) MessageBox.Show("Nie znaleziono ContourPlate na tym widoku.");
                    return edgesByNumber;
                }

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(contourPlate.GetCoordinateSystem()));

                var edgesInPlatePlane = GetContourPlateEdgesInPlane(contourPlate);

                if (edgesInPlatePlane == null || edgesInPlatePlane.Count == 0) {
                    if (showMessages)
                        MessageBox.Show("Nie znaleziono krawędzi w płaszczyźnie ContourPlate.");
                    return edgesByNumber;
                }

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
            bool showMessages
        ) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;

            const double minLengthMm = 100.0;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var modelParts = GetModelPartsFromDrawingView(selectedView, showMessages);

                var loftedPlate = FindFirstLoftedPlate(modelParts);
                if (loftedPlate == null) {
                    if (showMessages) MessageBox.Show("Nie znaleziono LoftedPlate na tym widoku.");
                    return edgesByNumber;
                }

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(loftedPlate.GetCoordinateSystem()));

                var edgesInPlatePlane = GetLoftedPlateEdgesInPlane(loftedPlate);

                if (edgesInPlatePlane == null || edgesInPlatePlane.Count == 0) {
                    if (showMessages)
                        MessageBox.Show("Nie znaleziono krawędzi w płaszczyźnie LoftedPlate.");
                    return edgesByNumber;
                }

                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem));

                var edgesInViewPlane = TransformEdgesBetweenCoordinateSystems(
                    loftedPlate.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    edgesInPlatePlane
                );

                var number = 0;

                for (var i = 0; i < edgesInViewPlane.Count; i++) {
                    var start = edgesInViewPlane[i].A;
                    var end = edgesInViewPlane[i].B;

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

        private static List<(TSG.Point A, TSG.Point B)> GetLoftedPlateEdgesInPlane(TSM.LoftedPlate loftedPlate) {
            var edges = new List<(TSG.Point A, TSG.Point B)>();
            if (loftedPlate == null) return edges;

            var solid = loftedPlate.GetSolid();
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
        private static List<TSM.Part> GetModelPartsFromDrawingView(TSD.View drawingView, bool showMessages) {
            var modelParts = new List<TSM.Part>();
            var addedModelIdentifiers = new HashSet<string>();
            var partTypeCounters = new Dictionary<string, int>();

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

                    var partTypeName = modelPart.GetType().Name;
                    if (partTypeCounters.ContainsKey(partTypeName))
                        partTypeCounters[partTypeName] += 1;
                    else
                        partTypeCounters.Add(partTypeName, 1);
                }
                catch {
                    // ignored
                }
            }

            if (!showMessages) return modelParts;

            switch (modelParts.Count) {
                case 0:
                    MessageBox.Show("Znaleziono: 0 elementów typu Part na tym widoku.");
                    return modelParts;
                case 1:
                    MessageBox.Show("Znaleziono: 1 element typu " + modelParts[0].GetType().Name + ".");
                    return modelParts;
            }

            var summaryLines = partTypeCounters
                .OrderByDescending(pair => pair.Value)
                .Select(pair => pair.Key + " (" + pair.Value + ")")
                .ToList();

            MessageBox.Show("Znaleziono: " + modelParts.Count + " elementów typów:\n" +
                            string.Join("\n", summaryLines));

            return modelParts;
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