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

            var edges = GetContourPlateEdges(selectedView);
            if (edges == null || edges.Count == 0) return;

            DrawContourPlateEdges(selectedView, edges);

            drawing.CommitChanges();
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

            var edges = GetContourPlateEdges(selectedView);
            if (edges == null || edges.Count == 0) return;

            var requested = ParseEdgeNumbers(edgeNumbersInput);

            if (requested.Count > 0) {
                var filtered = edges
                    .Where(kvp => requested.Contains(kvp.Key))
                    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                if (filtered.Count == 0) {
                    MessageBox.Show("Nie znaleziono krawędzi o podanych numerach.");
                    return;
                }

                var missing = requested.Where(n => !edges.ContainsKey(n)).OrderBy(n => n).ToList();
                if (missing.Count > 0) MessageBox.Show("Brak krawędzi o numerach: " + string.Join(", ", missing));

                edges = filtered;
            }

            CreateSectionViewsFromEdges(selectedView, edges);

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

            const double depthUp = 100.0;
            const double depthDown = 100.0;
            const double sectionLineLengthMm = 300.0;
            const double gap = 10.0;

            // bounding box głównego widoku
            var baseBox = baseView.GetAxisAlignedBoundingBox();

            // start odkładania przekrojów – prawa krawędź głównego widoku
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

                var half = sectionLineLengthMm * 0.5;

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

                // bounding box nowego przekroju
                var secBox = sectionView.GetAxisAlignedBoundingBox();

                // przesunięcie tak, aby lewa krawędź była dokładnie przy cursorX
                var deltaX = cursorX - secBox.LowerLeft.X;
                var deltaY = cursorY - secBox.UpperLeft.Y;

                sectionView.Origin = new TSG.Point(
                    sectionView.Origin.X + deltaX,
                    sectionView.Origin.Y + deltaY,
                    sectionView.Origin.Z
                );

                sectionView.Modify();

                // nowy cursor dla kolejnego przekroju
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

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> GetContourPlateEdges(TSD.View selectedView) {
            var edgesByNumber = new Dictionary<int, Tuple<TSG.Point, TSG.Point>>();
            if (selectedView == null) return edgesByNumber;

            const double minLengthMm = 100.0;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                var modelParts = GetModelPartsFromDrawingView(selectedView);

                var contourPlate = FindFirstContourPlate(modelParts);
                if (contourPlate == null) {
                    MessageBox.Show("Nie znaleziono ContourPlate na tym widoku.");
                    return edgesByNumber;
                }

                // 1) Pobierz krawędzie w płaszczyźnie blachy
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(contourPlate.GetCoordinateSystem()));
                var edgesInPlatePlane = GetContourPlateEdgesInPlane(contourPlate);

                if (edgesInPlatePlane == null || edgesInPlatePlane.Count == 0) {
                    MessageBox.Show("Nie znaleziono krawędzi (intersection points = 0) nawet w płaszczyźnie blachy.");
                    return edgesByNumber;
                }

                // 2) Transform do płaszczyzny widoku
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem));

                var edgesInViewPlane = TransformEdgesBetweenCoordinateSystems(
                    contourPlate.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    edgesInPlatePlane
                );

                // 3) Obrót + przesunięcie (jak u Ciebie było)
                var rotated = RotateEdgesAroundCenter2D(edgesInViewPlane, 90.0);
                var shiftedDown = TranslateEdgesDownByHalfHeight(rotated);

                // 4) Filtr długości + numeracja do dictionary
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

        private static List<TSM.Part> GetModelPartsFromDrawingView(TSD.View drawingView) {
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

            switch (modelParts.Count) {
                case 0:
                    MessageBox.Show("Znaleziono: 0 elementów typu Part na tym widoku.");
                    return modelParts;
                case 1: {
                    var singlePartTypeName = modelParts[0].GetType().Name;
                    MessageBox.Show("Znaleziono: 1 element typu " + singlePartTypeName + ".");
                    return modelParts;
                }
            }

            var summaryLines = partTypeCounters
                .OrderByDescending(pair => pair.Value)
                .Select(pair => pair.Key + " (" + pair.Value + ")")
                .ToList();

            MessageBox.Show("Znaleziono: " + modelParts.Count + " elementów typów:\n" +
                            string.Join("\n", summaryLines));
            return modelParts;
        }

        private static (TSD.View.ViewAttributes, TSD.SectionMarkBase.SectionMarkAttributes) GetSectionAttributes() {
            var view = new TSD.View.ViewAttributes(ViewAttributeName);
            var mark = new TSD.SectionMarkBase.SectionMarkAttributes();
            mark.LoadAttributes(MarkAttributeName);

            return (view, mark);
        }

        #endregion

        #region Other

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

        private static void DrawContourPlateEdges(
            TSD.View view,
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber
        ) {
            if (view == null) return;
            if (edgesByNumber == null || edgesByNumber.Count == 0) return;

            const double textOffsetMm = 10.0;

            var lineAttrs = new TSD.Line.LineAttributes {
                Line = new TSD.LineTypeAttributes(TSD.LineTypes.SolidLine, TSD.DrawingColors.Red)
            };

            foreach (var kvp in edgesByNumber.OrderBy(k => k.Key)) {
                var number = kvp.Key;
                var start = kvp.Value.Item1;
                var end = kvp.Value.Item2;

                // 1) Linia
                new TSD.Line(view, start, end, lineAttrs).Insert();

                // 2) Środek
                var mid = new TSG.Point(
                    (start.X + end.X) * 0.5,
                    (start.Y + end.Y) * 0.5,
                    0
                );

                // 3) Prostopadły offset w XY (żeby numer był obok linii)
                var dx = end.X - start.X;
                var dy = end.Y - start.Y;

                var nX = -dy;
                var nY = dx;
                var nLen = Math.Sqrt(nX * nX + nY * nY);

                if (nLen > 1e-9) {
                    nX /= nLen;
                    nY /= nLen;

                    mid.X += nX * textOffsetMm;
                    mid.Y += nY * textOffsetMm;
                }
                else
                    mid.Y += textOffsetMm;

                // 4) Tekst z numerem
                var text = new TSD.Text(view, mid, number.ToString());
                text.Insert();
            }
        }

        private static TSM.ContourPlate FindFirstContourPlate(List<TSM.Part> parts) {
            if (parts == null || parts.Count == 0) return null;

            foreach (var part in parts)
                if (part is TSM.ContourPlate contourPlate)
                    return contourPlate;

            return null;
        }

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