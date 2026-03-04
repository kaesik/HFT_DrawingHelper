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

        #region Add sections

        private static void AddSections() {
            var drawingHandler = new TSD.DrawingHandler();

            if (drawingHandler.GetConnectionStatus()) {
                var drawing = drawingHandler.GetActiveDrawing();
                if (drawing == null) return;

                DrawEdgeInContourPlate();
            }
        }

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            AddSections();
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

        private static List<(TSG.Point A, TSG.Point B)>
            GetContourPlateEdgesInPlatePlane(TSM.ContourPlate contourPlate) {
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

        #region Other

        private static void DrawEdgeInContourPlate() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var selector = drawingHandler.GetDrawingObjectSelector();
            var selected = selector.GetSelected();

            if (selected == null) {
                MessageBox.Show("Nie zaznaczono żadnego obiektu.");
                return;
            }

            TSD.View selectedView = null;
            selected.SelectInstances = false;

            while (selected.MoveNext()) {
                selectedView = selected.Current as TSD.View;
                if (selectedView != null) break;
            }

            if (selectedView == null) {
                MessageBox.Show("Zaznacz widok na rysunku, a potem uruchom funkcję.");
                return;
            }

            var model = new TSM.Model();

            var savedPlane = GetCurrentPlane();

            try {
                var modelParts = GetModelPartsFromDrawingView(selectedView);

                var contourPlate = FindFirstContourPlate(modelParts);
                if (contourPlate == null) {
                    MessageBox.Show("Nie znaleziono ContourPlate na tym widoku.");
                    return;
                }

                SetCurrentPlane(new TSM.TransformationPlane(contourPlate.GetCoordinateSystem()));
                var edgesInPlatePlane = GetContourPlateEdgesInPlatePlane(contourPlate);

                if (edgesInPlatePlane.Count == 0) {
                    MessageBox.Show("Nie znaleziono krawędzi (intersection points = 0) nawet w płaszczyźnie blachy.");
                    return;
                }

                SetCurrentPlane(new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem));

                var edgesInViewPlane = TransformEdgesBetweenCoordinateSystems(
                    contourPlate.GetCoordinateSystem(),
                    selectedView.DisplayCoordinateSystem,
                    edgesInPlatePlane
                );

                var rotated = RotateEdgesAroundCenter2D(edgesInViewPlane, 90.0);

                var shiftedDown = TranslateEdgesDownByHalfHeight(rotated);

                DrawEdgesOnDrawingAsRedLines(selectedView, shiftedDown);
                drawing.CommitChanges();
            }
            finally {
                SetCurrentPlane(savedPlane);
            }

            return;

            #region local functions

            void SetCurrentPlane(TSM.TransformationPlane plane) {
                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(plane);
            }

            TSM.TransformationPlane GetCurrentPlane() {
                return model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
            }

            TSM.ContourPlate FindFirstContourPlate(List<TSM.Part> parts) {
                if (parts == null || parts.Count == 0) return null;
                foreach (var p in parts)
                    if (p is TSM.ContourPlate cp)
                        return cp;
                return null;
            }

            TSG.Point ComputeEdgesCenter2D(List<(TSG.Point A, TSG.Point B)> edges) {
                var pts = new List<TSG.Point>(edges.Count * 2);
                foreach (var e in edges) {
                    pts.Add(e.A);
                    pts.Add(e.B);
                }

                var minX = pts.Min(p => p.X);
                var maxX = pts.Max(p => p.X);
                var minY = pts.Min(p => p.Y);
                var maxY = pts.Max(p => p.Y);

                return new TSG.Point((minX + maxX) * 0.5, (minY + maxY) * 0.5, 0);
            }

            TSG.Point RotatePointAroundCenter2D(TSG.Point p, TSG.Point c, double cos, double sin) {
                var dx = p.X - c.X;
                var dy = p.Y - c.Y;

                var rx = c.X + dx * cos - dy * sin;
                var ry = c.Y + dx * sin + dy * cos;

                return new TSG.Point(rx, ry, 0);
            }

            List<(TSG.Point A, TSG.Point B)> TransformEdgesBetweenCoordinateSystems(
                TSG.CoordinateSystem fromCs,
                TSG.CoordinateSystem toCs,
                List<(TSG.Point A, TSG.Point B)> edges
            ) {
                if (edges == null || edges.Count == 0) return new List<(TSG.Point A, TSG.Point B)>();

                var fromToGlobal = TSG.MatrixFactory.FromCoordinateSystem(fromCs);
                var globalToTo = TSG.MatrixFactory.ToCoordinateSystem(toCs);

                var result = new List<(TSG.Point A, TSG.Point B)>(edges.Count);

                foreach (var e in edges) {
                    var aGlobal = fromToGlobal.Transform(e.A);
                    var bGlobal = fromToGlobal.Transform(e.B);

                    var aTo = globalToTo.Transform(aGlobal);
                    var bTo = globalToTo.Transform(bGlobal);

                    aTo.Z = 0;
                    bTo.Z = 0;

                    result.Add((new TSG.Point(aTo.X, aTo.Y, 0), new TSG.Point(bTo.X, bTo.Y, 0)));
                }

                return result;
            }

            List<(TSG.Point A, TSG.Point B)> RotateEdgesAroundCenter2D(
                List<(TSG.Point A, TSG.Point B)> edges,
                double angleDegrees
            ) {
                if (edges == null || edges.Count == 0) return new List<(TSG.Point A, TSG.Point B)>();

                var beforePts = new List<TSG.Point>(edges.Count * 2);
                foreach (var e in edges) {
                    beforePts.Add(e.A);
                    beforePts.Add(e.B);
                }

                var beforeMinX = beforePts.Min(p => p.X);
                var beforeMinY = beforePts.Min(p => p.Y);

                var center = ComputeEdgesCenter2D(edges);

                var rad = angleDegrees * Math.PI / 180.0;
                var cos = Math.Cos(rad);
                var sin = Math.Sin(rad);

                var rotated = new List<(TSG.Point A, TSG.Point B)>(edges.Count);
                foreach (var e in edges)
                    rotated.Add((RotatePointAroundCenter2D(e.A, center, cos, sin),
                        RotatePointAroundCenter2D(e.B, center, cos, sin)));

                var afterPts = new List<TSG.Point>(rotated.Count * 2);
                foreach (var e in rotated) {
                    afterPts.Add(e.A);
                    afterPts.Add(e.B);
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

            void DrawEdgesOnDrawingAsRedLines(TSD.View view, List<(TSG.Point A, TSG.Point B)> edges) {
                if (view == null || edges == null || edges.Count == 0) return;

                var attrs = new TSD.Line.LineAttributes {
                    Line = new TSD.LineTypeAttributes(TSD.LineTypes.SolidLine, TSD.DrawingColors.Red)
                };

                foreach (var e in edges)
                    new TSD.Line(view, e.A, e.B, attrs).Insert();
            }

            List<(TSG.Point A, TSG.Point B)> TranslateEdgesDownByHalfHeight(List<(TSG.Point A, TSG.Point B)> edges) {
                if (edges == null || edges.Count == 0) return edges;

                var pts = new List<TSG.Point>(edges.Count * 2);
                foreach (var e in edges) {
                    pts.Add(e.A);
                    pts.Add(e.B);
                }

                var minY = pts.Min(p => p.Y);
                var maxY = pts.Max(p => p.Y);

                var halfHeight = (maxY - minY) * 0.5;

                var result = new List<(TSG.Point A, TSG.Point B)>(edges.Count);
                foreach (var e in edges)
                    result.Add((
                        new TSG.Point(e.A.X, e.A.Y - halfHeight, 0),
                        new TSG.Point(e.B.X, e.B.Y - halfHeight, 0)
                    ));

                return result;
            }

            #endregion
        }

        private static List<TSG.Point> RemoveNearDuplicates(List<TSG.Point> points, double epsilon) {
            if (points == null || points.Count == 0) return new List<TSG.Point>();

            var result = new List<TSG.Point>();
            foreach (var p in points) {
                var exists = result.Any(r => Math.Abs(r.X - p.X) <= epsilon && Math.Abs(r.Y - p.Y) <= epsilon);
                if (!exists) result.Add(p);
            }

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