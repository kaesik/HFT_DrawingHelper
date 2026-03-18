using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Main Dimension Flow

        private static void AddDimensions() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var drawingParts = GetDrawingPartsFromSelectedView(selectedView);
            if (drawingParts == null || drawingParts.Count == 0) {
                MessageBox.Show("Nie znaleziono elementów typu Part na zaznaczonym widoku.");
                return;
            }

            DrawPartFramesAndNames(drawingParts);

            var assemblyBounds = GetAssemblyBounds(drawingParts);
            if (assemblyBounds != null)
                CreateOverallAssemblyDimensions(selectedView, assemblyBounds);

            drawing.CommitChanges();
        }

        #endregion

        #region Overall Dimensions

        private static void CreateOverallAssemblyDimensions(
            TSD.View selectedView,
            PartBounds assemblyBounds
        ) {
            if (selectedView == null || assemblyBounds == null) return;

            var handler = new TSD.StraightDimensionSetHandler();

            var width = assemblyBounds.MaxX - assemblyBounds.MinX;
            var height = assemblyBounds.MaxY - assemblyBounds.MinY;

            if (width >= MinimumDimensionSpanMillimeters) {
                var horizontalPoints = new TSD.PointList {
                    new TSG.Point(assemblyBounds.MinX, assemblyBounds.MinY, 0),
                    new TSG.Point(assemblyBounds.MaxX, assemblyBounds.MinY, 0)
                };

                handler.CreateDimensionSet(
                    selectedView,
                    horizontalPoints,
                    new TSG.Vector(0.0, -1.0, 0.0),
                    OverallDimensionOffsetMillimeters
                );
            }

            if (height >= MinimumDimensionSpanMillimeters) {
                var verticalPoints = new TSD.PointList {
                    new TSG.Point(assemblyBounds.MinX, assemblyBounds.MinY, 0),
                    new TSG.Point(assemblyBounds.MinX, assemblyBounds.MaxY, 0)
                };

                handler.CreateDimensionSet(
                    selectedView,
                    verticalPoints,
                    new TSG.Vector(-1.0, 0.0, 0.0),
                    OverallDimensionOffsetMillimeters
                );
            }
        }

        #endregion

        #region Data Structures

        private sealed class DrawingPartWithBounds {
            public TSD.Part DrawingPart { get; set; }
            public PartBounds Bounds { get; set; }
        }

        private sealed class PartBounds {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
        }

        #endregion

        #region Drawing Parts

        private static List<DrawingPartWithBounds> GetDrawingPartsFromSelectedView(TSD.View selectedView) {
            var result = new List<DrawingPartWithBounds>();
            if (selectedView == null) return result;

            var allObjects = selectedView.GetAllObjects(typeof(TSD.Part));
            if (allObjects == null) return result;

            allObjects.SelectInstances = true;

            var addedModelIdentifiers = new HashSet<int>();

            while (allObjects.MoveNext()) {
                if (!(allObjects.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                var modelIdentifier = drawingPart.ModelIdentifier.ID;
                if (!addedModelIdentifiers.Add(modelIdentifier)) continue;

                if (!DoesPartIntersectViewDepth(drawingPart)) continue;

                var bounds = GetBoundsFromDrawingPart(drawingPart);
                if (bounds == null) continue;

                result.Add(new DrawingPartWithBounds {
                    DrawingPart = drawingPart,
                    Bounds = bounds
                });
            }

            return result;
        }

        private static bool DoesPartIntersectViewDepth(TSD.Part drawingPart) {
            if (drawingPart == null) return false;

            if (!(drawingPart.GetView() is TSD.View view)) return false;
            if (view.RestrictionBox == null) return true;

            var model = new TSM.Model();
            var modelObject = model.SelectModelObject(drawingPart.ModelIdentifier);
            if (!(modelObject is TSM.Part modelPart)) return false;

            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(view.DisplayCoordinateSystem)
                );

                TSM.Solid solid;
                try {
                    solid = modelPart.GetSolid();
                }
                catch {
                    return false;
                }

                var edgeEnumerator = solid?.GetEdgeEnumerator();
                if (edgeEnumerator == null) return false;

                var viewMinZ = Math.Min(view.RestrictionBox.MinPoint.Z, view.RestrictionBox.MaxPoint.Z);
                var viewMaxZ = Math.Max(view.RestrictionBox.MinPoint.Z, view.RestrictionBox.MaxPoint.Z);

                const double tolerance = 1.0;

                while (edgeEnumerator.MoveNext()) {
                    dynamic edge = edgeEnumerator.Current;
                    if (edge == null) continue;

                    if (!TryGetEdgePoints(edge, out TSG.Point startPoint, out TSG.Point endPoint)) continue;
                    if (startPoint == null || endPoint == null) continue;

                    if (DoesSegmentIntersectDepthRange(startPoint.Z, endPoint.Z, viewMinZ, viewMaxZ, tolerance))
                        return true;
                }

                return false;
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static bool DoesSegmentIntersectDepthRange(
            double startZ,
            double endZ,
            double rangeMinZ,
            double rangeMaxZ,
            double tolerance
        ) {
            var segmentMinZ = Math.Min(startZ, endZ);
            var segmentMaxZ = Math.Max(startZ, endZ);

            return segmentMaxZ >= rangeMinZ - tolerance &&
                   segmentMinZ <= rangeMaxZ + tolerance;
        }

        private static PartBounds GetBoundsFromDrawingPart(TSD.Part drawingPart) {
            if (drawingPart == null) return null;

            var model = new TSM.Model();
            if (!(drawingPart.GetView() is TSD.View view)) return null;

            var modelObject = model.SelectModelObject(drawingPart.ModelIdentifier);
            if (!(modelObject is TSM.Part modelPart)) return null;

            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(view.DisplayCoordinateSystem)
                );

                TSM.Solid solid;

                try {
                    solid = modelPart.GetSolid();
                }
                catch {
                    return null;
                }

                var edgeEnumerator = solid?.GetEdgeEnumerator();
                if (edgeEnumerator == null) return null;

                var minimumX = double.MaxValue;
                var minimumY = double.MaxValue;
                var maximumX = double.MinValue;
                var maximumY = double.MinValue;
                var foundAnyPoint = false;

                while (edgeEnumerator.MoveNext()) {
                    dynamic edge = edgeEnumerator.Current;
                    if (edge == null) continue;

                    if (!TryGetEdgePoints(edge, out TSG.Point startPoint, out TSG.Point endPoint)) continue;

                    UpdateBounds(startPoint, ref minimumX, ref minimumY, ref maximumX, ref maximumY);
                    UpdateBounds(endPoint, ref minimumX, ref minimumY, ref maximumX, ref maximumY);
                    foundAnyPoint = true;
                }

                if (!foundAnyPoint) return null;

                return new PartBounds {
                    MinX = minimumX,
                    MaxX = maximumX,
                    MinY = minimumY,
                    MaxY = maximumY
                };
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static PartBounds GetAssemblyBounds(List<DrawingPartWithBounds> drawingParts) {
            if (drawingParts == null || drawingParts.Count == 0) return null;

            var minimumX = double.MaxValue;
            var minimumY = double.MaxValue;
            var maximumX = double.MinValue;
            var maximumY = double.MinValue;
            var foundAny = false;

            foreach (var item in drawingParts.Where(item => item?.Bounds != null)) {
                if (item.Bounds.MinX < minimumX) minimumX = item.Bounds.MinX;
                if (item.Bounds.MinY < minimumY) minimumY = item.Bounds.MinY;
                if (item.Bounds.MaxX > maximumX) maximumX = item.Bounds.MaxX;
                if (item.Bounds.MaxY > maximumY) maximumY = item.Bounds.MaxY;
                foundAny = true;
            }

            if (!foundAny) return null;

            return new PartBounds {
                MinX = minimumX,
                MaxX = maximumX,
                MinY = minimumY,
                MaxY = maximumY
            };
        }

        #endregion

        #region Drawing

        private static void DrawPartFramesAndNames(List<DrawingPartWithBounds> drawingParts) {
            if (drawingParts == null || drawingParts.Count == 0) return;

            foreach (var item in drawingParts.Where(item => item?.DrawingPart != null && item.Bounds != null)) {
                if (!(item.DrawingPart.GetView() is TSD.View view)) continue;

                DrawRectangleForBounds(view, item.Bounds);

                var centerX = (item.Bounds.MinX + item.Bounds.MaxX) / 2.0;
                var centerY = (item.Bounds.MinY + item.Bounds.MaxY) / 2.0;
                var label = GetDrawingPartLabel(item.DrawingPart);

                new TSD.Text(view, new TSG.Point(centerX, centerY, 0), label).Insert();
            }
        }

        private static void DrawRectangleForBounds(
            TSD.View selectedView,
            PartBounds bounds
        ) {
            if (selectedView == null || bounds == null) return;

            var points = new TSD.PointList {
                new TSG.Point(bounds.MinX, bounds.MinY, 0),
                new TSG.Point(bounds.MinX, bounds.MaxY, 0),
                new TSG.Point(bounds.MaxX, bounds.MaxY, 0),
                new TSG.Point(bounds.MaxX, bounds.MinY, 0)
            };

            var polygon = new TSD.Polygon(selectedView, points);
            polygon.Attributes.Line.Color = TSD.DrawingColors.Red;
            polygon.Insert();
        }

        private static string GetDrawingPartLabel(TSD.Part drawingPart) {
            if (drawingPart == null) return "Unknown";

            var model = new TSM.Model();

            try {
                var modelObject = model.SelectModelObject(drawingPart.ModelIdentifier);
                if (modelObject is TSM.Part modelPart)
                    if (!string.IsNullOrWhiteSpace(modelPart.Name))
                        return modelPart.Name;

                return modelObject?.GetType().Name ?? drawingPart.GetType().Name;
            }
            catch {
                return drawingPart.GetType().Name;
            }
        }

        #endregion

        #region Geometry Extraction

        private static bool TryGetEdgePoints(
            dynamic edge,
            out TSG.Point startPoint,
            out TSG.Point endPoint
        ) {
            startPoint = null;
            endPoint = null;

            try {
                startPoint = edge.StartPoint as TSG.Point;
                endPoint = edge.EndPoint as TSG.Point;
                if (startPoint != null && endPoint != null) return true;
            }
            catch {
            }

            try {
                startPoint = edge.Point1 as TSG.Point;
                endPoint = edge.Point2 as TSG.Point;
                if (startPoint != null && endPoint != null) return true;
            }
            catch {
            }

            return false;
        }

        private static void UpdateBounds(
            TSG.Point point,
            ref double minimumX,
            ref double minimumY,
            ref double maximumX,
            ref double maximumY
        ) {
            if (point == null) return;

            if (point.X < minimumX) minimumX = point.X;
            if (point.Y < minimumY) minimumY = point.Y;
            if (point.X > maximumX) maximumX = point.X;
            if (point.Y > maximumY) maximumY = point.Y;
        }

        #endregion

        #region Constants

        private const double OverallDimensionOffsetMillimeters = 20.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;

        #endregion
    }
}