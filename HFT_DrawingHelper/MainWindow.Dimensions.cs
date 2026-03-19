using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Constants

        private const double OverallDimensionOffsetMillimeters = 20.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;
        private const double CurvedDimensionArcDepthRatio = 0.15;

        #endregion

        #region Data Structures

        private enum DimensionType {
            Straight,
            Curved
        }

        private enum HorizontalDimensionPlacement {
            Below,
            Above
        }

        private enum VerticalDimensionPlacement {
            Left,
            Right
        }

        private sealed class DimensionOptions {
            public DimensionType DimensionType { get; set; }
            public HorizontalDimensionPlacement HorizontalPlacement { get; set; }
            public VerticalDimensionPlacement VerticalPlacement { get; set; }
            public bool CreateHorizontal { get; set; }
            public bool CreateVertical { get; set; }
        }

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

        #region Main Dimension Flow

        private static void AddDimensions(DimensionOptions options) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var activeDrawing = drawingHandler.GetActiveDrawing();
            if (activeDrawing == null) return;

            var selectedParts = GetSelectedDrawingParts(drawingHandler);

            if (selectedParts == null || selectedParts.Count == 0) {
                AddDimensionsFromSelectedView(drawingHandler, activeDrawing, options);
                return;
            }

            AddDimensionsFromSelectedParts(activeDrawing, selectedParts, options);
        }

        private static void AddDimensionsFromSelectedView(
            TSD.DrawingHandler drawingHandler,
            TSD.Drawing activeDrawing,
            DimensionOptions options
        ) {
            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var allViewParts = GetDrawingPartsFromView(selectedView);
            var allPartsBounds = GetAssemblyBounds(allViewParts);
            if (allPartsBounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obwiedni elementów w widoku.");
                return;
            }

            CreateMaximumAssemblyDimensions(selectedView, allPartsBounds, options);
            activeDrawing.CommitChanges();
        }

        private static void AddDimensionsFromSelectedParts(
            TSD.Drawing activeDrawing,
            List<DrawingPartWithBounds> selectedParts,
            DimensionOptions options
        ) {
            var selectedView = GetCommonViewFromSelectedParts(selectedParts);
            if (selectedView == null) {
                MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                return;
            }

            DrawPartFramesAndNames(selectedParts);

            var allViewParts = GetDrawingPartsFromView(selectedView);
            var allPartsBounds = GetAssemblyBounds(allViewParts);
            if (allPartsBounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obwiedni wszystkich elementów w widoku.");
                return;
            }

            CreateOverallAssemblyDimensions(selectedView, allPartsBounds, selectedParts, options);
            activeDrawing.CommitChanges();
        }

        #endregion

        #region Dimension Creation

        private static void CreateMaximumAssemblyDimensions(
            TSD.View selectedView,
            PartBounds allPartsBounds,
            DimensionOptions options
        ) {
            if (selectedView == null || allPartsBounds == null) return;

            var totalWidth = allPartsBounds.MaxX - allPartsBounds.MinX;
            var totalHeight = allPartsBounds.MaxY - allPartsBounds.MinY;

            if (options.CreateHorizontal && totalWidth >= MinimumDimensionSpanMillimeters) {
                var anchorY = options.HorizontalPlacement == HorizontalDimensionPlacement.Above
                    ? allPartsBounds.MaxY
                    : allPartsBounds.MinY;

                var horizontalPoints = new TSD.PointList {
                    new TSG.Point(allPartsBounds.MinX, anchorY, 0),
                    new TSG.Point(allPartsBounds.MaxX, anchorY, 0)
                };

                CreateDimensionSet(
                    selectedView,
                    horizontalPoints,
                    GetHorizontalDirectionVector(options.HorizontalPlacement),
                    OverallDimensionOffsetMillimeters,
                    options.DimensionType
                );
            }

            if (options.CreateVertical && totalHeight >= MinimumDimensionSpanMillimeters) {
                var anchorX = options.VerticalPlacement == VerticalDimensionPlacement.Right
                    ? allPartsBounds.MaxX
                    : allPartsBounds.MinX;

                var verticalPoints = new TSD.PointList {
                    new TSG.Point(anchorX, allPartsBounds.MinY, 0),
                    new TSG.Point(anchorX, allPartsBounds.MaxY, 0)
                };

                CreateDimensionSet(
                    selectedView,
                    verticalPoints,
                    GetVerticalDirectionVector(options.VerticalPlacement),
                    OverallDimensionOffsetMillimeters,
                    options.DimensionType
                );
            }
        }

        private static void CreateOverallAssemblyDimensions(
            TSD.View selectedView,
            PartBounds allPartsBounds,
            List<DrawingPartWithBounds> selectedParts,
            DimensionOptions options
        ) {
            if (selectedView == null || allPartsBounds == null) return;

            var selectedBounds = GetAssemblyBounds(selectedParts);
            if (selectedBounds == null) return;

            var xCoordinates = selectedParts
                .Where(partWithBounds => partWithBounds?.Bounds != null)
                .SelectMany(partWithBounds => new[] { partWithBounds.Bounds.MinX, partWithBounds.Bounds.MaxX })
                .Append(allPartsBounds.MinX)
                .Append(allPartsBounds.MaxX)
                .GroupBy(xCoordinate => Math.Round(xCoordinate, 3))
                .Select(coordinateGroup => coordinateGroup.First())
                .OrderBy(xCoordinate => xCoordinate)
                .ToList();

            var yCoordinates = selectedParts
                .Where(partWithBounds => partWithBounds?.Bounds != null)
                .SelectMany(partWithBounds => new[] { partWithBounds.Bounds.MinY, partWithBounds.Bounds.MaxY })
                .Append(allPartsBounds.MinY)
                .Append(allPartsBounds.MaxY)
                .GroupBy(yCoordinate => Math.Round(yCoordinate, 3))
                .Select(coordinateGroup => coordinateGroup.First())
                .OrderBy(yCoordinate => yCoordinate)
                .ToList();

            var totalWidth = allPartsBounds.MaxX - allPartsBounds.MinX;
            var totalHeight = allPartsBounds.MaxY - allPartsBounds.MinY;

            if (options.CreateHorizontal && totalWidth >= MinimumDimensionSpanMillimeters && xCoordinates.Count >= 2) {
                var anchorY = options.HorizontalPlacement == HorizontalDimensionPlacement.Above
                    ? selectedBounds.MaxY
                    : selectedBounds.MinY;

                var horizontalPoints = new TSD.PointList();
                foreach (var xCoordinate in xCoordinates)
                    horizontalPoints.Add(new TSG.Point(xCoordinate, anchorY, 0));

                CreateDimensionSet(
                    selectedView,
                    horizontalPoints,
                    GetHorizontalDirectionVector(options.HorizontalPlacement),
                    OverallDimensionOffsetMillimeters,
                    options.DimensionType
                );
            }

            if (options.CreateVertical && totalHeight >= MinimumDimensionSpanMillimeters && yCoordinates.Count >= 2) {
                var anchorX = options.VerticalPlacement == VerticalDimensionPlacement.Right
                    ? selectedBounds.MaxX
                    : selectedBounds.MinX;

                var verticalPoints = new TSD.PointList();
                foreach (var yCoordinate in yCoordinates)
                    verticalPoints.Add(new TSG.Point(anchorX, yCoordinate, 0));

                CreateDimensionSet(
                    selectedView,
                    verticalPoints,
                    GetVerticalDirectionVector(options.VerticalPlacement),
                    OverallDimensionOffsetMillimeters,
                    options.DimensionType
                );
            }
        }

        private static TSG.Vector GetHorizontalDirectionVector(HorizontalDimensionPlacement placement) {
            return placement == HorizontalDimensionPlacement.Above
                ? new TSG.Vector(0.0, 1.0, 0.0)
                : new TSG.Vector(0.0, -1.0, 0.0);
        }

        private static TSG.Vector GetVerticalDirectionVector(VerticalDimensionPlacement placement) {
            return placement == VerticalDimensionPlacement.Right
                ? new TSG.Vector(1.0, 0.0, 0.0)
                : new TSG.Vector(-1.0, 0.0, 0.0);
        }

        private static void CreateDimensionSet(
            TSD.View selectedView,
            TSD.PointList dimensionPoints,
            TSG.Vector directionVector,
            double offsetMillimeters,
            DimensionType dimensionType
        ) {
            if (dimensionType == DimensionType.Straight) {
                var straightHandler = new TSD.StraightDimensionSetHandler();
                straightHandler.CreateDimensionSet(
                    selectedView,
                    dimensionPoints,
                    directionVector,
                    offsetMillimeters
                );
                return;
            }

            var arcPoints = ComputeArcPoints(dimensionPoints, directionVector);
            var curvedHandler = new TSD.CurvedDimensionSetHandler();
            curvedHandler.CreateCurvedDimensionSetOrthogonal(
                selectedView,
                arcPoints.Item1,
                arcPoints.Item2,
                arcPoints.Item3,
                dimensionPoints,
                offsetMillimeters
            );
        }

        private static Tuple<TSG.Point, TSG.Point, TSG.Point> ComputeArcPoints(
            TSD.PointList dimensionPoints,
            TSG.Vector directionVector
        ) {
            var points = new List<TSG.Point>();
            foreach (var item in dimensionPoints)
                if (item is TSG.Point point)
                    points.Add(point);

            var firstPoint = points.First();
            var lastPoint = points.Last();

            var midX = (firstPoint.X + lastPoint.X) / 2.0;
            var midY = (firstPoint.Y + lastPoint.Y) / 2.0;

            var spanX = lastPoint.X - firstPoint.X;
            var spanY = lastPoint.Y - firstPoint.Y;
            var span = Math.Sqrt(spanX * spanX + spanY * spanY);
            var arcDepth = span * CurvedDimensionArcDepthRatio;

            return Tuple.Create(
                new TSG.Point(firstPoint.X, firstPoint.Y, 0),
                new TSG.Point(midX + directionVector.X * arcDepth, midY + directionVector.Y * arcDepth, 0),
                new TSG.Point(lastPoint.X, lastPoint.Y, 0)
            );
        }

        #endregion

        #region Part Selection And View Validation

        private static List<DrawingPartWithBounds> GetSelectedDrawingParts(TSD.DrawingHandler drawingHandler) {
            if (drawingHandler == null) return new List<DrawingPartWithBounds>();

            var objectSelector = drawingHandler.GetDrawingObjectSelector();
            if (objectSelector == null) return new List<DrawingPartWithBounds>();

            var selectedObjects = objectSelector.GetSelected();
            if (selectedObjects == null) return new List<DrawingPartWithBounds>();

            selectedObjects.SelectInstances = true;
            return BuildPartList(selectedObjects, false);
        }

        private static TSD.View GetCommonViewFromSelectedParts(List<DrawingPartWithBounds> drawingParts) {
            if (drawingParts == null || drawingParts.Count == 0) return null;

            var firstDrawingPart = drawingParts[0].DrawingPart;

            if (!(firstDrawingPart?.GetView() is TSD.View commonView)) return null;

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

                var originsMatch =
                    Math.Abs(firstCoordinateSystem.Origin.X - secondCoordinateSystem.Origin.X) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.Origin.Y - secondCoordinateSystem.Origin.Y) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.Origin.Z - secondCoordinateSystem.Origin.Z) < 0.001;

                var axisXMatch =
                    Math.Abs(firstCoordinateSystem.AxisX.X - secondCoordinateSystem.AxisX.X) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisX.Y - secondCoordinateSystem.AxisX.Y) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisX.Z - secondCoordinateSystem.AxisX.Z) < 0.001;

                var axisYMatch =
                    Math.Abs(firstCoordinateSystem.AxisY.X - secondCoordinateSystem.AxisY.X) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisY.Y - secondCoordinateSystem.AxisY.Y) < 0.001 &&
                    Math.Abs(firstCoordinateSystem.AxisY.Z - secondCoordinateSystem.AxisY.Z) < 0.001;

                return originsMatch && axisXMatch && axisYMatch;
            }
            catch {
                return false;
            }
        }

        #endregion

        #region Part Collection

        private static List<DrawingPartWithBounds> GetDrawingPartsFromView(TSD.View selectedView) {
            if (selectedView == null) return new List<DrawingPartWithBounds>();

            var allObjects = selectedView.GetAllObjects(typeof(TSD.Part));
            if (allObjects == null) return new List<DrawingPartWithBounds>();

            allObjects.SelectInstances = true;
            return BuildPartList(allObjects, true);
        }

        private static List<DrawingPartWithBounds> BuildPartList(
            TSD.DrawingObjectEnumerator source,
            bool checkDepth
        ) {
            var result = new List<DrawingPartWithBounds>();
            var addedModelIdentifiers = new HashSet<int>();

            while (source.MoveNext()) {
                if (!(source.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;

                var modelIdentifier = drawingPart.ModelIdentifier.ID;
                if (!addedModelIdentifiers.Add(modelIdentifier)) continue;

                if (checkDepth && !DoesPartIntersectViewDepth(drawingPart)) continue;

                var partBounds = GetBoundsFromDrawingPart(drawingPart);
                if (partBounds == null) continue;

                result.Add(new DrawingPartWithBounds {
                    DrawingPart = drawingPart,
                    Bounds = partBounds
                });
            }

            return result;
        }

        #endregion

        #region Bounds Calculation

        private static PartBounds GetAssemblyBounds(List<DrawingPartWithBounds> drawingParts) {
            if (drawingParts == null || drawingParts.Count == 0) return null;

            var minimumX = double.MaxValue;
            var minimumY = double.MaxValue;
            var maximumX = double.MinValue;
            var maximumY = double.MinValue;
            var foundAny = false;

            foreach (var partWithBounds in drawingParts.Where(partWithBounds => partWithBounds?.Bounds != null)) {
                if (partWithBounds.Bounds.MinX < minimumX) minimumX = partWithBounds.Bounds.MinX;
                if (partWithBounds.Bounds.MinY < minimumY) minimumY = partWithBounds.Bounds.MinY;
                if (partWithBounds.Bounds.MaxX > maximumX) maximumX = partWithBounds.Bounds.MaxX;
                if (partWithBounds.Bounds.MaxY > maximumY) maximumY = partWithBounds.Bounds.MaxY;
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

        private static PartBounds GetBoundsFromDrawingPart(TSD.Part drawingPart) {
            if (!(drawingPart?.GetView() is TSD.View drawingView)) return null;

            var modelObject = MyModel.SelectModelObject(drawingPart.ModelIdentifier);
            if (!(modelObject is TSM.Part modelPart)) return null;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(drawingView.DisplayCoordinateSystem)
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
                    if (!TryGetEdgePoints(edgeEnumerator.Current, out var startPoint, out var endPoint)) continue;

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

        private static bool DoesPartIntersectViewDepth(TSD.Part drawingPart) {
            if (!(drawingPart?.GetView() is TSD.View drawingView)) return false;
            if (drawingView.RestrictionBox == null) return true;

            var modelObject = MyModel.SelectModelObject(drawingPart.ModelIdentifier);
            if (!(modelObject is TSM.Part modelPart)) return false;

            var workPlaneHandler = MyModel.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(drawingView.DisplayCoordinateSystem)
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

                var viewMinimumZ = Math.Min(drawingView.RestrictionBox.MinPoint.Z,
                    drawingView.RestrictionBox.MaxPoint.Z);
                var viewMaximumZ = Math.Max(drawingView.RestrictionBox.MinPoint.Z,
                    drawingView.RestrictionBox.MaxPoint.Z);
                const double tolerance = 1.0;

                while (edgeEnumerator.MoveNext()) {
                    if (!TryGetEdgePoints(edgeEnumerator.Current, out var startPoint, out var endPoint)) continue;
                    if (startPoint == null || endPoint == null) continue;

                    if (DoesSegmentIntersectDepthRange(startPoint.Z, endPoint.Z, viewMinimumZ, viewMaximumZ, tolerance))
                        return true;
                }

                return false;
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static bool DoesSegmentIntersectDepthRange(
            double segmentStartZ,
            double segmentEndZ,
            double rangeMinimumZ,
            double rangeMaximumZ,
            double tolerance
        ) {
            var segmentMinimumZ = Math.Min(segmentStartZ, segmentEndZ);
            var segmentMaximumZ = Math.Max(segmentStartZ, segmentEndZ);

            return segmentMaximumZ >= rangeMinimumZ - tolerance &&
                   segmentMinimumZ <= rangeMaximumZ + tolerance;
        }

        #endregion

        #region Part Frame Drawing

        private static void DrawPartFramesAndNames(List<DrawingPartWithBounds> drawingParts) {
            if (drawingParts == null || drawingParts.Count == 0) return;

            foreach (var partWithBounds in drawingParts.Where(partWithBounds =>
                         partWithBounds?.DrawingPart != null && partWithBounds.Bounds != null)) {
                if (!(partWithBounds.DrawingPart.GetView() is TSD.View drawingView)) continue;

                DrawRectangleForBounds(drawingView, partWithBounds.Bounds);

                var centerX = (partWithBounds.Bounds.MinX + partWithBounds.Bounds.MaxX) / 2.0;
                var centerY = (partWithBounds.Bounds.MinY + partWithBounds.Bounds.MaxY) / 2.0;

                new TSD.Text(drawingView, new TSG.Point(centerX, centerY, 0), GetPartLabel(partWithBounds.DrawingPart))
                    .Insert();
            }
        }

        private static void DrawRectangleForBounds(TSD.View drawingView, PartBounds bounds) {
            if (drawingView == null || bounds == null) return;

            var cornerPoints = new TSD.PointList {
                new TSG.Point(bounds.MinX, bounds.MinY, 0),
                new TSG.Point(bounds.MinX, bounds.MaxY, 0),
                new TSG.Point(bounds.MaxX, bounds.MaxY, 0),
                new TSG.Point(bounds.MaxX, bounds.MinY, 0)
            };

            var boundingPolygon = new TSD.Polygon(drawingView, cornerPoints);
            boundingPolygon.Attributes.Line.Color = TSD.DrawingColors.Red;
            boundingPolygon.Insert();
        }

        private static string GetPartLabel(TSD.Part drawingPart) {
            if (drawingPart == null) return "Unknown";

            try {
                var modelObject = MyModel.SelectModelObject(drawingPart.ModelIdentifier);
                if (modelObject is TSM.Part modelPart && !string.IsNullOrWhiteSpace(modelPart.Name))
                    return modelPart.Name;

                return modelObject?.GetType().Name ?? drawingPart.GetType().Name;
            }
            catch {
                return drawingPart.GetType().Name;
            }
        }

        #endregion

        #region Geometry Helpers

        private static bool TryGetEdgePoints(
            object edge,
            out TSG.Point startPoint,
            out TSG.Point endPoint
        ) {
            startPoint = null;
            endPoint = null;

            if (edge == null) return false;

            try {
                dynamic dynamicEdge = edge;

                var candidateStartPoint = dynamicEdge.StartPoint as TSG.Point;
                var candidateEndPoint = dynamicEdge.EndPoint as TSG.Point;

                if (candidateStartPoint != null && candidateEndPoint != null) {
                    startPoint = candidateStartPoint;
                    endPoint = candidateEndPoint;
                    return true;
                }

                candidateStartPoint = dynamicEdge.Point1 as TSG.Point;
                candidateEndPoint = dynamicEdge.Point2 as TSG.Point;

                if (candidateStartPoint != null && candidateEndPoint != null) {
                    startPoint = candidateStartPoint;
                    endPoint = candidateEndPoint;
                    return true;
                }
            }
            catch {
                // ignored
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
    }
}