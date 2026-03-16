using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Main Dimension Flow

        private static void AddDimensions(string assemblyInput) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var modelPartsInView = GetModelPartsFromDrawingView(selectedView);
            if (modelPartsInView == null || modelPartsInView.Count == 0) {
                MessageBox.Show("Nie znaleziono elementów modelu na zaznaczonym widoku.");
                return;
            }

            var targetAssembly = ResolveAssemblyToDimension(drawingHandler, modelPartsInView, assemblyInput);
            if (targetAssembly == null) {
                MessageBox.Show("Nie udało się ustalić assembly do wymiarowania.");
                return;
            }

            var assemblyParts = GetAssemblyPartsVisibleInView(targetAssembly, modelPartsInView);
            if (assemblyParts == null || assemblyParts.Count == 0) {
                MessageBox.Show("Nie znaleziono części wybranego assembly na zaznaczonym widoku.");
                return;
            }

            var bounds = GetBoundsFromParts(selectedView, assemblyParts);
            if (bounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obrysu całego elementu.");
                return;
            }

            CreateAssemblyDimensionSets(selectedView, bounds);
            DrawPartFramesAndNames(selectedView, assemblyParts, bounds);
            drawing.CommitChanges();
        }

        #endregion

        #region Dimension Creation

        private static void CreateAssemblyDimensionSets(
            TSD.View selectedView,
            AssemblyViewBounds bounds
        ) {
            var handler = new TSD.StraightDimensionSetHandler();

            var assemblyWidth = bounds.MaxX - bounds.MinX;
            var assemblyHeight = bounds.MaxY - bounds.MinY;

            var halfAssemblyWidth = assemblyWidth / 2.0;
            var halfAssemblyHeight = assemblyHeight / 2.0;

            if (assemblyWidth >= MinimumDimensionSpanMillimeters)
                CreateHorizontalOverallDimensionSet(
                    handler,
                    selectedView,
                    new TSG.Point(-halfAssemblyWidth, -halfAssemblyHeight, 0),
                    new TSG.Point(halfAssemblyWidth, -halfAssemblyHeight, 0),
                    OverallDimensionOffsetMillimeters,
                    new TSG.Vector(0.0, -1.0, 0.0)
                );

            if (assemblyHeight >= MinimumDimensionSpanMillimeters)
                CreateVerticalOverallDimensionSet(
                    handler,
                    selectedView,
                    new TSG.Point(-halfAssemblyWidth, -halfAssemblyHeight, 0),
                    new TSG.Point(-halfAssemblyWidth, halfAssemblyHeight, 0),
                    OverallDimensionOffsetMillimeters,
                    new TSG.Vector(-1.0, 0.0, 0.0)
                );
        }

        #endregion

        private static AssemblyViewBounds GetBoundsFromParts(
            TSD.View selectedView,
            List<TSM.Part> parts
        ) {
            if (selectedView == null || parts == null || parts.Count == 0) return null;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(new TSM.TransformationPlane());

                var displayCoordinateSystem = selectedView.DisplayCoordinateSystem;
                var origin = displayCoordinateSystem.Origin;
                var axisX = NormalizeVector(displayCoordinateSystem.AxisX);
                var axisY = NormalizeVector(displayCoordinateSystem.AxisY);

                var minimumX = double.MaxValue;
                var minimumY = double.MaxValue;
                var maximumX = double.MinValue;
                var maximumY = double.MinValue;
                var foundAnyPoint = false;

                foreach (var part in parts.Where(part => part != null)) {
                    TSM.Solid solid;

                    try {
                        solid = part.GetSolid();
                    }
                    catch {
                        continue;
                    }

                    var edgeEnumerator = solid?.GetEdgeEnumerator();
                    if (edgeEnumerator == null) continue;

                    while (edgeEnumerator.MoveNext()) {
                        dynamic edge = edgeEnumerator.Current;
                        if (edge == null) continue;

                        if (!TryGetEdgePoints(edge, out TSG.Point startPoint, out TSG.Point endPoint)) continue;

                        var projectedStartPoint = ProjectPointToViewPlane(startPoint, origin, axisX, axisY);
                        var projectedEndPoint = ProjectPointToViewPlane(endPoint, origin, axisX, axisY);

                        UpdateBounds(projectedStartPoint, ref minimumX, ref minimumY, ref maximumX, ref maximumY);
                        UpdateBounds(projectedEndPoint, ref minimumX, ref minimumY, ref maximumX, ref maximumY);
                        foundAnyPoint = true;
                    }
                }

                if (!foundAnyPoint) return null;

                return new AssemblyViewBounds {
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

        private static TSG.Point ProjectPointToViewPlane(
            TSG.Point point,
            TSG.Point origin,
            TSG.Vector axisX,
            TSG.Vector axisY
        ) {
            var vectorFromOrigin = new TSG.Vector(
                point.X - origin.X,
                point.Y - origin.Y,
                point.Z - origin.Z
            );

            var projectedX = Dot(vectorFromOrigin, axisX);
            var projectedY = Dot(vectorFromOrigin, axisY);

            return new TSG.Point(projectedX, projectedY, 0);
        }

        private static TSG.Vector NormalizeVector(TSG.Vector vector) {
            var length = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y + vector.Z * vector.Z);
            if (length < 0.000001) return new TSG.Vector(0, 0, 0);

            return new TSG.Vector(
                vector.X / length,
                vector.Y / length,
                vector.Z / length
            );
        }

        private static double Dot(TSG.Vector firstVector, TSG.Vector secondVector) {
            return firstVector.X * secondVector.X +
                   firstVector.Y * secondVector.Y +
                   firstVector.Z * secondVector.Z;
        }

        #region Data Structures

        private sealed class AssemblyViewBounds {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
        }

        #endregion

        #region Part Name Drawing

        private static void DrawPartFramesAndNames(
            TSD.View selectedView,
            List<TSM.Part> parts,
            AssemblyViewBounds assemblyBounds
        ) {
            if (selectedView == null || parts == null || parts.Count == 0 || assemblyBounds == null) return;

            var lineAttributes = new TSD.Line.LineAttributes {
                Line = new TSD.LineTypeAttributes(TSD.LineTypes.SolidLine, TSD.DrawingColors.Red)
            };

            var assemblyCenterX = (assemblyBounds.MinX + assemblyBounds.MaxX) / 2.0;
            var assemblyCenterY = (assemblyBounds.MinY + assemblyBounds.MaxY) / 2.0;

            foreach (var part in parts.Where(part => part != null)) {
                var rawBounds = GetPartBoundsInView(selectedView, part);
                if (rawBounds == null) continue;

                var shiftedBounds = ShiftBoundsToAssemblyCenter(rawBounds, assemblyCenterX, assemblyCenterY);

                DrawRectangleForBounds(selectedView, shiftedBounds, lineAttributes);

                var centerX = (shiftedBounds.MinX + shiftedBounds.MaxX) / 2.0;
                var centerY = (shiftedBounds.MinY + shiftedBounds.MaxY) / 2.0;
                var partName = string.IsNullOrWhiteSpace(part.Name) ? "Unnamed Part" : part.Name;

                var text = new TSD.Text(selectedView, new TSG.Point(centerX, centerY, 0), partName);
                text.Insert();
            }
        }

        private static AssemblyViewBounds GetPartBoundsInView(TSD.View selectedView, TSM.Part part) {
            if (selectedView == null || part == null) return null;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(new TSM.TransformationPlane());

                TSM.Solid solid;

                try {
                    solid = part.GetSolid();
                }
                catch {
                    return null;
                }

                var edgeEnumerator = solid?.GetEdgeEnumerator();
                if (edgeEnumerator == null) return null;

                var displayCoordinateSystem = selectedView.DisplayCoordinateSystem;
                var origin = displayCoordinateSystem.Origin;
                var axisX = NormalizeVector(displayCoordinateSystem.AxisX);
                var axisY = NormalizeVector(displayCoordinateSystem.AxisY);

                var minimumX = double.MaxValue;
                var minimumY = double.MaxValue;
                var maximumX = double.MinValue;
                var maximumY = double.MinValue;
                var foundAnyPoint = false;

                while (edgeEnumerator.MoveNext()) {
                    dynamic edge = edgeEnumerator.Current;
                    if (edge == null) continue;

                    if (!TryGetEdgePoints(edge, out TSG.Point startPoint, out TSG.Point endPoint)) continue;

                    var projectedStartPoint = ProjectPointToViewPlane(startPoint, origin, axisX, axisY);
                    var projectedEndPoint = ProjectPointToViewPlane(endPoint, origin, axisX, axisY);

                    UpdateBounds(projectedStartPoint, ref minimumX, ref minimumY, ref maximumX, ref maximumY);
                    UpdateBounds(projectedEndPoint, ref minimumX, ref minimumY, ref maximumX, ref maximumY);
                    foundAnyPoint = true;
                }

                if (!foundAnyPoint) return null;

                return new AssemblyViewBounds {
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

        private static AssemblyViewBounds ShiftBoundsToAssemblyCenter(
            AssemblyViewBounds bounds,
            double assemblyCenterX,
            double assemblyCenterY
        ) {
            if (bounds == null) return null;

            return new AssemblyViewBounds {
                MinX = bounds.MinX - assemblyCenterX,
                MaxX = bounds.MaxX - assemblyCenterX,
                MinY = bounds.MinY - assemblyCenterY,
                MaxY = bounds.MaxY - assemblyCenterY
            };
        }

        private static void DrawRectangleForBounds(
            TSD.View selectedView,
            AssemblyViewBounds bounds,
            TSD.Line.LineAttributes lineAttributes
        ) {
            if (selectedView == null || bounds == null) return;

            var bottomLeft = new TSG.Point(bounds.MinX, bounds.MinY, 0);
            var bottomRight = new TSG.Point(bounds.MaxX, bounds.MinY, 0);
            var topRight = new TSG.Point(bounds.MaxX, bounds.MaxY, 0);
            var topLeft = new TSG.Point(bounds.MinX, bounds.MaxY, 0);

            new TSD.Line(selectedView, bottomLeft, bottomRight, lineAttributes).Insert();
            new TSD.Line(selectedView, bottomRight, topRight, lineAttributes).Insert();
            new TSD.Line(selectedView, topRight, topLeft, lineAttributes).Insert();
            new TSD.Line(selectedView, topLeft, bottomLeft, lineAttributes).Insert();
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
            if (point.X < minimumX) minimumX = point.X;
            if (point.Y < minimumY) minimumY = point.Y;
            if (point.X > maximumX) maximumX = point.X;
            if (point.Y > maximumY) maximumY = point.Y;
        }

        #endregion

        #region Dimension Constants

        private const double OverallDimensionOffsetMillimeters = 20.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;

        #endregion

        #region Assembly Resolution

        private static TSM.Assembly ResolveAssemblyToDimension(
            TSD.DrawingHandler drawingHandler,
            List<TSM.Part> modelPartsInView,
            string assemblyInput
        ) {
            var assembliesInView = GetDistinctAssembliesFromParts(modelPartsInView);
            if (assembliesInView.Count == 0) return null;

            var assemblyFromInput = TryGetAssemblyFromInput(assembliesInView, modelPartsInView, assemblyInput);
            if (assemblyFromInput != null) return assemblyFromInput;

            var assemblyFromSelection = TryGetAssemblyFromSelectedDrawingPart(drawingHandler);
            if (assemblyFromSelection != null) return assemblyFromSelection;

            if (assembliesInView.Count == 1) return assembliesInView[0];

            var visibleAssemblyByPartCount = assembliesInView
                .Select(assembly => new {
                    Assembly = assembly,
                    VisiblePartCount = CountVisibleAssemblyPartsInView(assembly, modelPartsInView)
                })
                .OrderByDescending(item => item.VisiblePartCount)
                .FirstOrDefault();

            return visibleAssemblyByPartCount?.Assembly;
        }

        private static TSM.Assembly TryGetAssemblyFromSelectedDrawingPart(TSD.DrawingHandler drawingHandler) {
            var selector = drawingHandler.GetDrawingObjectSelector();
            var selectedObjects = selector.GetSelected();
            if (selectedObjects == null) return null;

            var model = new TSM.Model();
            selectedObjects.SelectInstances = true;

            while (selectedObjects.MoveNext()) {
                if (!(selectedObjects.Current is TSD.Part drawingPart)) continue;

                try {
                    var modelObject = model.SelectModelObject(drawingPart.ModelIdentifier);
                    if (!(modelObject is TSM.Part modelPart)) continue;

                    var assembly = modelPart.GetAssembly();
                    if (assembly != null) return assembly;
                }
                catch {
                }
            }

            return null;
        }

        private static List<TSM.Assembly> GetDistinctAssembliesFromParts(List<TSM.Part> parts) {
            var assemblies = new List<TSM.Assembly>();
            var addedAssemblyIdentifiers = new HashSet<int>();

            foreach (var part in parts.Where(part => part != null)) {
                TSM.Assembly assembly;

                try {
                    assembly = part.GetAssembly();
                }
                catch {
                    continue;
                }

                if (assembly == null) continue;

                var assemblyIdentifier = assembly.Identifier.ID;
                if (!addedAssemblyIdentifiers.Add(assemblyIdentifier)) continue;

                assemblies.Add(assembly);
            }

            return assemblies;
        }

        private static TSM.Assembly TryGetAssemblyFromInput(
            List<TSM.Assembly> assemblies,
            List<TSM.Part> modelPartsInView,
            string assemblyInput
        ) {
            if (assemblies == null || assemblies.Count == 0) return null;
            if (string.IsNullOrWhiteSpace(assemblyInput)) return null;

            var normalizedInput = assemblyInput.Trim();

            foreach (var assembly in assemblies) {
                if (!(assembly.GetMainPart() is TSM.Part mainPart)) continue;

                var assemblyPosition = string.Empty;
                var mainPartName = mainPart.Name ?? string.Empty;

                try {
                    mainPart.GetReportProperty("ASSEMBLY_POS", ref assemblyPosition);
                }
                catch {
                }

                if (assemblyPosition.IndexOf(normalizedInput, StringComparison.OrdinalIgnoreCase) >= 0)
                    return assembly;

                if (mainPartName.IndexOf(normalizedInput, StringComparison.OrdinalIgnoreCase) >= 0)
                    return assembly;
            }

            return (from assembly in assemblies
                let visibleParts = GetAssemblyPartsVisibleInView(assembly, modelPartsInView)
                where visibleParts.Select(part => part.Name ?? string.Empty).Any(partName =>
                    partName.IndexOf(normalizedInput, StringComparison.OrdinalIgnoreCase) >= 0)
                select assembly).FirstOrDefault();
        }

        #endregion

        #region Assembly Parts

        private static int CountVisibleAssemblyPartsInView(TSM.Assembly assembly, List<TSM.Part> modelPartsInView) {
            if (assembly == null || modelPartsInView == null || modelPartsInView.Count == 0) return 0;

            var visiblePartIdentifiers = new HashSet<int>(modelPartsInView.Select(part => part.Identifier.ID));
            var assemblyParts = GetAllAssemblyPartsRecursive(assembly);

            return assemblyParts.Count(part => visiblePartIdentifiers.Contains(part.Identifier.ID));
        }

        private static List<TSM.Part> GetAssemblyPartsVisibleInView(
            TSM.Assembly assembly,
            List<TSM.Part> modelPartsInView
        ) {
            var visiblePartIdentifiers = new HashSet<int>(modelPartsInView.Select(part => part.Identifier.ID));
            var assemblyParts = GetAllAssemblyPartsRecursive(assembly);

            return assemblyParts
                .Where(part => visiblePartIdentifiers.Contains(part.Identifier.ID))
                .ToList();
        }

        private static List<TSM.Part> GetAllAssemblyPartsRecursive(TSM.Assembly assembly) {
            var result = new Dictionary<int, TSM.Part>();
            CollectAssemblyPartsRecursive(assembly, result);
            return result.Values.ToList();
        }

        private static void CollectAssemblyPartsRecursive(
            TSM.Assembly assembly,
            Dictionary<int, TSM.Part> result
        ) {
            if (assembly == null) return;

            if (assembly.GetMainPart() is TSM.Part mainPart && !result.ContainsKey(mainPart.Identifier.ID))
                result[mainPart.Identifier.ID] = mainPart;

            var secondaries = assembly.GetSecondaries();
            if (secondaries != null)
                foreach (var secondary in secondaries) {
                    if (!(secondary is TSM.Part secondaryPart)) continue;

                    if (!result.ContainsKey(secondaryPart.Identifier.ID))
                        result[secondaryPart.Identifier.ID] = secondaryPart;
                }

            var subAssemblies = assembly.GetSubAssemblies();
            if (subAssemblies != null)
                foreach (var subAssemblyObject in subAssemblies) {
                    if (!(subAssemblyObject is TSM.Assembly subAssembly)) continue;
                    CollectAssemblyPartsRecursive(subAssembly, result);
                }
        }

        #endregion

        #region Dimension Set Builders

        private static void CreateHorizontalOverallDimensionSet(
            TSD.StraightDimensionSetHandler handler,
            TSD.View view,
            TSG.Point startPoint,
            TSG.Point endPoint,
            double offset,
            TSG.Vector directionVector
        ) {
            if (Math.Abs(endPoint.X - startPoint.X) < MinimumDimensionSpanMillimeters) return;

            var points = new TSD.PointList {
                startPoint,
                endPoint
            };

            handler.CreateDimensionSet(view, points, directionVector, offset);
        }

        private static void CreateVerticalOverallDimensionSet(
            TSD.StraightDimensionSetHandler handler,
            TSD.View view,
            TSG.Point startPoint,
            TSG.Point endPoint,
            double offset,
            TSG.Vector directionVector
        ) {
            if (Math.Abs(endPoint.Y - startPoint.Y) < MinimumDimensionSpanMillimeters) return;

            var points = new TSD.PointList {
                startPoint,
                endPoint
            };

            handler.CreateDimensionSet(view, points, directionVector, offset);
        }

        #endregion
    }
}