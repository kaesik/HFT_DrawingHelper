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
            if (assemblyParts.Count == 0) {
                MessageBox.Show("Nie znaleziono części wybranego assembly na zaznaczonym widoku.");
                return;
            }

            var bounds = GetAssemblyBoundsInView(selectedView, assemblyParts);
            if (bounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obrysu assembly.");
                return;
            }

            var xCoordinates = GetDistinctBoundaryCoordinates(
                selectedView,
                assemblyParts,
                bounds.MinX,
                bounds.MaxX,
                true
            );

            var yCoordinates = GetDistinctBoundaryCoordinates(
                selectedView,
                assemblyParts,
                bounds.MinY,
                bounds.MaxY,
                false
            );

            CreateAssemblyDimensionSets(selectedView, bounds, xCoordinates, yCoordinates);
            drawing.CommitChanges();
        }

        #endregion


        #region Dimension Creation

        private static void CreateAssemblyDimensionSets(
            TSD.View selectedView,
            AssemblyViewBounds bounds,
            List<double> xCoordinates,
            List<double> yCoordinates
        ) {
            var handler = new TSD.StraightDimensionSetHandler();

            CreateHorizontalChainDimensionSet(handler, selectedView, xCoordinates, bounds.MaxY,
                ChainDimensionOffsetMillimeters);

            CreateHorizontalOverallDimensionSet(handler, selectedView, bounds.MinX, bounds.MaxX, bounds.MaxY,
                OverallDimensionOffsetMillimeters);

            CreateVerticalChainDimensionSet(handler, selectedView, yCoordinates, bounds.MaxX,
                ChainDimensionOffsetMillimeters);

            CreateVerticalOverallDimensionSet(handler, selectedView, bounds.MinY, bounds.MaxY, bounds.MaxX,
                OverallDimensionOffsetMillimeters);
        }

        #endregion


        #region Data Structures

        private sealed class AssemblyViewBounds {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
        }

        #endregion

        #region Dimension Constants

        private const double ChainDimensionOffsetMillimeters = 80.0;
        private const double OverallDimensionOffsetMillimeters = 140.0;
        private const double CoordinateToleranceMillimeters = 1.0;
        private const double MinimumDimensionSpanMillimeters = 1.0;

        #endregion


        #region Assembly Resolution

        private static TSM.Assembly ResolveAssemblyToDimension(
            TSD.DrawingHandler drawingHandler,
            List<TSM.Part> modelPartsInView,
            string assemblyInput
        ) {
            var assemblyFromSelection = TryGetAssemblyFromSelectedDrawingPart(drawingHandler);
            if (assemblyFromSelection != null) return assemblyFromSelection;

            var assembliesInView = GetDistinctAssembliesFromParts(modelPartsInView);
            if (assembliesInView.Count == 0) return null;

            var assemblyFromInput = TryGetAssemblyFromInput(assembliesInView, assemblyInput);
            if (assemblyFromInput != null) return assemblyFromInput;

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
                    // ignored
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

        private static TSM.Assembly TryGetAssemblyFromInput(List<TSM.Assembly> assemblies, string assemblyInput) {
            if (assemblies == null || assemblies.Count == 0) return null;
            if (string.IsNullOrWhiteSpace(assemblyInput)) return null;

            var normalizedInput = assemblyInput.Trim();

            foreach (var assembly in assemblies) {
                if (!(assembly.GetMainPart() is TSM.Part mainPart)) continue;

                var assemblyPosition = string.Empty;
                var partName = mainPart.Name ?? string.Empty;

                try {
                    mainPart.GetReportProperty("ASSEMBLY_POS", ref assemblyPosition);
                }
                catch {
                    // ignroed
                }

                if (string.Equals(assemblyPosition, normalizedInput, StringComparison.OrdinalIgnoreCase))
                    return assembly;

                if (string.Equals(partName, normalizedInput, StringComparison.OrdinalIgnoreCase))
                    return assembly;
            }

            return null;
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


        #region Geometry Extraction

        private static AssemblyViewBounds GetAssemblyBoundsInView(
            TSD.View selectedView,
            List<TSM.Part> assemblyParts
        ) {
            if (selectedView == null || assemblyParts == null || assemblyParts.Count == 0) return null;

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem)
                );

                var allPoints = new List<TSG.Point>();

                foreach (var part in assemblyParts.Where(part => part != null)) {
                    TSM.Solid solid;

                    try {
                        solid = part.GetSolid();
                    }
                    catch {
                        continue;
                    }

                    if (solid == null) continue;

                    var min = solid.MinimumPoint;
                    var max = solid.MaximumPoint;

                    allPoints.Add(new TSG.Point(min.X, min.Y, 0));
                    allPoints.Add(new TSG.Point(max.X, min.Y, 0));
                    allPoints.Add(new TSG.Point(max.X, max.Y, 0));
                    allPoints.Add(new TSG.Point(min.X, max.Y, 0));
                }

                if (allPoints.Count == 0) return null;

                return new AssemblyViewBounds {
                    MinX = allPoints.Min(p => p.X),
                    MaxX = allPoints.Max(p => p.X),
                    MinY = allPoints.Min(p => p.Y),
                    MaxY = allPoints.Max(p => p.Y)
                };
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
        }

        private static List<double> GetDistinctBoundaryCoordinates(
            TSD.View selectedView,
            List<TSM.Part> assemblyParts,
            double min,
            double max,
            bool useXAxis
        ) {
            var result = new List<double>();

            var model = new TSM.Model();
            var workPlaneHandler = model.GetWorkPlaneHandler();
            var savedPlane = workPlaneHandler.GetCurrentTransformationPlane();

            try {
                workPlaneHandler.SetCurrentTransformationPlane(
                    new TSM.TransformationPlane(selectedView.DisplayCoordinateSystem)
                );

                foreach (var part in assemblyParts.Where(part => part != null)) {
                    TSM.Solid solid;

                    try {
                        solid = part.GetSolid();
                    }
                    catch {
                        continue;
                    }

                    if (solid == null) continue;

                    if (useXAxis) {
                        result.Add(solid.MinimumPoint.X);
                        result.Add(solid.MaximumPoint.X);
                    }
                    else {
                        result.Add(solid.MinimumPoint.Y);
                        result.Add(solid.MaximumPoint.Y);
                    }
                }
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }

            result.Add(min);
            result.Add(max);

            return MergeCloseCoordinates(result, CoordinateToleranceMillimeters);
        }

        private static List<double> MergeCloseCoordinates(List<double> coordinates, double tolerance) {
            if (coordinates == null || coordinates.Count == 0) return new List<double>();

            var ordered = coordinates.OrderBy(x => x).ToList();
            var result = new List<double> { ordered[0] };

            for (var i = 1; i < ordered.Count; i++)
                if (Math.Abs(ordered[i] - result.Last()) > tolerance)
                    result.Add(ordered[i]);

            return result;
        }

        #endregion


        #region Dimension Set Builders

        private static void CreateHorizontalChainDimensionSet(
            TSD.StraightDimensionSetHandler handler,
            TSD.View view,
            List<double> xs,
            double y,
            double offset
        ) {
            if (xs == null || xs.Count < 2) return;
            if (xs.Last() - xs.First() < MinimumDimensionSpanMillimeters) return;

            var points = new TSD.PointList();

            foreach (var x in xs)
                points.Add(new TSG.Point(x, y, 0));

            handler.CreateDimensionSet(view, points, new TSG.Vector(0, 1, 0), offset);
        }

        private static void CreateHorizontalOverallDimensionSet(
            TSD.StraightDimensionSetHandler handler,
            TSD.View view,
            double minX,
            double maxX,
            double y,
            double offset
        ) {
            if (maxX - minX < MinimumDimensionSpanMillimeters) return;

            var points = new TSD.PointList {
                new TSG.Point(minX, y, 0),
                new TSG.Point(maxX, y, 0)
            };

            handler.CreateDimensionSet(view, points, new TSG.Vector(0, 1, 0), offset);
        }

        private static void CreateVerticalChainDimensionSet(
            TSD.StraightDimensionSetHandler handler,
            TSD.View view,
            List<double> ys,
            double x,
            double offset
        ) {
            if (ys == null || ys.Count < 2) return;
            if (ys.Last() - ys.First() < MinimumDimensionSpanMillimeters) return;

            var points = new TSD.PointList();

            foreach (var y in ys)
                points.Add(new TSG.Point(x, y, 0));

            handler.CreateDimensionSet(view, points, new TSG.Vector(1, 0, 0), offset);
        }

        private static void CreateVerticalOverallDimensionSet(
            TSD.StraightDimensionSetHandler handler,
            TSD.View view,
            double minY,
            double maxY,
            double x,
            double offset
        ) {
            if (maxY - minY < MinimumDimensionSpanMillimeters) return;

            var points = new TSD.PointList {
                new TSG.Point(x, minY, 0),
                new TSG.Point(x, maxY, 0)
            };

            handler.CreateDimensionSet(view, points, new TSG.Vector(1, 0, 0), offset);
        }

        #endregion
    }
}