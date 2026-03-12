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

            var bounds = GetAssemblyBoundsInView(selectedView, assemblyParts);
            if (bounds == null) {
                MessageBox.Show("Nie udało się wyznaczyć obrysu całego elementu.");
                return;
            }

            CreateAssemblyDimensionSets(selectedView, bounds);
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
                    new TSG.Vector(0.0, 0.0, 0.0)
                );

            if (assemblyHeight >= MinimumDimensionSpanMillimeters)
                CreateVerticalOverallDimensionSet(
                    handler,
                    selectedView,
                    new TSG.Point(-halfAssemblyWidth, -halfAssemblyHeight, 0),
                    new TSG.Point(-halfAssemblyWidth, halfAssemblyHeight, 0),
                    OverallDimensionOffsetMillimeters,
                    new TSG.Vector(0.0, 0.0, 0.0)
                );
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
                    MinX = allPoints.Min(point => point.X),
                    MaxX = allPoints.Max(point => point.X),
                    MinY = allPoints.Min(point => point.Y),
                    MaxY = allPoints.Max(point => point.Y)
                };
            }
            finally {
                workPlaneHandler.SetCurrentTransformationPlane(savedPlane);
            }
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

        private const double OverallDimensionOffsetMillimeters = 20.0;
        private const double CoordinateToleranceMillimeters = 1.0;
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
                    // ignored
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