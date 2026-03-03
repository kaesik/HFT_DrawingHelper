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

        #endregion

        #region Button Clicks

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            AddSections();
        }

        private void AddDimensionsButton_Click(object sender, RoutedEventArgs e) {
            AddDimensions();
        }

        #endregion

        #region Add sections

        private static void AddSections() {
            var drawingHandler = new TSD.DrawingHandler();

            if (drawingHandler.GetConnectionStatus()) {
                var drawing = drawingHandler.GetActiveDrawing();

                if (drawing == null) return;

                CreateSectionViewOnSelectedView(drawingHandler, drawing);
            }
        }

        #endregion

        #region Add dimensions

        private void AddDimensions() {
            var elements = ElementsToDimensionTextBox.Text;
            var drawingHandler = new TSD.DrawingHandler();


            if (drawingHandler.GetConnectionStatus()) {
                var drawing = drawingHandler.GetActiveDrawing();

                if (drawing == null) return;

                ArrangeDrawingObject(drawing);
                ArrangeDrawingDim(drawing);
                ArrangeDrawingView(drawing);
            }
        }

        #endregion

        #region Helpers

        #region Arrange

        private static void ArrangeDrawingObject(TSD.Drawing drawing) {
            var containerView = drawing.GetSheet();
            var viewEnumerator = containerView.GetViews();

            if (viewEnumerator != null) {
                viewEnumerator.SelectInstances = false;

                while (viewEnumerator.MoveNext()) {
                    if (viewEnumerator.Current is TSD.View view) {
                        var drawingMarkEnumerator = view.GetAllObjects(typeof(TSD.MarkBase));

                        // Przykładowe działanie: przesunięcie o pewną odległość
                        while (drawingMarkEnumerator.MoveNext()) {
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
            }
        }

        private static void ArrangeDrawingDim(TSD.Drawing drawing) {
            var containerView = drawing.GetSheet();
            var viewEnumerator = containerView.GetViews();

            if (viewEnumerator != null) {
                while (viewEnumerator.MoveNext()) {
                    if (viewEnumerator.Current is TSD.View view) {
                        var dimensionEnumerator = view.GetAllObjects(typeof(TSD.StraightDimensionSet));

                        while (dimensionEnumerator.MoveNext()) {
                            if (dimensionEnumerator.Current is TSD.StraightDimensionSet dimensionSet) {
                                dimensionSet.Modify();
                            }
                        }
                    }
                }
            }
        }

        private static void ArrangeDrawingView(TSD.Drawing drawing) {
            var containerView = drawing.GetSheet();
            var viewEnumerator = containerView.GetViews();

            if (viewEnumerator != null) {
                while (viewEnumerator.MoveNext()) {
                    var view = viewEnumerator.Current as TSD.View;

                    // Przykładowe działanie: przesunięcie widoku o pewną odległość
                    if (view != null) {
                        var newOrigin = new TSG.Point(view.Origin.X + 50, view.Origin.Y + 50, view.Origin.Z);
                        view.Origin = newOrigin;
                    }

                    view?.Modify();
                }
            }
        }

        #endregion

        #region Create

        private static void CreateSectionViewOnSelectedView(TSD.DrawingHandler drawingHandler, TSD.Drawing drawing) {
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
                if (selectedView != null)
                    break;
            }

            if (selectedView == null) {
                MessageBox.Show("Zaznacz widok na rysunku, a potem uruchom funkcję.");
                return;
            }

            // CreateSingleSectionView(selectedView);
            var modelParts = GetModelPartsFromDrawingView(selectedView);

            var contourPlate = FindFirstContourPlate(modelParts);
            if (contourPlate == null) {
                MessageBox.Show("Nie znaleziono ContourPlate na tym widoku.");
                return;
            }

            GetContourPlateEdgePoints(contourPlate);

            drawing.CommitChanges();
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
                    if (partTypeCounters.ContainsKey(partTypeName)) {
                        partTypeCounters[partTypeName] += 1;
                    }
                    else {
                        partTypeCounters.Add(partTypeName, 1);
                    }
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

        private static void GetContourPlateEdgePoints(TSM.ContourPlate contourPlate) {
            if (contourPlate == null) {
                MessageBox.Show("ContourPlate jest null.");
                return;
            }

            if (contourPlate.Contour?.ContourPoints == null) {
                MessageBox.Show("ContourPlate nie ma konturu lub punktów konturu.");
                return;
            }

            var contourPoints = contourPlate.Contour.ContourPoints;
            if (contourPoints.Count < 2) {
                MessageBox.Show("Za mało punktów konturu, aby utworzyć krawędzie.");
                return;
            }

            var outputLines = new List<string> {
                "Punkty krawędzi ContourPlate (kolejne punkty konturu):",
                ""
            };

            for (var index = 0; index < contourPoints.Count; index++) {
                var firstContourPointObject = contourPoints[index] as TSM.ContourPoint;
                var secondContourPointObject = contourPoints[(index + 1) % contourPoints.Count] as TSM.ContourPoint;

                if (firstContourPointObject == null || secondContourPointObject == null) continue;

                var firstPoint = new TSG.Point(firstContourPointObject.X, firstContourPointObject.Y,
                    firstContourPointObject.Z);
                var secondPoint = new TSG.Point(secondContourPointObject.X, secondContourPointObject.Y,
                    secondContourPointObject.Z);

                outputLines.Add(
                    "Krawędź " + (index + 1) + ": " +
                    "(" + firstPoint.X + ", " + firstPoint.Y + ", " + firstPoint.Z + ")" +
                    " -> " +
                    "(" + secondPoint.X + ", " + secondPoint.Y + ", " + secondPoint.Z + ")"
                );
            }

            MessageBox.Show(string.Join("\n", outputLines));
        }

        private static (TSD.View.ViewAttributes, TSD.SectionMarkBase.SectionMarkAttributes) GetSectionAttributes() {
            var view = new TSD.View.ViewAttributes("#HFT_Kant_Section");
            var mark = new TSD.SectionMarkBase.SectionMarkAttributes();
            mark.LoadAttributes("#HFT_SECTION_V");

            return (view, mark);
        }

        #endregion

        #region Other

        private static TSM.ContourPlate FindFirstContourPlate(List<TSM.Part> modelParts) {
            if (modelParts == null || modelParts.Count == 0) return null;

            foreach (var modelPart in modelParts) {
                if (modelPart is TSM.ContourPlate contourPlate) return contourPlate;
            }

            return null;
        }

        #endregion

        #endregion
    }
}