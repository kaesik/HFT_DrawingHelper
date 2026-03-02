using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.IO;
using System.Collections;

using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TS = Tekla.Structures;
using TSD = Tekla.Structures.Drawing;
using TSMO = Tekla.Structures.Model.Operations;

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

            CreateSingleSectionView(selectedView);

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
        
        private static List<TSG.Point> GetContourPointsInViewPlane(TSM.ContourPlate contourPlate, TSD.View view) {
            var contourPointsInViewPlane = new List<TSG.Point>();
            var contourPoints = contourPlate.Contour.ContourPoints;

            foreach (var point in contourPoints) {
                var contourPoint = point as TSM.ContourPoint;
                if (contourPoint == null) continue;

                switch (view.ViewType) {
                    case TSD.View.ViewTypes.FrontView:
                        contourPointsInViewPlane.Add(new TSG.Point(contourPoint.X, contourPoint.Y, 0));
                        break;
                    case TSD.View.ViewTypes.TopView:
                        contourPointsInViewPlane.Add(new TSG.Point(contourPoint.X, contourPoint.Z, 0));
                        break;
                    default:
                        contourPointsInViewPlane.Add(new TSG.Point(contourPoint.Y, contourPoint.Z, 0));
                        break;
                }
            }

            return contourPointsInViewPlane;
        }
        
        private static (TSD.View.ViewAttributes, TSD.SectionMarkBase.SectionMarkAttributes) GetSectionAttributes() {
            var view = new TSD.View.ViewAttributes("#HFT_Kant_Section");
            var mark = new TSD.SectionMarkBase.SectionMarkAttributes();
            mark.LoadAttributes("#HFT_SECTION_V");

            return (view, mark);
        }

        #endregion
    }
}