using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        private static readonly TSM.Model MyModel = new TSM.Model();

        public MainWindow() {
            if (MyModel.GetConnectionStatus()) {
                InitializeComponent();
                ModelDrawingLabel.Text = MyModel.GetInfo().ModelName.Replace(".db1", "");
            }
            else
                MessageBox.Show("Brak połączenia z Tekla Structures.");
        }

        private void DrawEdgesButton_Click(object sender, RoutedEventArgs e) {
            var formattedNumbers = DrawEdgesWithNumbers();

            if (!string.IsNullOrWhiteSpace(formattedNumbers)) EdgeNumbersTextBox.Text = formattedNumbers;
        }

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            AddSections(EdgeNumbersTextBox.Text);
        }

        private void AddDimensionsButton_Click(object sender, RoutedEventArgs e) {
            var dimensionType = CurvedDimensionRadioButton.IsChecked == true
                ? DimensionType.Curved
                : DimensionType.Straight;

            var horizontalPlacement = DimensionAboveRadioButton.IsChecked == true
                ? HorizontalDimensionPlacement.Above
                : HorizontalDimensionPlacement.Below;

            var verticalPlacement = DimensionRightRadioButton.IsChecked == true
                ? VerticalDimensionPlacement.Right
                : VerticalDimensionPlacement.Left;

            var options = new DimensionOptions {
                DimensionType = dimensionType,
                HorizontalPlacement = horizontalPlacement,
                VerticalPlacement = verticalPlacement,
                CreateHorizontal = HorizontalDimensionCheckBox.IsChecked == true,
                CreateVertical = VerticalDimensionCheckBox.IsChecked == true
            };

            if (!options.CreateHorizontal && !options.CreateVertical) {
                MessageBox.Show("Zaznacz co najmniej jeden wymiar do utworzenia.");
                return;
            }

            AddDimensions(options);
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        private void ThemeToggle_Click(object sender, RoutedEventArgs e) {
        }

        #region Helpers

        private static TSD.View GetSelectedViewOrShowMessage(TSD.DrawingHandler drawingHandler) {
            var objectSelector = drawingHandler.GetDrawingObjectSelector();
            var selectedObjects = objectSelector.GetSelected();

            if (selectedObjects == null) {
                MessageBox.Show("Nie zaznaczono żadnego obiektu.");
                return null;
            }

            TSD.View selectedView = null;

            selectedObjects.SelectInstances = false;
            while (selectedObjects.MoveNext()) {
                selectedView = selectedObjects.Current as TSD.View;
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
            var addedModelIdentifiers = new HashSet<int>();

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

                var identifier = (TS.Identifier)modelIdentifierObject;
                if (!addedModelIdentifiers.Add(identifier.ID)) continue;

                try {
                    var modelObject = MyModel.SelectModelObject(identifier);
                    if (!(modelObject is TSM.Part modelPart)) continue;

                    modelParts.Add(modelPart);
                }
                catch {
                    // ignored
                }
            }

            return modelParts;
        }

        #endregion
    }
}