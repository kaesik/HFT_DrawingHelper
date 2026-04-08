using System.Collections.Generic;
using System.Linq;
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
                return;
            }

            var connectionErrorWindow = new TeklaConnectionErrorWindow();
            connectionErrorWindow.ShowDialog();
            Close();
        }

        private void DrawEdgesButton_Click(object sender, RoutedEventArgs e) {
            var formattedNumbers = DrawEdgesWithNumbers();

            if (!string.IsNullOrWhiteSpace(formattedNumbers))
                EdgeNumbersTextBox.Text = formattedNumbers;
        }

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            AddSections(EdgeNumbersTextBox.Text);
        }

        private void AddDimensionsButton_Click(object sender, RoutedEventArgs e) {
            if (!TryApplyCheckedPartOverride()) return;

            var optionsList = BuildDimensionOptionsFromUi();

            if (optionsList.Count == 0) {
                MessageBox.Show("Nie zaznaczono żadnego położenia wymiarów do utworzenia.");
                _overrideSelectedParts = null;
                return;
            }

            try {
                AddDimensions(optionsList);
            }
            finally {
                _overrideSelectedParts = null;
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            DragMove();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e) {
            Close();
        }

        private void ThemeToggle_Click(object sender, RoutedEventArgs e) {
            ThemeService.Toggle();
        }

        #region Helpers

        private bool TryApplyCheckedPartOverride() {
            if (SidePanelBorder.Visibility != Visibility.Visible || _partItems.Count == 0) {
                _overrideSelectedParts = null;
                return true;
            }

            var checkedParts = _partItems
                .Where(item => item.IsChecked && item.DrawingPart != null)
                .Select(item => item.DrawingPart)
                .ToList();

            if (checkedParts.Count == 0) {
                MessageBox.Show("Nie zaznaczono żadnych elementów na liście.");
                return false;
            }

            _overrideSelectedParts = checkedParts;
            return true;
        }

        private List<DimensionOptions> BuildDimensionOptionsFromUi() {
            var dimensionType = CurvedDimensionRadioButton.IsChecked == true
                ? DimensionType.Curved
                : DimensionType.Straight;

            var horizontalScope = HorizontalTotalDimensionCheckBox.IsChecked == true
                ? DimensionScope.Overall
                : DimensionScope.PerElement;

            var verticalScope = VerticalTotalDimensionCheckBox.IsChecked == true
                ? DimensionScope.Overall
                : DimensionScope.PerElement;

            var optionsList = new List<DimensionOptions>();

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionAboveCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Horizontal,
                DimensionPlacement.Above,
                horizontalScope
            );

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionBelowCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Horizontal,
                DimensionPlacement.Below,
                horizontalScope
            );

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionRightCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Vertical,
                DimensionPlacement.Right,
                verticalScope
            );

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionLeftCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Vertical,
                DimensionPlacement.Left,
                verticalScope
            );

            return optionsList;
        }

        private static void AddDimensionOptionIfChecked(
            List<DimensionOptions> optionsList,
            bool isChecked,
            DimensionType dimensionType,
            DimensionAxis axis,
            DimensionPlacement placement,
            DimensionScope scope
        ) {
            if (!isChecked) return;

            optionsList.Add(new DimensionOptions {
                DimensionType = dimensionType,
                Axis = axis,
                Placement = placement,
                Scope = scope
            });
        }

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