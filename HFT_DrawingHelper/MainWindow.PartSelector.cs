using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        private static List<TSD.Part> _overrideSelectedParts;

        private readonly ObservableCollection<SelectedPartGroup> _partGroups =
            new ObservableCollection<SelectedPartGroup>();

        private readonly ObservableCollection<SelectedPartItem> _partItems =
            new ObservableCollection<SelectedPartItem>();

        private void ShowSelectedPartsButton_Click(object sender, RoutedEventArgs e) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) {
                MessageBox.Show("Brak połączenia z Tekla Structures.");
                return;
            }

            var drawingParts = FilterIgnoredSectionDrawingParts(GetSelectedDrawingParts(drawingHandler));

            _partItems.Clear();
            _partGroups.Clear();

            foreach (var dp in drawingParts) {
                if (dp?.DrawingPart == null) continue;

                var name = "(brak nazwy)";
                var secondary = "";

                try {
                    var modelObject = MyModel.SelectModelObject(dp.DrawingPart.ModelIdentifier);
                    if (modelObject is TSM.Part modelPart) {
                        name = string.IsNullOrWhiteSpace(modelPart.Name)
                            ? "(brak nazwy)"
                            : modelPart.Name.Trim();

                        secondary = modelPart.Profile?.ProfileString ?? "";
                    }
                }
                catch {
                    // ignored
                }

                _partItems.Add(new SelectedPartItem {
                    DisplayName = name,
                    GroupName = name,
                    SecondaryInfo = secondary,
                    DrawingPart = dp.DrawingPart,
                    IsChecked = true,
                    IsPreviewSelected = false
                });
            }

            if (_partItems.Count == 0) {
                MessageBox.Show("Nie zaznaczono żadnych elementów na rysunku.");
                return;
            }

            var groupedItems = _partItems
                .GroupBy(item => item.GroupName)
                .OrderBy(group => group.Key);

            foreach (var group in groupedItems)
                _partGroups.Add(new SelectedPartGroup(
                    group.Key,
                    new ObservableCollection<SelectedPartItem>(
                        group
                            .OrderBy(item => item.SecondaryInfo)
                            .ThenBy(item => item.DisplayName)
                    )
                ));

            PartItemsList.ItemsSource = _partGroups;
            SetSidePanelMode(SidePanelMode.Parts);
        }

        private void GroupCheckBox_Click(object sender, RoutedEventArgs e) {
            if (!(sender is CheckBox checkBox)) return;
            if (!(checkBox.DataContext is SelectedPartGroup group)) return;

            group.SetCheckedForAll(group.IsChecked != true);

            e.Handled = true;
        }

        private void GroupHeader_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (!(sender is FrameworkElement element)) return;
            if (!(element.DataContext is SelectedPartGroup group)) return;

            var shouldSelectGroup = !group.Items.All(item => item.IsPreviewSelected);

            foreach (var item in group.Items)
                item.IsPreviewSelected = shouldSelectGroup;

            RefreshGroupPreviewStates();
            SyncPreviewSelectionToDrawing();

            e.Handled = true;
        }

        private void PartItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (!(sender is FrameworkElement element)) return;
            if (!(element.DataContext is SelectedPartItem item)) return;
            if (item.DrawingPart == null) return;

            item.IsPreviewSelected = !item.IsPreviewSelected;

            RefreshGroupPreviewStates();
            SyncPreviewSelectionToDrawing();

            e.Handled = true;
        }

        private void RefreshGroupPreviewStates() {
            foreach (var group in _partGroups)
                group.IsPreviewSelected = group.Items.Any(item => item.IsPreviewSelected);
        }

        private void SyncPreviewSelectionToDrawing() {
            var selectedDrawingParts = _partGroups
                .SelectMany(group => group.Items)
                .Where(item => item.IsPreviewSelected && item.DrawingPart != null)
                .Select(item => item.DrawingPart)
                .ToList();

            if (selectedDrawingParts.Count == 0) {
                ClearDrawingSelectionInEditor();
                return;
            }

            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var selector = drawingHandler.GetDrawingObjectSelector();
            if (selector == null) return;

            var objectsToSelect = new ArrayList();
            foreach (var drawingPart in selectedDrawingParts)
                objectsToSelect.Add(drawingPart);

            selector.UnselectAllObjects();
            selector.SelectObjects(objectsToSelect, false);
        }

        private void ClearPreviewSelection() {
            foreach (var group in _partGroups) {
                group.IsPreviewSelected = false;

                foreach (var item in group.Items)
                    item.IsPreviewSelected = false;
            }
        }

        private static void ClearDrawingSelectionInEditor() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var selector = drawingHandler.GetDrawingObjectSelector();
            selector?.UnselectAllObjects();
        }

        private void SelectAllPartsButton_Click(object sender, RoutedEventArgs e) {
            foreach (var group in _partGroups)
                group.SetCheckedForAll(true);
        }

        private void DeselectAllPartsButton_Click(object sender, RoutedEventArgs e) {
            foreach (var group in _partGroups)
                group.SetCheckedForAll(false);
        }

        private void ResetPartSelectionPanelState() {
            _overrideSelectedParts = null;
            ClearPreviewSelection();
            ClearDrawingSelectionInEditor();
        }
    }
}