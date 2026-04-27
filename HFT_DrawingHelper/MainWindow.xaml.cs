using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Tekla.Structures.Dialog.UIControls;
using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using TSD = Tekla.Structures.Drawing;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        private static readonly TSM.Model MyModel = new TSM.Model();

        private readonly ObservableCollection<string> _availableMarkAttributeNames =
            new ObservableCollection<string>();

        private readonly ObservableCollection<string> _availableViewAttributeNames =
            new ObservableCollection<string>();

        private readonly ObservableCollection<SelectableEdgeGroup> _edgeGroups =
            new ObservableCollection<SelectableEdgeGroup>();

        private readonly Dictionary<int, List<TSD.DrawingObject>> _edgePreviewObjectsByGroupNumber =
            new Dictionary<int, List<TSD.DrawingObject>>();

        private bool _sectionAttributeOptionsLoaded;

        private SidePanelMode _sidePanelMode = SidePanelMode.None;

        public MainWindow() {
            if (MyModel.GetConnectionStatus()) {
                InitializeComponent();
                EdgeGroupsList.ItemsSource = _edgeGroups;
                ViewAttributeNameComboBox.ItemsSource = _availableViewAttributeNames;
                MarkAttributeNameComboBox.ItemsSource = _availableMarkAttributeNames;
                LoadSectionSettingsFallbackOnly();
                LoadSectionSettingsIntoPanel();
                ModelDrawingLabel.Text = MyModel.GetInfo().ModelName.Replace(".db1", "");
                return;
            }

            var connectionErrorWindow = new TeklaConnectionErrorWindow();
            connectionErrorWindow.ShowDialog();
            Close();
        }

        private void DrawEdgesButton_Click(object sender, RoutedEventArgs e) {
            DrawEdgesWithNumbers();
        }

        private void GetEdgesButton_Click(object sender, RoutedEventArgs e) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            ClearAllEdgePreviews();

            var selectableGroups = BuildSelectableEdgeGroups(drawingHandler);

            if (selectableGroups.Count == 0) {
                MessageBox.Show("Nie znaleziono krawędzi do wyboru.");
                return;
            }

            _edgeGroups.Clear();

            foreach (var group in selectableGroups)
                _edgeGroups.Add(group);

            SetSidePanelMode(SidePanelMode.Edges);
        }

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            var checkedEdgeNumbers = _edgeGroups
                .Where(group => group.IsChecked)
                .Select(group => group.GroupNumber)
                .OrderBy(number => number)
                .ToList();

            if (checkedEdgeNumbers.Count == 0) {
                MessageBox.Show("Nie zaznaczono żadnych krawędzi do przekrojów.");
                return;
            }

            AddSections(FormatEdgeNumbersForInput(checkedEdgeNumbers));
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

        private void SelectAllEdgesButton_Click(object sender, RoutedEventArgs e) {
            foreach (var group in _edgeGroups)
                group.IsChecked = true;
        }

        private void DeselectAllEdgesButton_Click(object sender, RoutedEventArgs e) {
            foreach (var group in _edgeGroups)
                group.IsChecked = false;
        }

        private void CloseSidePanelButton_Click(object sender, RoutedEventArgs e) {
            if (_sidePanelMode == SidePanelMode.Parts)
                ResetPartSelectionPanelState();

            if (_sidePanelMode == SidePanelMode.Edges) {
                ClearAllEdgePreviews();
                _edgeGroups.Clear();
            }

            SetSidePanelMode(SidePanelMode.None);
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

        private void SettingsToggleButton_Click(object sender, RoutedEventArgs e) {
            ToggleSettingsPanel();
        }

        private void CloseSettingsPanelButton_Click(object sender, RoutedEventArgs e) {
            SetSettingsPanelVisibility(false);
        }

        private void ApplySectionSettingsButton_Click(object sender, RoutedEventArgs e) {
            var viewAttributeName = ViewAttributeNameComboBox.SelectedItem as string ??
                                    ViewAttributeNameComboBox.Text?.Trim();
            var markAttributeName = MarkAttributeNameComboBox.SelectedItem as string ??
                                    MarkAttributeNameComboBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(viewAttributeName) || string.IsNullOrWhiteSpace(markAttributeName)) {
                MessageBox.Show("Wybierz ViewAttributeName i MarkAttributeName z listy.");
                return;
            }

            UpdateSectionAttributeNames(viewAttributeName, markAttributeName);
            LoadSectionSettingsIntoPanel();
        }

        private void RefreshSectionSettingsButton_Click(object sender, RoutedEventArgs e) {
            ReloadSectionAttributeOptions();
        }

        private void ToggleSettingsPanel() {
            var isVisible = SettingsPanelBorder.Visibility != Visibility.Visible;
            SetSettingsPanelVisibility(isVisible);

            if (!isVisible) return;

            if (!_sectionAttributeOptionsLoaded)
                ReloadSectionAttributeOptions();
            else
                LoadSectionSettingsIntoPanel();
        }

        private void SetSettingsPanelVisibility(bool isVisible) {
            if (SettingsPanelBorder == null) return;
            SettingsPanelBorder.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ReloadSectionAttributeOptions() {
            try {
                ApplyAttributeCollectionSnapshot(_availableViewAttributeNames, GetAvailableViewAttributeNames());
                ApplyAttributeCollectionSnapshot(_availableMarkAttributeNames, GetAvailableMarkAttributeNames());
                _sectionAttributeOptionsLoaded = true;
            }
            catch {
                LoadSectionSettingsFallbackOnly();
            }

            LoadSectionSettingsIntoPanel();
        }

        private void LoadSectionSettingsIntoPanel() {
            if (ViewAttributeNameComboBox == null || MarkAttributeNameComboBox == null) return;

            EnsureAttributeOptionExists(_availableViewAttributeNames, _viewAttributeName, DefaultViewAttributeName);
            EnsureAttributeOptionExists(_availableMarkAttributeNames, _markAttributeName, DefaultMarkAttributeName);

            ViewAttributeNameComboBox.SelectedItem = _availableViewAttributeNames
                .FirstOrDefault(item => string.Equals(item, _viewAttributeName, StringComparison.OrdinalIgnoreCase));

            MarkAttributeNameComboBox.SelectedItem = _availableMarkAttributeNames
                .FirstOrDefault(item => string.Equals(item, _markAttributeName, StringComparison.OrdinalIgnoreCase));
        }

        private void LoadSectionSettingsFallbackOnly() {
            ApplyAttributeCollectionSnapshot(
                _availableViewAttributeNames,
                OrderAttributeNames(new[] { _viewAttributeName, DefaultViewAttributeName })
            );

            ApplyAttributeCollectionSnapshot(
                _availableMarkAttributeNames,
                OrderAttributeNames(new[] { _markAttributeName, DefaultMarkAttributeName })
            );

            LoadSectionSettingsIntoPanel();
        }

        private SectionAttributeOptionsSnapshot BuildSectionAttributeOptionsSnapshot() {
            return new SectionAttributeOptionsSnapshot(
                GetAvailableViewAttributeNames(),
                GetAvailableMarkAttributeNames()
            );
        }

        private static void ApplyAttributeCollectionSnapshot(
            ObservableCollection<string> targetCollection,
            IEnumerable<string> discoveredNames
        ) {
            if (targetCollection == null) return;

            targetCollection.Clear();
            foreach (var item in OrderAttributeNames(discoveredNames ?? Enumerable.Empty<string>()))
                targetCollection.Add(item);
        }

        private static void EnsureAttributeOptionExists(
            ObservableCollection<string> targetCollection,
            params string[] values
        ) {
            if (targetCollection == null || values == null) return;

            foreach (var value in values) {
                if (string.IsNullOrWhiteSpace(value)) continue;
                if (targetCollection.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase)))
                    continue;
                targetCollection.Add(value.Trim());
            }

            var orderedItems = OrderAttributeNames(targetCollection);
            targetCollection.Clear();
            foreach (var item in orderedItems)
                targetCollection.Add(item);
        }

        private static List<string> GetAvailableViewAttributeNames() {
            var candidateNames = GetAvailableAttributeNamesFromExtensions(
                new[] { "vi", "vw", "view" },
                _viewAttributeName,
                DefaultViewAttributeName
            );

            foreach (var attributeName in GetShallowAttributeNamesFromStandardDirectories())
                AddNormalizedAttributeName(candidateNames, attributeName);

            return OrderAttributeNames(candidateNames);
        }

        private static List<string> GetAvailableMarkAttributeNames() {
            var candidateNames = GetAvailableAttributeNamesFromExtensions(
                new[] { "csm", "mrk", "mark", "pm" },
                _markAttributeName,
                DefaultMarkAttributeName
            );

            foreach (var attributeName in GetShallowAttributeNamesFromStandardDirectories())
                AddNormalizedAttributeName(candidateNames, attributeName);

            return OrderAttributeNames(candidateNames);
        }

        private static HashSet<string> GetAvailableAttributeNamesFromExtensions(
            IEnumerable<string> candidateExtensions,
            params string[] fallbackNames
        ) {
            var candidateNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (candidateExtensions != null)
                foreach (var extension in candidateExtensions) {
                    if (string.IsNullOrWhiteSpace(extension)) continue;

                    try {
                        var fileNames = EnvironmentFiles.GetMultiDirectoryFileList(extension);
                        if (fileNames != null)
                            foreach (var fileName in fileNames)
                                AddNormalizedAttributeName(candidateNames, fileName);
                    }
                    catch {
                    }

                    try {
                        var fileNames = EnvironmentFiles.GetAttributeFiles(extension);
                        if (fileNames != null)
                            foreach (var fileName in fileNames)
                                AddNormalizedAttributeName(candidateNames, fileName);
                    }
                    catch {
                    }
                }

            foreach (var fallbackName in fallbackNames)
                AddNormalizedAttributeName(candidateNames, fallbackName);

            return candidateNames;
        }

        private static IEnumerable<string> GetShallowAttributeNamesFromStandardDirectories() {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            List<string> directories;
            try {
                directories = EnvironmentFiles.GetStandardPropertyFileDirectories();
            }
            catch {
                return result;
            }

            if (directories == null || directories.Count == 0)
                return result;

            foreach (var directory in directories.Where(directory => !string.IsNullOrWhiteSpace(directory)))
            foreach (var candidateDirectory in GetShallowSearchDirectories(directory))
            foreach (var attributeName in EnumerateTopLevelAttributeNames(candidateDirectory))
                result.Add(attributeName);

            return result;
        }

        private static IEnumerable<string> GetShallowSearchDirectories(string directory) {
            var result = new List<string>();

            if (string.IsNullOrWhiteSpace(directory))
                return result;

            if (EnvironmentFiles.IsValidDirectory(directory))
                result.Add(directory);

            var attributesDirectory = Path.Combine(directory, "attributes");
            if (EnvironmentFiles.IsValidDirectory(attributesDirectory))
                result.Add(attributesDirectory);

            try {
                if (EnvironmentFiles.IsValidDirectory(directory))
                    foreach (var subDirectory in Directory.EnumerateDirectories(directory)
                                 .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
                        result.Add(subDirectory);

                        var subAttributesDirectory = Path.Combine(subDirectory, "attributes");
                        if (EnvironmentFiles.IsValidDirectory(subAttributesDirectory))
                            result.Add(subAttributesDirectory);
                    }
            }
            catch {
            }

            return result;
        }

        private static IEnumerable<string> EnumerateTopLevelAttributeNames(string directory) {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(directory) || !EnvironmentFiles.IsValidDirectory(directory))
                return result;

            try {
                foreach (var filePath in Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly)) {
                    var fileName = Path.GetFileName(filePath);
                    if (string.IsNullOrWhiteSpace(fileName))
                        continue;

                    if (!LooksLikeSectionAttributeFile(fileName))
                        continue;

                    var normalizedName = NormalizeAttributeName(fileName);
                    if (!string.IsNullOrWhiteSpace(normalizedName))
                        result.Add(normalizedName);
                }
            }
            catch {
            }

            return result;
        }

        private static bool LooksLikeSectionAttributeFile(string fileName) {
            if (string.IsNullOrWhiteSpace(fileName))
                return false;

            var normalizedName = NormalizeAttributeName(fileName);
            if (string.IsNullOrWhiteSpace(normalizedName))
                return false;

            if (normalizedName.StartsWith("#", StringComparison.OrdinalIgnoreCase))
                return true;

            if (normalizedName.IndexOf("section", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            if (normalizedName.IndexOf("schnitt", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            if (normalizedName.IndexOf("mark", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            return string.Equals(normalizedName, "standard", StringComparison.OrdinalIgnoreCase);
        }

        private static void AddCandidateIfLoadable(
            ICollection<string> targetCollection,
            string filePath,
            Func<string, bool> canLoadAttributes
        ) {
            var normalizedValue = NormalizeAttributeName(Path.GetFileName(filePath));
            if (string.IsNullOrWhiteSpace(normalizedValue)) return;
            if (!canLoadAttributes(normalizedValue)) return;

            targetCollection.Add(normalizedValue);
        }

        private static IEnumerable<string> EnumerateLikelyPropertyFiles() {
            List<string> directories;

            try {
                directories = EnvironmentFiles.GetStandardPropertyFileDirectories();
            }
            catch {
                return Enumerable.Empty<string>();
            }

            if (directories == null || directories.Count == 0)
                return Enumerable.Empty<string>();

            var uniqueFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<string>();

            foreach (var directory in directories.Where(directory => !string.IsNullOrWhiteSpace(directory))) {
                foreach (var filePath in EnumerateFilesFromDirectoryAndAttributes(directory))
                    if (uniqueFiles.Add(filePath))
                        result.Add(filePath);

                foreach (var subDirectory in EnumerateImmediateDirectories(directory)
                             .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
                foreach (var filePath in EnumerateFilesFromDirectoryAndAttributes(subDirectory))
                    if (uniqueFiles.Add(filePath))
                        result.Add(filePath);
            }

            return result;
        }

        private static IEnumerable<string> EnumerateFilesFromDirectoryAndAttributes(string directory) {
            var result = new List<string>();

            if (string.IsNullOrWhiteSpace(directory))
                return result;

            foreach (var currentDirectory in GetDirectoryCandidates(directory))
            foreach (var filePath in EnumerateTopLevelFiles(currentDirectory))
                result.Add(filePath);

            return result;
        }

        private static IEnumerable<string> GetDirectoryCandidates(string directory) {
            var result = new List<string>();

            if (string.IsNullOrWhiteSpace(directory))
                return result;

            if (EnvironmentFiles.IsValidDirectory(directory))
                result.Add(directory);

            var attributesDirectory = Path.Combine(directory, "attributes");
            if (EnvironmentFiles.IsValidDirectory(attributesDirectory))
                result.Add(attributesDirectory);

            return result;
        }

        private static IEnumerable<string> EnumerateImmediateDirectories(string directory) {
            if (string.IsNullOrWhiteSpace(directory) || !EnvironmentFiles.IsValidDirectory(directory))
                return Enumerable.Empty<string>();

            try {
                return Directory.EnumerateDirectories(directory).ToList();
            }
            catch {
                return Enumerable.Empty<string>();
            }
        }

        private static IEnumerable<string> EnumerateTopLevelFiles(string directory) {
            if (string.IsNullOrWhiteSpace(directory) || !EnvironmentFiles.IsValidDirectory(directory))
                return Enumerable.Empty<string>();

            try {
                return Directory.EnumerateFiles(directory, "*.*", SearchOption.TopDirectoryOnly).ToList();
            }
            catch {
                return Enumerable.Empty<string>();
            }
        }

        private static bool CanLoadViewAttributes(string attributeName) {
            if (string.IsNullOrWhiteSpace(attributeName)) return false;

            try {
                var attributes = new TSD.View.ViewAttributes();
                return attributes.LoadAttributes(attributeName.Trim());
            }
            catch {
                return false;
            }
        }

        private static bool CanLoadMarkAttributes(string attributeName) {
            if (string.IsNullOrWhiteSpace(attributeName)) return false;

            try {
                var attributes = new TSD.SectionMarkBase.SectionMarkAttributes();
                return attributes.LoadAttributes(attributeName.Trim());
            }
            catch {
                return false;
            }
        }

        private static void AddNormalizedAttributeName(ICollection<string> targetCollection, string value) {
            if (targetCollection == null) return;

            var normalizedValue = NormalizeAttributeName(value);
            if (string.IsNullOrWhiteSpace(normalizedValue)) return;

            targetCollection.Add(normalizedValue);
        }

        private static string NormalizeAttributeName(string value) {
            if (string.IsNullOrWhiteSpace(value)) return null;

            var normalizedValue = Path.GetFileNameWithoutExtension(value.Trim());
            return string.IsNullOrWhiteSpace(normalizedValue) ? null : normalizedValue;
        }

        private static List<string> OrderAttributeNames(IEnumerable<string> names) {
            return names
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => string.Equals(name, "standard", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private sealed class SectionAttributeOptionsSnapshot {
            public SectionAttributeOptionsSnapshot(
                IEnumerable<string> viewAttributeNames,
                IEnumerable<string> markAttributeNames
            ) {
                ViewAttributeNames = OrderAttributeNames(viewAttributeNames ?? Enumerable.Empty<string>());
                MarkAttributeNames = OrderAttributeNames(markAttributeNames ?? Enumerable.Empty<string>());
            }

            public IReadOnlyList<string> ViewAttributeNames { get; }

            public IReadOnlyList<string> MarkAttributeNames { get; }
        }

        private enum SidePanelMode {
            None,
            Parts,
            Edges
        }

        private sealed class SelectableEdgeGroup : INotifyPropertyChanged {
            private bool _isChecked;
            private bool _isPreviewVisible;

            public int GroupNumber { get; set; }

            public string DisplayName { get; set; }

            public string SecondaryInfo { get; set; }

            public NumberedEdgeGroup EdgeGroup { get; set; }

            public TSD.View TargetView { get; set; }

            public bool IsChecked {
                get => _isChecked;
                set {
                    if (_isChecked == value) return;
                    _isChecked = value;
                    OnPropertyChanged();
                }
            }

            public bool IsPreviewVisible {
                get => _isPreviewVisible;
                set {
                    if (_isPreviewVisible == value) return;
                    _isPreviewVisible = value;
                    OnPropertyChanged();
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;

            private void OnPropertyChanged([CallerMemberName] string propertyName = null) {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #region Edge Preview

        private void EdgeGroupCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            e.Handled = false;
        }

        private void EdgeGroupItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (IsOriginalSourceInsideCheckBox(e.OriginalSource)) return;
            if (!(sender is FrameworkElement element)) return;
            if (!(element.DataContext is SelectableEdgeGroup edgeGroup)) return;

            if (edgeGroup.IsPreviewVisible)
                HideEdgePreview(edgeGroup);
            else
                ShowEdgePreview(edgeGroup);

            e.Handled = true;
        }

        private static bool IsOriginalSourceInsideCheckBox(object originalSource) {
            var dependencyObject = originalSource as DependencyObject;

            while (dependencyObject != null) {
                if (dependencyObject is CheckBox)
                    return true;

                dependencyObject = VisualTreeHelper.GetParent(dependencyObject);
            }

            return false;
        }

        private void ShowEdgePreview(SelectableEdgeGroup selectableEdgeGroup) {
            if (selectableEdgeGroup?.EdgeGroup == null || selectableEdgeGroup.TargetView == null) return;
            if (selectableEdgeGroup.IsPreviewVisible) return;

            HideEdgePreview(selectableEdgeGroup);

            var createdObjects = new List<TSD.DrawingObject>();
            var edgeGroup = selectableEdgeGroup.EdgeGroup;
            var view = selectableEdgeGroup.TargetView;

            try {
                if (!edgeGroup.IsPolyline || edgeGroup.PolylinePoints == null || edgeGroup.PolylinePoints.Count < 2) {
                    if (edgeGroup.SectionEdge != null) {
                        var points = new List<TSG.Point> {
                            new TSG.Point(edgeGroup.SectionEdge.Item1.X, edgeGroup.SectionEdge.Item1.Y, 0),
                            new TSG.Point(edgeGroup.SectionEdge.Item2.X, edgeGroup.SectionEdge.Item2.Y, 0)
                        };

                        var drawingObject = DrawPreviewPolyline(view, points);
                        if (drawingObject != null)
                            createdObjects.Add(drawingObject);
                    }
                }
                else {
                    var drawingObject = DrawPreviewPolyline(
                        view,
                        edgeGroup.PolylinePoints.Select(point => new TSG.Point(point.X, point.Y, 0)).ToList()
                    );

                    if (drawingObject != null)
                        createdObjects.Add(drawingObject);
                }

                if (createdObjects.Count > 0) {
                    _edgePreviewObjectsByGroupNumber[selectableEdgeGroup.GroupNumber] = createdObjects;
                    selectableEdgeGroup.IsPreviewVisible = true;
                }

                CommitActiveDrawingChanges();
            }
            catch {
                foreach (var drawingObject in createdObjects)
                    try {
                        drawingObject?.Delete();
                    }
                    catch {
                        // ignored
                    }
            }
        }

        private void HideEdgePreview(SelectableEdgeGroup selectableEdgeGroup) {
            if (selectableEdgeGroup == null) return;

            if (_edgePreviewObjectsByGroupNumber.TryGetValue(selectableEdgeGroup.GroupNumber, out var previewObjects)) {
                foreach (var drawingObject in previewObjects)
                    try {
                        drawingObject?.Delete();
                    }
                    catch {
                        // ignored
                    }

                _edgePreviewObjectsByGroupNumber.Remove(selectableEdgeGroup.GroupNumber);
                CommitActiveDrawingChanges();
            }

            selectableEdgeGroup.IsPreviewVisible = false;
        }

        private void ClearAllEdgePreviews() {
            foreach (var drawingObject in _edgePreviewObjectsByGroupNumber.Values.SelectMany(previewGroup =>
                         previewGroup))
                try {
                    drawingObject?.Delete();
                }
                catch {
                    // ignored
                }

            _edgePreviewObjectsByGroupNumber.Clear();

            foreach (var edgeGroup in _edgeGroups)
                edgeGroup.IsPreviewVisible = false;

            CommitActiveDrawingChanges();
        }

        private static TSD.DrawingObject DrawPreviewPolyline(TSD.View view, List<TSG.Point> points) {
            if (view == null || points == null || points.Count < 2) return null;

            var cleanedPoints = RemoveNearDuplicates(
                new List<TSG.Point>(points),
                DuplicateToleranceMillimeters
            );

            if (cleanedPoints == null || cleanedPoints.Count < 2) return null;

            switch (cleanedPoints.Count) {
                case 2:
                    return DrawPreviewStraightSegment(view, cleanedPoints[0], cleanedPoints[1]);
                case 3: {
                    var first = cleanedPoints[0];
                    var last = cleanedPoints[cleanedPoints.Count - 1];

                    var distance = Math.Sqrt(
                        (first.X - last.X) * (first.X - last.X) +
                        (first.Y - last.Y) * (first.Y - last.Y)
                    );

                    if (distance <= DuplicateToleranceMillimeters)
                        return DrawPreviewStraightSegment(view, cleanedPoints[0], cleanedPoints[1]);

                    break;
                }
            }

            var polylinePoints = new TSD.PointList();
            foreach (var point in cleanedPoints)
                polylinePoints.Add(new TSG.Point(point.X, point.Y, 0));

            var polyline = new TSD.Polyline(view, polylinePoints);
            polyline.Attributes.Line.Color = TSD.DrawingColors.Red;
            polyline.Insert();

            return polyline;
        }

        private static TSD.DrawingObject DrawPreviewStraightSegment(TSD.View view, TSG.Point startPoint,
            TSG.Point endPoint) {
            if (view == null || startPoint == null || endPoint == null) return null;

            var distance = Math.Sqrt(
                (startPoint.X - endPoint.X) * (startPoint.X - endPoint.X) +
                (startPoint.Y - endPoint.Y) * (startPoint.Y - endPoint.Y)
            );

            if (distance <= DuplicateToleranceMillimeters) return null;

            var points = new TSD.PointList {
                new TSG.Point(startPoint.X, startPoint.Y, 0),
                new TSG.Point(endPoint.X, endPoint.Y, 0)
            };

            var polyline = new TSD.Polyline(view, points);
            polyline.Attributes.Line.Color = TSD.DrawingColors.Red;
            polyline.Insert();

            return polyline;
        }

        private static void CommitActiveDrawingChanges() {
            try {
                var drawingHandler = new TSD.DrawingHandler();
                if (!drawingHandler.GetConnectionStatus()) return;

                var activeDrawing = drawingHandler.GetActiveDrawing();
                activeDrawing?.CommitChanges();
            }
            catch {
                // ignored
            }
        }

        #endregion

        #region Helpers

        private bool TryApplyCheckedPartOverride() {
            if (_sidePanelMode != SidePanelMode.Parts || SidePanelBorder.Visibility != Visibility.Visible ||
                _partItems.Count == 0) {
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
            var dimensionType = AngledDimensionRadioButton.IsChecked == true
                ? DimensionType.Angled
                : CurvedDimensionRadioButton.IsChecked == true
                    ? DimensionType.Curved
                    : DimensionType.Straight;

            var horizontalScope = HorizontalTotalDimensionCheckBox.IsChecked == true
                ? DimensionScope.Overall
                : DimensionScope.PerElement;

            var verticalScope = VerticalTotalDimensionCheckBox.IsChecked == true
                ? DimensionScope.Overall
                : DimensionScope.PerElement;

            var useShortExtensionLine = ShortExtensionLineCheckBox.IsChecked == true;

            var optionsList = new List<DimensionOptions>();

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionAboveCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Horizontal,
                DimensionPlacement.Above,
                horizontalScope,
                useShortExtensionLine
            );

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionBelowCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Horizontal,
                DimensionPlacement.Below,
                horizontalScope,
                useShortExtensionLine
            );

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionRightCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Vertical,
                DimensionPlacement.Right,
                verticalScope,
                useShortExtensionLine
            );

            AddDimensionOptionIfChecked(
                optionsList,
                DimensionLeftCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Vertical,
                DimensionPlacement.Left,
                verticalScope,
                useShortExtensionLine
            );

            return optionsList;
        }

        private void SetSidePanelMode(SidePanelMode sidePanelMode) {
            _sidePanelMode = sidePanelMode;

            SidePanelBorder.Visibility = sidePanelMode == SidePanelMode.None
                ? Visibility.Collapsed
                : Visibility.Visible;

            if (PartSelectionPanel != null)
                PartSelectionPanel.Visibility = sidePanelMode == SidePanelMode.Parts
                    ? Visibility.Visible
                    : Visibility.Collapsed;

            if (EdgeSelectionPanel != null)
                EdgeSelectionPanel.Visibility = sidePanelMode == SidePanelMode.Edges
                    ? Visibility.Visible
                    : Visibility.Collapsed;
        }

        private static List<SelectableEdgeGroup> BuildSelectableEdgeGroups(TSD.DrawingHandler drawingHandler) {
            var selectedDrawingParts = FilterIgnoredSectionDrawingParts(GetSelectedDrawingParts(drawingHandler));
            TSD.View targetView = null;

            if (selectedDrawingParts.Count > 0) {
                targetView = GetCommonViewFromSelectedParts(selectedDrawingParts);

                if (targetView == null) {
                    MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                    return new List<SelectableEdgeGroup>();
                }
            }
            else {
                targetView = GetSelectedViewOrShowMessage(drawingHandler);
                if (targetView == null) return new List<SelectableEdgeGroup>();

                selectedDrawingParts = GetDrawingPartsFromView(targetView);
                if (selectedDrawingParts.Count == 0) return new List<SelectableEdgeGroup>();
            }

            var outlineSnapshots = GetPartOutlineSnapshots(targetView, selectedDrawingParts);
            if (outlineSnapshots.Count == 0) return new List<SelectableEdgeGroup>();

            var edgesByNumber = BuildEdgesByNumberFromOutlines(outlineSnapshots);
            if (edgesByNumber.Count == 0) return new List<SelectableEdgeGroup>();

            var groupedEdges = BuildNumberedEdgeGroups(
                edgesByNumber,
                JoinToleranceMillimeters,
                NearStraightAngleDegrees
            );

            groupedEdges = FilterShortNumberedEdgeGroups(
                groupedEdges,
                MinimumNumberedEdgeLengthMillimeters
            );

            return (from pair in groupedEdges.OrderBy(item => item.Key)
                select pair.Value
                into @group
                where @group != null
                let segmentCount = @group.IsPolyline
                    ? @group.EdgeSegments.Count
                    : 1
                let length = Math.Round(ComputeGroupLength(@group), 0)
                let typeLabel = @group.IsPolyline ? "łamana" : "prosta"
                select new SelectableEdgeGroup {
                    GroupNumber = @group.GroupNumber,
                    DisplayName = "Krawędź " + @group.GroupNumber,
                    SecondaryInfo = typeLabel + ", odcinki: " + segmentCount + ", długość: " + length + " mm",
                    EdgeGroup = @group,
                    TargetView = targetView,
                    IsChecked = false,
                    IsPreviewVisible = false
                }).ToList();
        }

        private static string FormatEdgeNumbersForInput(List<int> numbers) {
            if (numbers == null || numbers.Count == 0) return string.Empty;

            var orderedNumbers = numbers
                .Distinct()
                .OrderBy(number => number)
                .ToList();

            var result = new List<string>();
            var rangeStart = orderedNumbers[0];
            var previous = orderedNumbers[0];

            for (var index = 1; index < orderedNumbers.Count; index++) {
                var current = orderedNumbers[index];

                if (current == previous + 1) {
                    previous = current;
                    continue;
                }

                result.Add(rangeStart == previous
                    ? rangeStart.ToString()
                    : rangeStart + "-" + previous);

                rangeStart = current;
                previous = current;
            }

            result.Add(rangeStart == previous
                ? rangeStart.ToString()
                : rangeStart + "-" + previous);

            return string.Join(",", result);
        }

        private static void AddDimensionOptionIfChecked(
            List<DimensionOptions> optionsList,
            bool isChecked,
            DimensionType dimensionType,
            DimensionAxis axis,
            DimensionPlacement placement,
            DimensionScope scope,
            bool useShortExtensionLine
        ) {
            if (!isChecked) return;

            optionsList.Add(new DimensionOptions {
                DimensionType = dimensionType,
                Axis = axis,
                Placement = placement,
                Scope = scope,
                UseShortExtensionLine = useShortExtensionLine
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
                    if (ShouldIgnoreSectionModelPart(modelPart)) continue;

                    modelParts.Add(modelPart);
                }
                catch {
                    // ignored
                }
            }

            return modelParts;
        }


        private static List<DrawingPartWithBounds> FilterIgnoredSectionDrawingParts(
            List<DrawingPartWithBounds> drawingParts
        ) {
            if (drawingParts == null || drawingParts.Count == 0)
                return new List<DrawingPartWithBounds>();

            return drawingParts
                .Where(drawingPart => !ShouldIgnoreSectionDrawingPart(drawingPart))
                .ToList();
        }

        private static bool ShouldIgnoreSectionDrawingPart(DrawingPartWithBounds drawingPartWithBounds) {
            return drawingPartWithBounds == null || ShouldIgnoreSectionDrawingPart(drawingPartWithBounds.DrawingPart);
        }

        private static bool ShouldIgnoreSectionDrawingPart(TSD.Part drawingPart) {
            if (drawingPart == null) return true;
            if (drawingPart.ModelIdentifier == null) return false;

            try {
                var modelObject = MyModel.SelectModelObject(drawingPart.ModelIdentifier);
                return modelObject is TSM.Part modelPart && ShouldIgnoreSectionModelPart(modelPart);
            }
            catch {
                return false;
            }
        }

        private static bool ShouldIgnoreSectionModelPart(TSM.Part modelPart) {
            if (modelPart == null) return true;

            var name = modelPart.Name;
            if (string.IsNullOrWhiteSpace(name)) return false;

            var normalizedName = name.Trim().ToUpperInvariant();
            return normalizedName.StartsWith("WS") || normalizedName.StartsWith("MS");
        }

        private static List<DrawingPartWithBounds> GetDrawingPartsFromView(TSD.View drawingView) {
            var result = new List<DrawingPartWithBounds>();
            if (drawingView == null) return result;

            var drawingObjects = drawingView.GetAllObjects(typeof(TSD.ModelObject));
            if (drawingObjects == null) return result;

            var addedModelIdentifiers = new HashSet<int>();

            while (drawingObjects.MoveNext()) {
                if (!(drawingObjects.Current is TSD.Part drawingPart)) continue;
                if (drawingPart.Hideable != null && drawingPart.Hideable.IsHidden) continue;
                if (ShouldIgnoreSectionDrawingPart(drawingPart)) continue;

                var identifier = drawingPart.ModelIdentifier;
                if (identifier == null) continue;
                if (!addedModelIdentifiers.Add(identifier.ID)) continue;

                result.Add(new DrawingPartWithBounds {
                    DrawingPart = drawingPart
                });
            }

            return result;
        }

        #endregion
    }
}