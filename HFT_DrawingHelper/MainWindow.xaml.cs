using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Tekla.Structures.Dialog.UIControls;
using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using TSD = Tekla.Structures.Drawing;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        private static readonly TSM.Model MyModel = new TSM.Model();

        private readonly ObservableCollection<string> _availableDimensionAttributeNames =
            new ObservableCollection<string>();

        private readonly ObservableCollection<string> _availableMarkAttributeNames =
            new ObservableCollection<string>();

        private readonly ObservableCollection<string> _availableViewAttributeNames =
            new ObservableCollection<string>();

        private readonly ObservableCollection<SelectableEdgeGroup> _edgeGroups =
            new ObservableCollection<SelectableEdgeGroup>();

        private readonly Dictionary<int, List<TSD.DrawingObject>> _edgePreviewObjectsByGroupNumber =
            new Dictionary<int, List<TSD.DrawingObject>>();

        private RadioButton _angledDimensionRadioButton;
        private RadioButton _curvedDimensionRadioButton;
        private CheckBox _dimensionAboveCheckBox;
        private ComboBox _dimensionAttributeNameComboBox;
        private CheckBox _dimensionBelowCheckBox;
        private CheckBox _dimensionLeftCheckBox;
        private CheckBox _dimensionRightCheckBox;
        private ListBox _edgeGroupsList;
        private FrameworkElement _edgeSelectionPanel;
        private CheckBox _horizontalTotalDimensionCheckBox;

        private bool _isResettingMainTabsStartupState;
        private bool _mainTabSelectionWasRequestedByUser;
        private ComboBox _markAttributeNameComboBox;
        private ListBox _partItemsList;
        private FrameworkElement _partSelectionPanel;

        private bool _sectionAttributeOptionsLoaded;
        private CheckBox _shortExtensionLineCheckBox;

        private SidePanelMode _sidePanelMode = SidePanelMode.None;
        private CheckBox _verticalTotalDimensionCheckBox;

        private ComboBox _viewAttributeNameComboBox;

        public MainWindow() {
            if (MyModel.GetConnectionStatus()) {
                InitializeComponent();
                InitializeSplitXamlReferences();
                WireSplitXamlEvents();
                _edgeGroupsList.ItemsSource = _edgeGroups;
                _viewAttributeNameComboBox.ItemsSource = _availableViewAttributeNames;
                _markAttributeNameComboBox.ItemsSource = _availableMarkAttributeNames;
                _dimensionAttributeNameComboBox.ItemsSource = _availableDimensionAttributeNames;
                LoadSavedSectionSettings();
                LoadSectionSettingsFallbackOnly();
                LoadSectionSettingsIntoPanel();
                var modelName = MyModel.GetInfo().ModelName.Replace(".db1", "");
                ModelDrawingLabel.Text = modelName;
                InitializeStartScreen(modelName);
                return;
            }

            var connectionErrorWindow = new TeklaConnectionErrorWindow();
            connectionErrorWindow.ShowDialog();
            Close();
        }


        private void InitializeSplitXamlReferences() {
            UpdateLayout();

            _viewAttributeNameComboBox = FindNamedDescendant<ComboBox>("ViewAttributeNameComboBox");
            _markAttributeNameComboBox = FindNamedDescendant<ComboBox>("MarkAttributeNameComboBox");
            _dimensionAttributeNameComboBox = FindNamedDescendant<ComboBox>("DimensionAttributeNameComboBox");
            _partItemsList = FindNamedDescendant<ListBox>("PartItemsList");
            _edgeGroupsList = FindNamedDescendant<ListBox>("EdgeGroupsList");
            _partSelectionPanel = FindNamedDescendant<FrameworkElement>("PartSelectionPanel");
            _edgeSelectionPanel = FindNamedDescendant<FrameworkElement>("EdgeSelectionPanel");
            FindNamedDescendant<Button>("DrawEdgesButton");
            FindNamedDescendant<Button>("GetEdgesButton");
            FindNamedDescendant<Button>("AddSectionsButton");
            FindNamedDescendant<Button>("ShowSelectedPartsButton");
            FindNamedDescendant<Button>("AddDimensionsButton");
            FindNamedDescendant<Button>("RefreshSectionSettingsButton");
            FindNamedDescendant<Button>("ApplySectionSettingsButton");
            _angledDimensionRadioButton = FindNamedDescendant<RadioButton>("AngledDimensionRadioButton");
            _curvedDimensionRadioButton = FindNamedDescendant<RadioButton>("CurvedDimensionRadioButton");
            _dimensionAboveCheckBox = FindNamedDescendant<CheckBox>("DimensionAboveCheckBox");
            _dimensionBelowCheckBox = FindNamedDescendant<CheckBox>("DimensionBelowCheckBox");
            _horizontalTotalDimensionCheckBox = FindNamedDescendant<CheckBox>("HorizontalTotalDimensionCheckBox");
            _dimensionRightCheckBox = FindNamedDescendant<CheckBox>("DimensionRightCheckBox");
            _dimensionLeftCheckBox = FindNamedDescendant<CheckBox>("DimensionLeftCheckBox");
            _verticalTotalDimensionCheckBox = FindNamedDescendant<CheckBox>("VerticalTotalDimensionCheckBox");
            _shortExtensionLineCheckBox = FindNamedDescendant<CheckBox>("ShortExtensionLineCheckBox");
        }

        private void WireSplitXamlEvents() {
            AddClickHandler("CloseSettingsPanelButton", CloseSettingsPanelButton_Click);
            AddClickHandler("RefreshSectionSettingsButton", RefreshSectionSettingsButton_Click);
            AddClickHandler("ApplySectionSettingsButton", ApplySectionSettingsButton_Click);
            AddClickHandler("DrawEdgesButton", DrawEdgesButton_Click);
            AddClickHandler("GetEdgesButton", GetEdgesButton_Click);
            AddClickHandler("AddSectionsButton", AddSectionsButton_Click);
            AddClickHandler("ShowSelectedPartsButton", ShowSelectedPartsButton_Click);
            AddClickHandler("AddDimensionsButton", AddDimensionsButton_Click);
            AddClickHandler("ClosePartsSidePanelButton", CloseSidePanelButton_Click);
            AddClickHandler("CloseEdgesSidePanelButton", CloseSidePanelButton_Click);
            AddClickHandler("SelectAllPartsButton", SelectAllPartsButton_Click);
            AddClickHandler("DeselectAllPartsButton", DeselectAllPartsButton_Click);
            AddClickHandler("SelectAllEdgesButton", SelectAllEdgesButton_Click);
            AddClickHandler("DeselectAllEdgesButton", DeselectAllEdgesButton_Click);

            if (_partItemsList != null) {
                _partItemsList.AddHandler(ButtonBase.ClickEvent, new RoutedEventHandler(PartItemsList_Click), true);
                _partItemsList.AddHandler(MouseLeftButtonDownEvent,
                    new MouseButtonEventHandler(PartItemsList_MouseLeftButtonDown), true);
            }

            if (_edgeGroupsList != null)
                _edgeGroupsList.AddHandler(MouseLeftButtonDownEvent,
                    new MouseButtonEventHandler(EdgeGroupsList_MouseLeftButtonDown), true);
        }

        private void AddClickHandler(string elementName, RoutedEventHandler handler) {
            var button = FindNamedDescendant<ButtonBase>(elementName);
            if (button != null)
                button.Click += handler;
        }

        private void InitializeStartScreen(string modelName) {
            _isResettingMainTabsStartupState = true;
            _mainTabSelectionWasRequestedByUser = false;

            MainTabs.PreviewMouseLeftButtonDown += MainTabs_PreviewMouseLeftButtonDown;
            MainTabs.PreviewKeyDown += MainTabs_PreviewKeyDown;

            SetStartScreenModelName(modelName);
            ResetMainTabsToStartScreen();

            Dispatcher.BeginInvoke(new Action(() => {
                ResetMainTabsToStartScreen();

                Dispatcher.BeginInvoke(new Action(() => {
                    ResetMainTabsToStartScreen();
                    _isResettingMainTabsStartupState = false;
                }), DispatcherPriority.ApplicationIdle);
            }), DispatcherPriority.Loaded);
        }

        private void MainTabs_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (FindAncestorOrSelf<TabItem>(e.OriginalSource as DependencyObject) == null)
                return;

            _isResettingMainTabsStartupState = false;
            _mainTabSelectionWasRequestedByUser = true;
        }

        private void MainTabs_PreviewKeyDown(object sender, KeyEventArgs e) {
            if (e.Key != Key.Left && e.Key != Key.Right && e.Key != Key.Up && e.Key != Key.Down &&
                e.Key != Key.Home && e.Key != Key.End && e.Key != Key.Enter && e.Key != Key.Space)
                return;

            _isResettingMainTabsStartupState = false;
            _mainTabSelectionWasRequestedByUser = true;
        }

        private void MainTabs_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            if (!ReferenceEquals(e.OriginalSource, MainTabs))
                return;

            if (_isResettingMainTabsStartupState || !_mainTabSelectionWasRequestedByUser) {
                ResetMainTabsToStartScreen();
                return;
            }

            SetStartScreenVisibility(MainTabs.SelectedIndex < 0);
        }

        private void ResetMainTabsToStartScreen() {
            if (MainTabs == null)
                return;

            MainTabs.ApplyTemplate();
            SetStartScreenVisibility(true);

            MainTabs.SelectionChanged -= MainTabs_SelectionChanged;
            MainTabs.SelectedItem = null;
            MainTabs.SelectedIndex = -1;

            foreach (var tabItem in MainTabs.Items.OfType<TabItem>())
                tabItem.IsSelected = false;

            MainTabs.SelectionChanged += MainTabs_SelectionChanged;
        }

        private void SetStartScreenVisibility(bool isVisible) {
            var startScreen = GetMainTabsTemplateElement<FrameworkElement>("StartScreen");
            if (startScreen != null)
                startScreen.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }

        private void SetStartScreenModelName(string modelName) {
            var startScreenModelName = GetMainTabsTemplateElement<TextBlock>("StartScreenModelName");
            if (startScreenModelName != null)
                startScreenModelName.Text = modelName;
        }

        private T GetMainTabsTemplateElement<T>(string elementName) where T : FrameworkElement {
            if (MainTabs == null)
                return null;

            MainTabs.ApplyTemplate();
            return MainTabs.Template?.FindName(elementName, MainTabs) as T;
        }

        private void PartItemsList_Click(object sender, RoutedEventArgs e) {
            var checkBox = FindAncestorOrSelf<CheckBox>(e.OriginalSource as DependencyObject);
            if (checkBox == null || !checkBox.IsThreeState)
                return;

            GroupCheckBox_Click(checkBox, e);
        }

        private void PartItemsList_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (IsOriginalSourceInsideCheckBox(e.OriginalSource))
                return;

            var element = FindNearestFrameworkElementWithDataContext(e.OriginalSource as DependencyObject);
            if (element == null || element.DataContext == null)
                return;

            if (HasReadableProperty(element.DataContext, "Items") &&
                HasReadableProperty(element.DataContext, "GroupName"))
                GroupHeader_MouseLeftButtonDown(element, e);
            else
                PartItem_MouseLeftButtonDown(element, e);
        }

        private void EdgeGroupsList_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) {
            if (IsOriginalSourceInsideCheckBox(e.OriginalSource))
                return;

            var element = FindNearestFrameworkElementWithDataContext(e.OriginalSource as DependencyObject);
            if (element?.DataContext is SelectableEdgeGroup)
                EdgeGroupItem_MouseLeftButtonDown(element, e);
        }

        private static bool HasReadableProperty(object source, string propertyName) {
            return source?.GetType().GetProperty(propertyName) != null;
        }

        private T FindNamedDescendant<T>(string elementName) where T : FrameworkElement {
            return FindNamedDescendant<T>(this, elementName, new HashSet<DependencyObject>());
        }

        private static T FindNamedDescendant<T>(DependencyObject parent, string elementName,
            HashSet<DependencyObject> visited) where T : FrameworkElement {
            if (parent == null || !visited.Add(parent))
                return null;

            if (parent is T typedElement && typedElement.Name == elementName)
                return typedElement;

            var logicalChildren = LogicalTreeHelper.GetChildren(parent).OfType<DependencyObject>().ToList();
            foreach (var result in logicalChildren.Select(child => FindNamedDescendant<T>(child, elementName, visited))
                         .Where(result => result != null)) return result;

            int visualChildrenCount;
            try {
                visualChildrenCount = VisualTreeHelper.GetChildrenCount(parent);
            }
            catch {
                return null;
            }

            for (var index = 0; index < visualChildrenCount; index++) {
                var child = VisualTreeHelper.GetChild(parent, index);
                var result = FindNamedDescendant<T>(child, elementName, visited);
                if (result != null)
                    return result;
            }

            return null;
        }

        private static T FindAncestorOrSelf<T>(DependencyObject source) where T : DependencyObject {
            var current = source;

            while (current != null) {
                if (current is T typed)
                    return typed;

                current = GetParentDependencyObject(current);
            }

            return null;
        }

        private static FrameworkElement FindNearestFrameworkElementWithDataContext(DependencyObject source) {
            var current = source;

            while (current != null) {
                if (current is FrameworkElement frameworkElement && frameworkElement.DataContext != null)
                    return frameworkElement;

                current = GetParentDependencyObject(current);
            }

            return null;
        }

        private static DependencyObject GetParentDependencyObject(DependencyObject source) {
            if (source == null)
                return null;

            try {
                var visualParent = VisualTreeHelper.GetParent(source);
                if (visualParent != null)
                    return visualParent;
            }
            catch {
                // ignored
            }

            try {
                return LogicalTreeHelper.GetParent(source);
            }
            catch {
                return null;
            }
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
            var viewAttributeName = _viewAttributeNameComboBox.SelectedItem as string ??
                                    _viewAttributeNameComboBox.Text?.Trim();
            var markAttributeName = _markAttributeNameComboBox.SelectedItem as string ??
                                    _markAttributeNameComboBox.Text?.Trim();
            var dimensionAttributeName = _dimensionAttributeNameComboBox.SelectedItem as string ??
                                         _dimensionAttributeNameComboBox.Text?.Trim();

            if (string.IsNullOrWhiteSpace(viewAttributeName) ||
                string.IsNullOrWhiteSpace(markAttributeName) ||
                string.IsNullOrWhiteSpace(dimensionAttributeName)) {
                MessageBox.Show("Wybierz ViewAttributeName, MarkAttributeName i DimensionAttributeName z listy.");
                return;
            }

            UpdateSectionAttributeNames(viewAttributeName, markAttributeName);
            UpdateDimensionAttributeName(dimensionAttributeName);
            SaveCurrentSectionSettings();
            LoadSectionSettingsIntoPanel();
        }

        private void RefreshSectionSettingsButton_Click(object sender, RoutedEventArgs e) {
            ReloadSectionAttributeOptions();
        }

        private void LoadSavedSectionSettings() {
            var savedSettings = DrawingAttributeSettingsService.Load();
            if (savedSettings == null)
                return;

            UpdateSectionAttributeNames(
                string.IsNullOrWhiteSpace(savedSettings.ViewAttributeName)
                    ? _viewAttributeName
                    : savedSettings.ViewAttributeName,
                string.IsNullOrWhiteSpace(savedSettings.MarkAttributeName)
                    ? _markAttributeName
                    : savedSettings.MarkAttributeName
            );

            if (!string.IsNullOrWhiteSpace(savedSettings.DimensionAttributeName))
                UpdateDimensionAttributeName(savedSettings.DimensionAttributeName);
        }

        private static void SaveCurrentSectionSettings() {
            DrawingAttributeSettingsService.Save(
                _viewAttributeName,
                _markAttributeName,
                _dimensionAttributeName
            );
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
                ApplyAttributeCollectionSnapshot(_availableDimensionAttributeNames,
                    GetAvailableDimensionAttributeNames());
                _sectionAttributeOptionsLoaded = true;
            }
            catch {
                LoadSectionSettingsFallbackOnly();
            }

            LoadSectionSettingsIntoPanel();
        }

        private void LoadSectionSettingsIntoPanel() {
            if (_viewAttributeNameComboBox == null ||
                _markAttributeNameComboBox == null ||
                _dimensionAttributeNameComboBox == null) return;

            EnsureAttributeOptionExists(_availableViewAttributeNames, _viewAttributeName, DefaultViewAttributeName);
            EnsureAttributeOptionExists(_availableMarkAttributeNames, _markAttributeName, DefaultMarkAttributeName);
            EnsureAttributeOptionExists(_availableDimensionAttributeNames, _dimensionAttributeName,
                DefaultDimensionAttributeName);

            _viewAttributeNameComboBox.SelectedItem = _availableViewAttributeNames
                .FirstOrDefault(item => string.Equals(item, _viewAttributeName, StringComparison.OrdinalIgnoreCase));

            _markAttributeNameComboBox.SelectedItem = _availableMarkAttributeNames
                .FirstOrDefault(item => string.Equals(item, _markAttributeName, StringComparison.OrdinalIgnoreCase));

            _dimensionAttributeNameComboBox.SelectedItem = _availableDimensionAttributeNames
                .FirstOrDefault(
                    item => string.Equals(item, _dimensionAttributeName, StringComparison.OrdinalIgnoreCase));
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

            ApplyAttributeCollectionSnapshot(
                _availableDimensionAttributeNames,
                OrderAttributeNames(new[] { _dimensionAttributeName, DefaultDimensionAttributeName })
            );

            LoadSectionSettingsIntoPanel();
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

        private static List<string> GetAvailableDimensionAttributeNames() {
            var candidateNames = GetAvailableAttributeNamesFromExtensions(
                new[] { "dim", "sdim", "adim", "rdim", "cdim" },
                _dimensionAttributeName,
                DefaultDimensionAttributeName
            );

            foreach (var attributeName in GetShallowDimensionAttributeNamesFromStandardDirectories())
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
                        // ignored
                    }

                    try {
                        var fileNames = EnvironmentFiles.GetAttributeFiles(extension);
                        if (fileNames != null)
                            foreach (var fileName in fileNames)
                                AddNormalizedAttributeName(candidateNames, fileName);
                    }
                    catch {
                        // ignored
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

        private static IEnumerable<string> GetShallowDimensionAttributeNamesFromStandardDirectories() {
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
            foreach (var attributeName in EnumerateTopLevelDimensionAttributeNames(candidateDirectory))
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
                // ignored
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
                // ignored
            }

            return result;
        }

        private static IEnumerable<string> EnumerateTopLevelDimensionAttributeNames(string directory) {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(directory) || !EnvironmentFiles.IsValidDirectory(directory))
                return result;

            try {
                foreach (var filePath in Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly)) {
                    var fileName = Path.GetFileName(filePath);
                    if (string.IsNullOrWhiteSpace(fileName))
                        continue;

                    if (!LooksLikeDimensionAttributeFile(fileName))
                        continue;

                    var normalizedName = NormalizeAttributeName(fileName);
                    if (!string.IsNullOrWhiteSpace(normalizedName))
                        result.Add(normalizedName);
                }
            }
            catch {
                // ignored
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

            return normalizedName.IndexOf("mark", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   string.Equals(normalizedName, "standard", StringComparison.OrdinalIgnoreCase);
        }

        private static bool LooksLikeDimensionAttributeFile(string fileName) {
            if (string.IsNullOrWhiteSpace(fileName))
                return false;

            var normalizedName = NormalizeAttributeName(fileName);
            if (string.IsNullOrWhiteSpace(normalizedName))
                return false;

            var extension = Path.GetExtension(fileName)?.TrimStart('.');
            if (!string.IsNullOrWhiteSpace(extension)) {
                var dimensionExtensions = new[] { "dim", "sdim", "adim", "rdim", "cdim" };
                if (dimensionExtensions.Any(item => string.Equals(item, extension, StringComparison.OrdinalIgnoreCase)))
                    return true;
            }

            if (normalizedName.IndexOf("dim", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            if (normalizedName.IndexOf("dimension", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            return normalizedName.IndexOf("wymiar", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   string.Equals(normalizedName, "standard", StringComparison.OrdinalIgnoreCase);
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

        private sealed class SavedDrawingAttributeSettings {
            public string ViewAttributeName { get; set; }
            public string MarkAttributeName { get; set; }
            public string DimensionAttributeName { get; set; }
        }

        private static class DrawingAttributeSettingsService {
            private static readonly string SettingsPath =
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "HFT.DrawingHelper",
                    "drawing_attributes.txt"
                );

            public static SavedDrawingAttributeSettings Load() {
                try {
                    if (!File.Exists(SettingsPath))
                        return null;

                    var savedSettings = new SavedDrawingAttributeSettings();
                    var lines = File.ReadAllLines(SettingsPath);

                    foreach (var line in lines) {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        var separatorIndex = line.IndexOf('=');
                        if (separatorIndex < 0)
                            continue;

                        var key = line.Substring(0, separatorIndex).Trim();
                        var value = line.Substring(separatorIndex + 1).Trim();

                        if (string.Equals(key, "ViewAttributeName", StringComparison.OrdinalIgnoreCase))
                            savedSettings.ViewAttributeName = value;
                        else if (string.Equals(key, "MarkAttributeName", StringComparison.OrdinalIgnoreCase))
                            savedSettings.MarkAttributeName = value;
                        else if (string.Equals(key, "DimensionAttributeName", StringComparison.OrdinalIgnoreCase))
                            savedSettings.DimensionAttributeName = value;
                    }

                    return savedSettings;
                }
                catch {
                    return null;
                }
            }

            public static void Save(
                string viewAttributeName,
                string markAttributeName,
                string dimensionAttributeName
            ) {
                try {
                    var directoryPath = Path.GetDirectoryName(SettingsPath);
                    if (!string.IsNullOrEmpty(directoryPath))
                        Directory.CreateDirectory(directoryPath);

                    File.WriteAllLines(
                        SettingsPath,
                        new[] {
                            @"ViewAttributeName=" + NormalizeSavedValue(viewAttributeName),
                            @"MarkAttributeName=" + NormalizeSavedValue(markAttributeName),
                            @"DimensionAttributeName=" + NormalizeSavedValue(dimensionAttributeName)
                        }
                    );
                }
                catch {
                    // ignored
                }
            }

            private static string NormalizeSavedValue(string value) {
                return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
            }
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
            var dimensionType = _angledDimensionRadioButton.IsChecked == true
                ? DimensionType.Angled
                : _curvedDimensionRadioButton.IsChecked == true
                    ? DimensionType.Curved
                    : DimensionType.Straight;

            var horizontalScope = _horizontalTotalDimensionCheckBox.IsChecked == true
                ? DimensionScope.Overall
                : DimensionScope.PerElement;

            var verticalScope = _verticalTotalDimensionCheckBox.IsChecked == true
                ? DimensionScope.Overall
                : DimensionScope.PerElement;

            var useShortExtensionLine = _shortExtensionLineCheckBox.IsChecked == true;
            var dimensionAttributeName = _dimensionAttributeNameComboBox?.SelectedItem as string ??
                                         _dimensionAttributeNameComboBox?.Text?.Trim();

            if (!string.IsNullOrWhiteSpace(dimensionAttributeName))
                UpdateDimensionAttributeName(dimensionAttributeName);

            dimensionAttributeName = GetDimensionAttributeNameOrDefault(_dimensionAttributeName,
                DefaultDimensionAttributeName);

            var optionsList = new List<DimensionOptions>();

            AddDimensionOptionIfChecked(
                optionsList,
                _dimensionAboveCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Horizontal,
                DimensionPlacement.Above,
                horizontalScope,
                useShortExtensionLine,
                dimensionAttributeName
            );

            AddDimensionOptionIfChecked(
                optionsList,
                _dimensionBelowCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Horizontal,
                DimensionPlacement.Below,
                horizontalScope,
                useShortExtensionLine,
                dimensionAttributeName
            );

            AddDimensionOptionIfChecked(
                optionsList,
                _dimensionRightCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Vertical,
                DimensionPlacement.Right,
                verticalScope,
                useShortExtensionLine,
                dimensionAttributeName
            );

            AddDimensionOptionIfChecked(
                optionsList,
                _dimensionLeftCheckBox.IsChecked == true,
                dimensionType,
                DimensionAxis.Vertical,
                DimensionPlacement.Left,
                verticalScope,
                useShortExtensionLine,
                dimensionAttributeName
            );

            return optionsList;
        }

        private void SetSidePanelMode(SidePanelMode sidePanelMode) {
            _sidePanelMode = sidePanelMode;

            SidePanelBorder.Visibility = sidePanelMode == SidePanelMode.None
                ? Visibility.Collapsed
                : Visibility.Visible;

            SetSidePanelContentVisibility(
                "PartSelectionPanel",
                "PartsSidePanelHost",
                sidePanelMode == SidePanelMode.Parts
            );

            SetSidePanelContentVisibility(
                "EdgeSelectionPanel",
                "EdgesSidePanelHost",
                sidePanelMode == SidePanelMode.Edges
            );
        }

        private void SetSidePanelContentVisibility(string panelElementName, string hostElementName, bool isVisible) {
            var visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;

            var panelElement = FindNamedDescendant<FrameworkElement>(panelElementName);
            if (panelElement != null)
                panelElement.Visibility = visibility;

            var hostElement = FindNamedDescendant<ContentControl>(hostElementName);
            if (hostElement == null)
                return;

            hostElement.Visibility = visibility;

            if (hostElement.Content is FrameworkElement contentElement)
                contentElement.Visibility = visibility;
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
            bool useShortExtensionLine,
            string dimensionAttributeName
        ) {
            if (!isChecked) return;

            optionsList.Add(new DimensionOptions {
                DimensionType = dimensionType,
                Axis = axis,
                Placement = placement,
                Scope = scope,
                UseShortExtensionLine = useShortExtensionLine,
                AttributeFileName = dimensionAttributeName
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