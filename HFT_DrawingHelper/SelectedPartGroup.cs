using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;

namespace HFT_DrawingHelper {
    public sealed class SelectedPartGroup : INotifyPropertyChanged {
        private bool? _isChecked;
        private bool _isExpanded;
        private bool _isPreviewSelected;
        private bool _isUpdatingChildren;

        public SelectedPartGroup(string groupName, ObservableCollection<SelectedPartItem> items) {
            GroupName = groupName;
            Items = items ?? new ObservableCollection<SelectedPartItem>();

            Items.CollectionChanged += Items_CollectionChanged;

            foreach (var item in Items)
                SubscribeItem(item);

            RefreshCheckState();
        }

        public string GroupName { get; }

        public ObservableCollection<SelectedPartItem> Items { get; }

        public bool IsExpanded {
            get => _isExpanded;
            set {
                if (_isExpanded == value) return;
                _isExpanded = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsExpanded)));
            }
        }

        public bool? IsChecked {
            get => _isChecked;
            private set {
                if (_isChecked == value) return;
                _isChecked = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsChecked)));
            }
        }

        public bool IsPreviewSelected {
            get => _isPreviewSelected;
            set {
                if (_isPreviewSelected == value) return;
                _isPreviewSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsPreviewSelected)));
            }
        }

        public int ItemCount => Items.Count;

        public event PropertyChangedEventHandler PropertyChanged;

        public void SetCheckedForAll(bool value) {
            _isUpdatingChildren = true;

            foreach (var item in Items)
                item.IsChecked = value;

            _isUpdatingChildren = false;

            RefreshCheckState();
        }

        private void Items_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e) {
            if (e.OldItems != null)
                foreach (SelectedPartItem item in e.OldItems)
                    UnsubscribeItem(item);

            if (e.NewItems != null)
                foreach (SelectedPartItem item in e.NewItems)
                    SubscribeItem(item);

            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ItemCount)));
            RefreshCheckState();
        }

        private void SubscribeItem(SelectedPartItem item) {
            if (item == null) return;
            item.PropertyChanged += Item_PropertyChanged;
        }

        private void UnsubscribeItem(SelectedPartItem item) {
            if (item == null) return;
            item.PropertyChanged -= Item_PropertyChanged;
        }

        private void Item_PropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (_isUpdatingChildren) return;

            if (e.PropertyName == nameof(SelectedPartItem.IsChecked))
                RefreshCheckState();
        }

        private void RefreshCheckState() {
            if (Items.Count == 0) {
                IsChecked = false;
                return;
            }

            var checkedCount = Items.Count(item => item.IsChecked);

            if (checkedCount == 0) {
                IsChecked = false;
                return;
            }

            if (checkedCount == Items.Count) {
                IsChecked = true;
                return;
            }

            IsChecked = null;
        }
    }
}