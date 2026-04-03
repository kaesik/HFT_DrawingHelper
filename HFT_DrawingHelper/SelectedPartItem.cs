using System.ComponentModel;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public sealed class SelectedPartItem : INotifyPropertyChanged {
        private bool _isChecked = true;
        private bool _isPreviewSelected;

        public bool IsChecked {
            get => _isChecked;
            set {
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

        public string DisplayName { get; set; }
        public string SecondaryInfo { get; set; }
        public string GroupName { get; set; }
        public TSD.Part DrawingPart { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}