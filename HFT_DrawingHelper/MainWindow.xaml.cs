using System.Windows;
using TSM = Tekla.Structures.Model;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        public MainWindow() {
            var myModel = new TSM.Model();

            if (myModel.GetConnectionStatus()) {
                InitializeComponent();
                ModelDrawingLabel.Content = myModel.GetInfo().ModelName.Replace(".db1", "");
            }
            else
                MessageBox.Show("Keine Verbindung zu Tekla Structures");
        }

        private void DrawEdgesButton_Click(object sender, RoutedEventArgs e) {
            var formattedNumbers = DrawEdgesWithNumbers();

            if (!string.IsNullOrWhiteSpace(formattedNumbers)) EdgeNumbersTextBox.Text = formattedNumbers;
        }

        private void AddSectionsButton_Click(object sender, RoutedEventArgs e) {
            AddSections(EdgeNumbersTextBox.Text);
        }

        private void AddDimensionsButton_Click(object sender, RoutedEventArgs e) {
            AddDimensions();
        }

        private void AddDimensions() {
            var elements = ElementsToDimensionTextBox.Text;
        }
    }
}