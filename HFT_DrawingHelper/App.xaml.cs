using System.Windows;

namespace HFT_DrawingHelper {
    public partial class App {
        protected override void OnStartup(StartupEventArgs e) {
            base.OnStartup(e);
            ThemeService.Initialize();
        }
    }
}