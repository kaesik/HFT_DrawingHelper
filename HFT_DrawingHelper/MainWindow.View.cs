using System;
using System.Windows;
using System.Windows.Controls;
using TSD = Tekla.Structures.Drawing;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        private const double RotateViewMinimumPointDistance = 0.001;
        private const string RubberbandDrawingPluginName = "RubberbandDrawing";

        private void RotateViewByTwoPoints_Click(object sender, RoutedEventArgs e) {
            RotateViewByTwoPoints();
        }

        private void RotateViewByTwoPoints() {
            var drawingHandler = new TSD.DrawingHandler();

            if (!drawingHandler.GetConnectionStatus()) {
                MessageBox.Show("Brak połączenia z Tekla Structures.");
                return;
            }

            var drawing = drawingHandler.GetActiveDrawing();

            if (drawing == null) {
                MessageBox.Show("Brak aktywnego rysunku.");
                return;
            }

            try {
                var picker = drawingHandler.GetPicker();

                TSG.Point basePoint;
                TSG.Point secondPoint;
                TSD.ViewBase pickedView;

                picker.PickTwoPoints(
                    "Wskaż punkt bazowy obrotu widoku.",
                    "Wskaż drugi punkt wyznaczający kierunek poziomy.",
                    out basePoint,
                    out secondPoint,
                    out pickedView
                );

                var drawingView = pickedView as TSD.View;

                if (drawingView == null) {
                    MessageBox.Show("Wskazane punkty nie należą do zwykłego widoku rysunku.");
                    return;
                }

                var deltaX = secondPoint.X - basePoint.X;
                var deltaY = secondPoint.Y - basePoint.Y;

                var distance = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);

                if (distance < RotateViewMinimumPointDistance) {
                    MessageBox.Show("Wskazane punkty są zbyt blisko siebie.");
                    return;
                }

                var currentLineAngleDegrees = Math.Atan2(deltaY, deltaX) * 180.0 / Math.PI;
                var rotationAngleDegrees = NormalizeRotationAngle(-currentLineAngleDegrees);

                drawingView.Select();

                var rotated = drawingView.RotateViewOnDrawingPlane(rotationAngleDegrees);

                if (!rotated) {
                    MessageBox.Show("Nie udało się obrócić widoku.");
                    return;
                }

                drawing.CommitChanges("Obrócono widok");

                if (ShouldRunRubberbandDrawingPlugin())
                    PickObjectAndInsertRubberbandDrawingPlugin(drawingHandler, drawing);
            }
            catch (TSD.PickerInterruptedException) {
            }
            catch (Exception exception) {
                MessageBox.Show("Nie udało się obrócić widoku.\n" + exception.Message);
            }
        }

        private bool ShouldRunRubberbandDrawingPlugin() {
            var checkBox = FindNamedDescendant<CheckBox>("RunRubberbandDrawingPluginCheckBox");
            return checkBox != null && checkBox.IsChecked == true;
        }

        private static void PickObjectAndInsertRubberbandDrawingPlugin(TSD.DrawingHandler drawingHandler,
            TSD.Drawing drawing) {
            try {
                var picker = drawingHandler.GetPicker();

                TSD.DrawingObject pickedObject;
                TSD.ViewBase pickedView;

                picker.PickObject(
                    "Wskaż element w widoku, dla którego ma zostać uruchomiony RubberbandDrawing.",
                    out pickedObject,
                    out pickedView
                );

                var drawingView = pickedView as TSD.View;

                if (pickedObject == null || drawingView == null) {
                    MessageBox.Show("Nie wskazano poprawnego elementu w zwykłym widoku rysunku.");
                    return;
                }

                drawingView.Select();

                var plugin = new TSD.Plugin(drawingView, RubberbandDrawingPluginName);

                var pickerInput = new TSD.PluginPickerInput();
                pickerInput.Add(new TSD.PickerInputObject(pickedObject));

                plugin.SetPickerInput(pickerInput);

                var inserted = plugin.Insert();

                if (!inserted) {
                    MessageBox.Show("Nie udało się uruchomić pluginu RubberbandDrawing.");
                    return;
                }

                drawing.CommitChanges("Uruchomiono RubberbandDrawing");
            }
            catch (TSD.PickerInterruptedException) {
            }
            catch (Exception exception) {
                MessageBox.Show("Nie udało się uruchomić pluginu RubberbandDrawing.\n" + exception.Message);
            }
        }

        private static double NormalizeRotationAngle(double angleDegrees) {
            while (angleDegrees <= -180.0) angleDegrees += 360.0;

            while (angleDegrees > 180.0) angleDegrees -= 360.0;

            return angleDegrees;
        }
    }
}