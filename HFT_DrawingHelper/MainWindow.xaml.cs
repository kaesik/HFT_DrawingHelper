using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.IO;
using System.Collections;

using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TS = Tekla.Structures;
using TSD = Tekla.Structures.Drawing;
using TSMO = Tekla.Structures.Model.Operations;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        public MainWindow() {
            var myModel = new TSM.Model();
            
            if (myModel.GetConnectionStatus()) {
                InitializeComponent();
                ModelDrawingLabel.Content = myModel.GetInfo().ModelName.Replace(".db1", "");
            }
            else MessageBox.Show("Keine Verbindung zu Tekla Structures");
        }
    }
}