using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSG = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Main Section Creation Flow

        private static void AddSections(string edgeNumbersInput) {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return;

            var selectedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (selectedView == null) return;

            var detectionResult = DetectEdgesFromSelectedView(selectedView);

            if (!detectionResult.HasEdges) {
                MessageBox.Show(detectionResult.ErrorMessage);
                return;
            }

            var requestedGroupNumbers = ParseEdgeNumbers(edgeNumbersInput);

            var groupsByGroupNumber = BuildNumberedEdgeGroups(
                detectionResult.EdgesByNumber,
                JoinToleranceMillimeters,
                NearStraightAngleDegrees
            );

            var sectionEdgesByGroupNumber = groupsByGroupNumber
                .Where(pair => pair.Value?.SectionEdge != null)
                .ToDictionary(pair => pair.Key, pair => pair.Value.SectionEdge);

            var filteredSectionEdges = FilterEdgesOrShowMessage(sectionEdgesByGroupNumber, requestedGroupNumbers);
            if (filteredSectionEdges == null) return;

            CreateSectionViewsFromEdges(selectedView, filteredSectionEdges);
            drawing.CommitChanges();
        }

        #endregion


        #region Section View Creation

        private static void CreateSectionViewsFromEdges(
            TSD.View baseView,
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edgesByNumber
        ) {
            if (baseView == null) return;
            if (edgesByNumber == null || edgesByNumber.Count == 0) return;

            var (viewAttrs, markAttrs) = GetSectionAttributes();

            var baseBox = baseView.GetAxisAlignedBoundingBox();

            var cursorX = baseBox.UpperRight.X + Gap;
            var cursorY = baseBox.UpperRight.Y;

            foreach (var pair in edgesByNumber.OrderBy(x => x.Key)) {
                var edge = pair.Value;
                if (edge == null) continue;

                var edgeA = edge.Item1;
                var edgeB = edge.Item2;

                var midPoint = new TSG.Point(
                    (edgeA.X + edgeB.X) * 0.5,
                    (edgeA.Y + edgeB.Y) * 0.5,
                    (edgeA.Z + edgeB.Z) * 0.5
                );

                var dx = edgeB.X - edgeA.X;
                var dy = edgeB.Y - edgeA.Y;

                var length = Math.Sqrt(dx * dx + dy * dy);
                if (length < 1e-6) continue;

                var nx = -dy / length;
                var ny = dx / length;

                const double halfLength = SectionLineLengthMillimeters * 0.5;

                var startPoint = new TSG.Point(
                    midPoint.X - nx * halfLength,
                    midPoint.Y - ny * halfLength,
                    midPoint.Z
                );

                var endPoint = new TSG.Point(
                    midPoint.X + nx * halfLength,
                    midPoint.Y + ny * halfLength,
                    midPoint.Z
                );

                ForceSectionLookLeftOrUp(ref startPoint, ref endPoint);

                var insertionPoint = new TSG.Point(cursorX, cursorY, baseView.Origin.Z);

                var ok = TSD.View.CreateSectionView(
                    baseView,
                    startPoint,
                    endPoint,
                    insertionPoint,
                    DepthUp,
                    DepthDown,
                    viewAttrs,
                    markAttrs,
                    out var sectionView,
                    out var sectionMark
                );

                if (!ok || sectionView == null || sectionMark == null) continue;

                sectionView.Modify();
                sectionMark.Modify();

                var sectionBox = sectionView.GetAxisAlignedBoundingBox();

                var deltaX = cursorX - sectionBox.LowerLeft.X;
                var deltaY = cursorY - sectionBox.UpperLeft.Y;

                sectionView.Origin = new TSG.Point(
                    sectionView.Origin.X + deltaX,
                    sectionView.Origin.Y + deltaY,
                    sectionView.Origin.Z
                );

                sectionView.Modify();

                sectionBox = sectionView.GetAxisAlignedBoundingBox();
                cursorX = sectionBox.UpperRight.X + Gap;
            }
        }

        #endregion


        #region Section Attributes

        private static (TSD.View.ViewAttributes, TSD.SectionMarkBase.SectionMarkAttributes) GetSectionAttributes() {
            var view = new TSD.View.ViewAttributes(ViewAttributeName);
            var mark = new TSD.SectionMarkBase.SectionMarkAttributes();
            mark.LoadAttributes(MarkAttributeName);

            return (view, mark);
        }

        #endregion


        #region Section Orientation

        private static void ForceSectionLookLeftOrUp(ref TSG.Point startPoint, ref TSG.Point endPoint) {
            var dx = endPoint.X - startPoint.X;
            var dy = endPoint.Y - startPoint.Y;

            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 1e-9) return;

            var nx = -dy / length;
            var ny = dx / length;

            if (Math.Abs(ny) >= Math.Abs(nx)) {
                if (ny < 0) (startPoint, endPoint) = (endPoint, startPoint);
            }
            else {
                if (nx > 0) (startPoint, endPoint) = (endPoint, startPoint);
            }
        }

        #endregion


        #region Edge Filtering

        private static Dictionary<int, Tuple<TSG.Point, TSG.Point>> FilterEdgesOrShowMessage(
            Dictionary<int, Tuple<TSG.Point, TSG.Point>> edges,
            HashSet<int> requested
        ) {
            if (edges == null || edges.Count == 0) return null;
            if (requested == null || requested.Count == 0) return edges;

            var filtered = edges
                .Where(pair => requested.Contains(pair.Key))
                .ToDictionary(pair => pair.Key, pair => pair.Value);

            if (filtered.Count == 0) {
                MessageBox.Show("Nie znaleziono krawędzi o podanych numerach.");
                return null;
            }

            var missing = requested.Where(number => !edges.ContainsKey(number)).OrderBy(number => number).ToList();
            if (missing.Count > 0) MessageBox.Show("Brak krawędzi o numerach: " + string.Join(", ", missing));

            return filtered;
        }

        #endregion

        #region Section Constants

        private const double DepthUp = 1.0;
        private const double DepthDown = 1.0;
        private const double SectionLineLengthMillimeters = 300.0;
        private const double Gap = 10.0;

        #endregion
    }
}