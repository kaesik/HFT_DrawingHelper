using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using TSD = Tekla.Structures.Drawing;
using TSG = Tekla.Structures.Geometry3d;

namespace HFT_DrawingHelper {
    public partial class MainWindow {
        #region Main Edge Drawing Flow

        private static string DrawEdgesWithNumbers() {
            var drawingHandler = new TSD.DrawingHandler();
            if (!drawingHandler.GetConnectionStatus()) return null;

            var drawing = drawingHandler.GetActiveDrawing();
            if (drawing == null) return null;

            var pickedView = GetSelectedViewOrShowMessage(drawingHandler);
            if (pickedView == null) return null;

            var drawingParts = GetSelectedDrawingParts(drawingHandler);
            TSD.View commonView;

            if (drawingParts.Count == 0) {
                var singleDrawingPart = GetSingleDrawingPartFromPickedView(pickedView);
                if (singleDrawingPart == null) {
                    MessageBox.Show(
                        "Jeśli nie zaznaczasz elementu, wskazany widok musi zawierać dokładnie jeden element typu Part.");
                    return null;
                }

                drawingParts = new List<DrawingPartWithBounds> {
                    singleDrawingPart
                };

                commonView = singleDrawingPart.DrawingPart.GetView() as TSD.View;
                if (commonView == null) {
                    MessageBox.Show("Nie udało się ustalić widoku elementu.");
                    return null;
                }
            }
            else {
                commonView = GetCommonViewFromSelectedParts(drawingParts);
                if (commonView == null) {
                    MessageBox.Show("Zaznaczone elementy muszą należeć do jednego widoku.");
                    return null;
                }
            }

            var outlineSnapshots = GetPartOutlineSnapshots(commonView, drawingParts);
            if (outlineSnapshots == null || outlineSnapshots.Count == 0) {
                MessageBox.Show("Nie udało się wyznaczyć krawędzi dla zaznaczonych elementów.");
                return null;
            }

            DrawOutlineSnapshots(commonView, outlineSnapshots);

            if (!HasExactlyOnePart(commonView)) {
                drawing.CommitChanges();
                return string.Empty;
            }

            var edgesByNumber = BuildEdgesByNumberFromOutlines(outlineSnapshots);
            if (edgesByNumber.Count == 0) {
                drawing.CommitChanges();
                return string.Empty;
            }

            var numberedGroups = BuildNumberedEdgeGroups(
                edgesByNumber,
                JoinToleranceMillimeters,
                NearStraightAngleDegrees
            );

            numberedGroups = FilterShortNumberedEdgeGroups(
                numberedGroups,
                MinimumNumberedEdgeLengthMillimeters
            );

            DrawEdgeNumbers(commonView, numberedGroups);

            drawing.CommitChanges();

            return FormatEdgeNumbersForTextBox(
                numberedGroups.Keys.OrderBy(number => number).ToList()
            );
        }

        private static void DrawOutlineSnapshots(
            TSD.View view,
            List<PartOutlineSnapshot> outlineSnapshots
        ) {
            if (view == null || outlineSnapshots == null || outlineSnapshots.Count == 0) return;

            foreach (var group in outlineSnapshots
                         .Where(snapshot => snapshot?.SegmentGroups != null && snapshot.SegmentGroups.Count != 0)
                         .SelectMany(snapshot =>
                             snapshot.SegmentGroups.Where(currentGroup =>
                                 currentGroup?.Points != null && currentGroup.Points.Count >= 2)))
                if (!group.IsPolyline || group.Points.Count == 2)
                    DrawStraightSegmentPrimitive(view, group.Points[0], group.Points[1], TSD.DrawingColors.Red);
                else
                    DrawPolylinePrimitive(view, group.Points, TSD.DrawingColors.Red);
        }

        #endregion

        #region Drawing And Numbering

        private static void DrawEdgeNumbers(
            TSD.View view,
            Dictionary<int, NumberedEdgeGroup> numberedGroups
        ) {
            if (view == null || numberedGroups == null || numberedGroups.Count == 0) return;

            foreach (var pair in numberedGroups.OrderBy(group => group.Key)) {
                var group = pair.Value;
                if (group == null) continue;

                TSG.Point numberPoint;

                if (!group.IsPolyline) {
                    var singleSegment = group.EdgeSegments[0];
                    numberPoint = ComputeTextInsertionPointForSegment(
                        singleSegment.StartPoint,
                        singleSegment.EndPoint,
                        NumberOffsetMillimeters
                    );
                }
                else
                    numberPoint = ComputeTextInsertionPointForPolyline(
                        group.PolylinePoints,
                        NumberOffsetMillimeters
                    );

                new TSD.Text(view, numberPoint, group.GroupNumber.ToString()).Insert();
            }
        }

        private static string FormatEdgeNumbersForTextBox(List<int> numbers) {
            if (numbers == null || numbers.Count == 0) return string.Empty;

            var orderedNumbers = numbers
                .Distinct()
                .OrderBy(number => number)
                .ToList();

            if (orderedNumbers.Count <= MaximumInlineEdgeCount)
                return string.Join(",", orderedNumbers);

            return orderedNumbers.First() + "-" + orderedNumbers.Last();
        }

        private static TSG.Point ComputeTextInsertionPointForSegment(
            TSG.Point startPoint,
            TSG.Point endPoint,
            double offsetMillimeters
        ) {
            var midpoint = new TSG.Point(
                (startPoint.X + endPoint.X) * 0.5,
                (startPoint.Y + endPoint.Y) * 0.5,
                0
            );

            var directionX = endPoint.X - startPoint.X;
            var directionY = endPoint.Y - startPoint.Y;

            var directionLength = Math.Sqrt(directionX * directionX + directionY * directionY);
            if (directionLength < 1e-9) return new TSG.Point(midpoint.X, midpoint.Y + offsetMillimeters, 0);

            var normalX = -directionY / directionLength;
            var normalY = directionX / directionLength;

            return new TSG.Point(
                midpoint.X + normalX * offsetMillimeters,
                midpoint.Y + normalY * offsetMillimeters,
                0
            );
        }

        private static TSG.Point ComputeTextInsertionPointForPolyline(
            List<TSG.Point> polylinePoints,
            double offsetMillimeters
        ) {
            if (polylinePoints == null || polylinePoints.Count < 2) return new TSG.Point(0, 0, 0);

            var sumX = 0.0;
            var sumY = 0.0;

            foreach (var polylinePoint in polylinePoints) {
                sumX += polylinePoint.X;
                sumY += polylinePoint.Y;
            }

            var averageX = sumX / polylinePoints.Count;
            var averageY = sumY / polylinePoints.Count;

            var firstPoint = polylinePoints[0];
            var lastPoint = polylinePoints[polylinePoints.Count - 1];

            var directionX = lastPoint.X - firstPoint.X;
            var directionY = lastPoint.Y - firstPoint.Y;

            var directionLength = Math.Sqrt(directionX * directionX + directionY * directionY);
            if (directionLength < 1e-9) return new TSG.Point(averageX, averageY + offsetMillimeters, 0);

            var normalX = -directionY / directionLength;
            var normalY = directionX / directionLength;

            return new TSG.Point(
                averageX + normalX * offsetMillimeters,
                averageY + normalY * offsetMillimeters,
                0
            );
        }

        private static void DrawStraightSegmentPrimitive(
            TSD.View view,
            TSG.Point startPoint,
            TSG.Point endPoint,
            TSD.DrawingColors color
        ) {
            if (view == null || startPoint == null || endPoint == null) return;

            var distance = Math.Sqrt(
                (startPoint.X - endPoint.X) * (startPoint.X - endPoint.X) +
                (startPoint.Y - endPoint.Y) * (startPoint.Y - endPoint.Y)
            );

            if (distance <= DuplicateToleranceMillimeters) return;

            var points = new TSD.PointList {
                new TSG.Point(startPoint.X, startPoint.Y, 0),
                new TSG.Point(endPoint.X, endPoint.Y, 0)
            };

            var polyline = new TSD.Polyline(view, points);
            polyline.Attributes.Line.Color = color;
            polyline.Insert();
        }

        private static void DrawPolylinePrimitive(
            TSD.View view,
            List<TSG.Point> points,
            TSD.DrawingColors color
        ) {
            if (view == null || points == null || points.Count < 2) return;

            var cleanedPoints = RemoveNearDuplicates(
                new List<TSG.Point>(points),
                DuplicateToleranceMillimeters
            );

            if (cleanedPoints == null || cleanedPoints.Count < 2) return;

            switch (cleanedPoints.Count) {
                case 2:
                    DrawStraightSegmentPrimitive(view, cleanedPoints[0], cleanedPoints[1], color);
                    return;
                case 3: {
                    var first = cleanedPoints[0];
                    var last = cleanedPoints[cleanedPoints.Count - 1];

                    var distance = Math.Sqrt(
                        (first.X - last.X) * (first.X - last.X) +
                        (first.Y - last.Y) * (first.Y - last.Y)
                    );

                    if (distance <= DuplicateToleranceMillimeters) {
                        DrawStraightSegmentPrimitive(view, cleanedPoints[0], cleanedPoints[1], color);
                        return;
                    }

                    break;
                }
            }

            var polylinePoints = new TSD.PointList();
            foreach (var point in cleanedPoints)
                polylinePoints.Add(new TSG.Point(point.X, point.Y, 0));

            var polyline = new TSD.Polyline(view, polylinePoints);
            polyline.Attributes.Line.Color = color;
            polyline.Insert();
        }

        #endregion
    }
}