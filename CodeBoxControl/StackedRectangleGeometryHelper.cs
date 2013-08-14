using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Diagnostics;
using System.Windows;
namespace CodeBoxControl
{
  public  class StackedRectangleGeometryHelper
    {
      private Geometry mOrginalGeometry;
      public StackedRectangleGeometryHelper(Geometry geom)
      {
          mOrginalGeometry = geom;
      }

       
      /// <summary>
      /// Creates List of geometries for underline
      /// </summary>
      /// <returns>List of Geometries</returns>
      public List<Geometry> BottomEdgeRectangleGeometries()
      {
          List<Geometry> geoms = new List<Geometry>();
          PathGeometry pg = (PathGeometry)mOrginalGeometry;
          foreach (PathFigure fg in pg.Figures)
          {
              PolyLineSegment pls = (PolyLineSegment)fg.Segments[0]; 
              PointCollectionHelper pch = new PointCollectionHelper(pls.Points, fg.StartPoint);
              List<double> distinctY = pch.DistinctY;
              for (int i = 0; i < distinctY.Count - 1; i++)
              {
                  double bottom = distinctY[i + 1] - 3;
                  double top = bottom + 2;
                  // ordered values of X that are present for both Y values
                  List<double> rlMatches = pch.XAtY(distinctY[i], distinctY[i + 1]); 
                  double left = rlMatches[0];
                  double right = rlMatches[rlMatches.Count - 1];

                  PathGeometry rpg = CreateGeometry(top, bottom, left, right);
                  geoms.Add(rpg);
              }
          }
          return geoms;
      }

      /// <summary>
      /// Creates List of geometries for Strikethru
      /// </summary>
      /// <returns>List of Geometries</returns>
      public List<Geometry> CenterLineRectangleGeometries()
      {
          List<Geometry> geoms = new List<Geometry>();
          PathGeometry pg = (PathGeometry)mOrginalGeometry;

          foreach (PathFigure fg in pg.Figures)
          {
              PolyLineSegment pls = (PolyLineSegment)fg.Segments[0];
              PointCollectionHelper pch = new PointCollectionHelper(pls.Points, fg.StartPoint);
              List<double> distinctY = pch.DistinctY;
              for (int i = 0; i < distinctY.Count - 1; i++)
              {
                  double top = (distinctY[i] + distinctY[i + 1])/ 2 + 1;
                  double bottom = (distinctY[i] + distinctY[i + 1]) / 2 - 1;
                  // ordered values of X that are present for both Y values
                  List<double> rlMatches = pch.XAtY(distinctY[i], distinctY[i + 1]);
                  double left = rlMatches[0];
                  double right = rlMatches[rlMatches.Count - 1];

                  PathGeometry rpg = CreateGeometry(top, bottom, left, right);
                  geoms.Add(rpg);
              }
          }
          return geoms;
      }

      private static PathGeometry CreateGeometry(double top, double bottom, double left, double right)
      {
          PolyLineSegment plr = new PolyLineSegment();

          plr.Points.Add(new Point(right, top));
          plr.Points.Add(new Point(right, bottom));
          plr.Points.Add(new Point(left, bottom));
          PathFigure rpf = new PathFigure();
          rpf.StartPoint = new Point(left, top);
          rpf.Segments.Add(plr);
          rpf.IsClosed = true;
          PathGeometry rpg = new PathGeometry();
          rpg.Figures.Add(rpf);
          return rpg;
      }

    }
}
