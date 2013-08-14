using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
namespace CodeBoxControl
{
  public class PointCollectionHelper
    {
      private PointCollection mPointCollection;
      private Point mStartPoint;
      private Double[] mYValues;
      private Double[] mXValues;
      private bool mAreXValuesGenerated;
      private bool mAreYValuesGenerated;
      private Dictionary<double, List<Double>> mXRanges = new Dictionary<double, List<Double>>();
      private Dictionary<double, List<Double>> mYRanges = new Dictionary<double, List<Double>>();
      private bool mAreXRangesGenerated;
      private bool mAreYRangesGenerated;

      public PointCollectionHelper(PointCollection pointCollection, Point StartPoint)
      {
          mPointCollection = pointCollection;
          mStartPoint = StartPoint;
      }
      private void EnsureXValues()
      {
          if (!mAreXValuesGenerated)
          {
              HashSet<double> distinct = new HashSet<double>();
              foreach (Point p in mPointCollection)
              {
                  distinct.Add(p.X);
              }
              if (mStartPoint != null)
              {
                  distinct.Add(mStartPoint.X);
              }
              mXValues = new Double[distinct.Count];
              distinct.CopyTo(mXValues);
              Array.Sort(mXValues);
              mAreXValuesGenerated = true;
          }

      }

      private void EnsureYValues()
      {
          if (!mAreYValuesGenerated)
          {
              HashSet<double> distinct = new HashSet<double>();
              foreach (Point p in mPointCollection)
              {
                  distinct.Add(p.Y);
              }
              if (mStartPoint != null)
              {
                  distinct.Add(mStartPoint.Y );
              }
              mYValues = new Double[distinct.Count];
              distinct.CopyTo(mYValues);
              Array.Sort(mYValues);
              mAreYValuesGenerated = true;
          }

      }

      private void EnsureXRanges()
      {
          if (!mAreXRangesGenerated)
          {
              EnsureXValues();
              foreach (double d in mXValues)
              {
                  mXRanges.Add(d, YAtXInternal(d));
              }
              mAreXRangesGenerated = true;
          }

      }

      private List<double> YAtXInternal(double x)
      {
          List<double> yVals = new List<double>();
          foreach (Point p in mPointCollection)
          {
              if (p.X == x)
              {
                  yVals.Add(p.Y);
              }
          }
          if (mStartPoint != null)
          {
              if (mStartPoint.X == x)
              {
                  if (!yVals.Contains(mStartPoint.Y))
                  {
                      yVals.Add(mStartPoint.Y);
                  }
              }
          }
          yVals.Sort();
          return yVals;

      }

      private void EnsureYRanges()
      {
          if (!mAreYRangesGenerated)
          {
              EnsureYValues();
              foreach (double d in mYValues)
              {
                  mYRanges.Add(d, XAtYInternal(d));
              }
   
              mAreYRangesGenerated = true;
          }
      }

      private List<double> XAtYInternal(Double y)
      {
          List<double> xVals = new List<double>();
          foreach (Point p in mPointCollection)
          {
              if (p.Y == y)
              {
                  if (!xVals.Contains(p.X))
                  {
                      xVals.Add(p.X);
                  }
              }
          }
          if (mStartPoint != null)
          {
              if (mStartPoint.Y == y)
              {
                  if (!xVals.Contains(mStartPoint.X))
                  {
                      xVals.Add(mStartPoint.X);
                  }
              }
          }
          xVals.Sort();
          return xVals;
      }
 
 

      public List<double> YAtX(double x)
      {
          List<double> yVals = new List<double>();
          foreach (Point p in mPointCollection)
          {
              if (p.X == x)
              {
                  yVals.Add(p.Y);
              }
          }
          if (mStartPoint != null)
          {
              if (mStartPoint.X == x)
              {
                  yVals.Add(mStartPoint.Y);
              }
          }
          yVals.Sort();
          return yVals;

      }

      public List<double> XAtY(Double y)
      {
          EnsureYRanges();
          return mYRanges[y];
           
      }


      public List<double> XAtY(double y1, double y2)
      {
          EnsureYRanges();
          HashSet<double> y1Hash = new HashSet<double>();
          foreach (double d in XAtY(y1))
          {
              y1Hash.Add(d);
          }

         IEnumerable<double> matches =  y1Hash.Intersect(XAtY(y2));
         List<Double> matchList = new List<double>();
         matchList.AddRange(matches);
        return matchList;
      }

 




      public double MaxX 
      {
          get
          {
              if (mPointCollection.Count == 0) { return Double.NaN; }
              EnsureXValues();
              return mXValues[mXValues.Length - 1];
          }
      }

      public   double MaxY 
      {
          get
          {
              if (mPointCollection.Count == 0) { return Double.NaN; }
              EnsureYValues();
              return mXValues[mYValues.Length - 1];
          }

      }

      public   double MinX  
      {
          get
          {
              EnsureXValues();
              return mXValues[0];
          }
      }

      public  double MinY 
      {
          get
          {
              EnsureXValues();
              return mYValues[0];
          }


      }

      public   double Height 
      {
          get
          {
              return MaxY - MinY;
          }
      }
      public   double Width 
      {
          get
          {
              return MaxX - MinX;
          }
      }



  
      public   List<Double> DistinctX 
      {
          get
          {
              EnsureXValues();
              List<Double> distinctX = new List<Double>();
              distinctX.AddRange(mXValues);
              return distinctX;
          }
      }
  
      public   List<Double> DistinctY 
      {
          get
          {
              EnsureYValues();
              List<Double> distinctY = new List<Double>();
              distinctY.AddRange(mYValues);
              return distinctY;
          }
      }

      public   Rect BoundingRectangle(PointCollection pc)
      {
          Rect rect = new Rect(MinX , MinY , Width , Height );
          return rect;
      }


     

    }
}
