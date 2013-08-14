using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Windows;
namespace CodeBoxControl.Decorations
{
    /// <summary>
    /// Abstract base class for decorations
    /// </summary>
  public abstract  class Decoration:DependencyObject 
    {
      /// <summary>
      /// EDecoration Type of the decoration , Default is TextColor
      /// </summary>
      public static DependencyProperty DecorationTypeProperty = DependencyProperty.Register("DecorationType", typeof(EDecorationType), typeof(Decoration),
    new PropertyMetadata(EDecorationType.TextColor));

        public EDecorationType DecorationType
        {
            get { return (EDecorationType)GetValue(DecorationTypeProperty); }
            set { SetValue(DecorationTypeProperty, value); }
        }

      /// <summary>
      /// Brushed used for the decoration
      /// </summary>
        public static DependencyProperty BrushProperty = DependencyProperty.Register("Brush", typeof(Brush), typeof(Decoration),
        new PropertyMetadata(null));

        public Brush Brush
        {
            get { return (Brush)GetValue(BrushProperty); }
            set { SetValue(BrushProperty, value); }
        }

        public abstract List<Pair> Ranges(string Text);
    }
}
