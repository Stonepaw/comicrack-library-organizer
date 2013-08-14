using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using CodeBoxControl;
using CodeBoxControl.Decorations;
using System.Windows.Media;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Stonepaw.LibraryOrganizer.Controls
{
    public class TemplateTextBox : CodeBox
    {
        
        
        public static readonly DependencyProperty BindableSelectionStartProperty =
            DependencyProperty.Register(
            "BindableSelectionStart",
            typeof(int),
            typeof(TemplateTextBox),
            new PropertyMetadata(OnBindableSelectionStartChanged));

        public static readonly DependencyProperty BindableSelectionLengthProperty =
            DependencyProperty.Register(
            "BindableSelectionLength",
            typeof(int),
            typeof(TemplateTextBox),
            new PropertyMetadata(OnBindableSelectionLengthChanged));

        private bool changeFromUI;

        public TemplateTextBox()
            : base()
        {
            this.SelectionChanged += this.OnSelectionChanged;
            RegexDecoration fields = new RegexDecoration();
            fields.Brush = new SolidColorBrush(Colors.Blue);
            fields.RegexString = "{.*?<(\\w+).*>.*?}";
            this.Decorations.Add(fields);

            CommonOpenFileDialog c = new CommonOpenFileDialog();
            c.IsFolderPicker = true;
            

        }

        public int BindableSelectionStart
        {
            get
            {
                return (int)this.GetValue(BindableSelectionStartProperty);
            }

            set
            {
                this.SetValue(BindableSelectionStartProperty, value);
            }
        }

        public int BindableSelectionLength
        {
            get
            {
                return (int)this.GetValue(BindableSelectionLengthProperty);
            }

            set
            {
                this.SetValue(BindableSelectionLengthProperty, value);
            }
        }

        private static void OnBindableSelectionStartChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs args)
        {
            var textBox = dependencyObject as TemplateTextBox;

            if (!textBox.changeFromUI)
            {
                int newValue = (int)args.NewValue;
                textBox.SelectionStart = newValue;
            }
            else
            {
                textBox.changeFromUI = false;
            }
        }

        private static void OnBindableSelectionLengthChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs args)
        {
            var textBox = dependencyObject as TemplateTextBox;

            if (!textBox.changeFromUI)
            {
                int newValue = (int)args.NewValue;
                textBox.SelectionLength = newValue;
            }
            else
            {
                textBox.changeFromUI = false;
            }
        }

        private void OnSelectionChanged(object sender, RoutedEventArgs e)
        {
            if (this.BindableSelectionStart != this.SelectionStart)
            {
                this.changeFromUI = true;
                this.BindableSelectionStart = this.SelectionStart;
            }

            if (this.BindableSelectionLength != this.SelectionLength)
            {
                this.changeFromUI = true;
                this.BindableSelectionLength = this.SelectionLength;
            }
        }
    }
}
