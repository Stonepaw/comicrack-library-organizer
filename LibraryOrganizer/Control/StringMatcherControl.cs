using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Control
{
    internal partial class StringMatcherControl : UserControl
    {
        private static IReadOnlyCollection<
            KeyValuePair<StringMatcherMode, string>
        > StringOperators = Enum.GetValues(typeof(StringMatcherMode))
            .Cast<StringMatcherMode>()
            .Select(x => new KeyValuePair<StringMatcherMode, string>(
                x,
                (
                    Attribute.GetCustomAttribute(
                        x.GetType().GetField(x.ToString()),
                        typeof(DescriptionAttribute)
                    ) as DescriptionAttribute
                )?.Description
                    ?? x.ToString()
            ))
            .ToList();

        private BookFieldStringMatcherViewModel _viewModel;

        public new int Width
        {
            get => base.Width;
            set
            {
                Debug.WriteLine($"Setting matcher width to {value}");

                base.Width = value;
                MaximumSize = new Size(value, int.MaxValue);
                MinimumSize = new Size(value, 0);
            }
        }

        public StringMatcherControl()
            : this(
                new BookFieldStringMatcherViewModel(
                    new GroupMatcherViewModel(null),
                    ComicBookStringField.AgeRating
                )
            ) { }

        public StringMatcherControl(BookFieldStringMatcherViewModel viewModel)
        {
            _viewModel = viewModel;

            InitializeComponent();

            bookFieldStringMatcherViewModelBindingSource.DataSource = _viewModel;

            field.DataSource = ComicBookField.ComicBookField.Fields;
            field.DisplayMember = "Label";
            mode.DataSource = StringOperators;
            mode.DisplayMember = "Value";
            mode.ValueMember = "Key";
        }

        private void actionsButton_Click(object sender, System.EventArgs e)
        {
            matcherGroupActions.Show(actionsButton, new Point(0, actionsButton.Height));
        }

        private void deleteGroupAction_Click(object sender, System.EventArgs e)
        {
            _viewModel.Delete();
        }
    }
}
