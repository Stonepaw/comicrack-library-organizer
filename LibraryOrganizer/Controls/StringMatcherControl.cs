using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class StringMatcherControl : UserControl
    {
        private static readonly IReadOnlyCollection<
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

        private readonly BookFieldStringMatcherViewModel _viewModel;

        public new int Width
        {
            get => base.Width;
            set
            {
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

            SuspendLayout();

            actionsButton.Size = MatcherControl.ActionButtonSize;

            tableLayoutPanel1.ColumnStyles[1] = new ColumnStyle(
                SizeType.Absolute,
                MatcherControl.FieldComboBoxWidth
            );
            tableLayoutPanel1.ColumnStyles[2] = new ColumnStyle(
                SizeType.Absolute,
                MatcherControl.ModeComboBoxWidth
            );

            tableLayoutPanel1.ResumeLayout();

            bookFieldStringMatcherViewModelBindingSource.DataSource = _viewModel;

            field.DataSource = ComicBookField.ComicBookField.Fields;
            field.DisplayMember = "Label";
            mode.DataSource = StringOperators;
            mode.DisplayMember = "Value";
            mode.ValueMember = "Key";

            ResumeLayout();
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
