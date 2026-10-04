using System.Drawing;
using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class MatcherControl : UserControl
    {
        private readonly MatcherViewModel _viewModel;

        public new int Width
        {
            get => base.Width;
            set
            {
                SuspendLayout();
                base.Width = value;
                MaximumSize = new Size(value, int.MaxValue);
                MinimumSize = new Size(value, 0);
                ResumeLayout();
            }
        }

        public MatcherControl()
            : this(new MatcherViewModel(new GroupMatcherViewModel(null))) { }

        public MatcherControl(MatcherViewModel viewModel)
        {
            _viewModel = viewModel;

            InitializeComponent();

            SuspendLayout();

            matcherViewModelBindingSource.DataSource = _viewModel;
            field.DataSource = ComicBookField.ComicBookField.Fields;
            field.DisplayMember = "Label";

            _viewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(_viewModel.Value))
                {
                    UpdateValueControl();
                }
            };
            UpdateValueControl();

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

        private void UpdateValueControl()
        {
            if (_viewModel.Value is StringMatcherValueViewModel stringMatcherValueViewModel)
            {
                UseStringMatcherValueControl(stringMatcherValueViewModel);
            }
            else
            {
                matcherValuePanel.SuspendLayout();
                matcherValuePanel.Controls.Clear();
                matcherValuePanel.ResumeLayout();
            }
        }

        private void UseStringMatcherValueControl(StringMatcherValueViewModel viewModel)
        {
            if (matcherValuePanel.Controls.Count > 0)
            {
                if (matcherValuePanel.Controls[0] is StringMatcherValueControl)
                {
                    return;
                }
            }

            matcherValuePanel.SuspendLayout();
            matcherValuePanel.Controls.Clear();

            StringMatcherValueControl stringMatcherValueControl = new StringMatcherValueControl(
                viewModel
            );
            stringMatcherValueControl.SuspendLayout();
            stringMatcherValueControl.Dock = DockStyle.Fill;
            matcherValuePanel.Controls.Add(stringMatcherValueControl);
            stringMatcherValueControl.ResumeLayout();
            matcherValuePanel.ResumeLayout();
        }
    }
}
