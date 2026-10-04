using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class NumberMatcherValueControl : UserControl
    {
        private readonly NumberMatcherValueViewModel _viewModel;

        public NumberMatcherValueControl()
            : this(new NumberMatcherValueViewModel()) { }

        public NumberMatcherValueControl(NumberMatcherValueViewModel viewModel)
        {
            _viewModel = viewModel;
            InitializeComponent();
            numberMatcherValueViewModelBindingSource.DataSource = _viewModel;

            mode.DataSource = NumberMatcherValueViewModel.Modes;
            mode.ValueMember = "Key";
            mode.DisplayMember = "Value";
            ToggleValue2();

            _viewModel.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(_viewModel.Value2Enabled))
                {
                    ToggleValue2();
                }
            };
        }

        private void ToggleValue2()
        {
            SuspendLayout();
            valueTableLayoutPanel.SuspendLayout();

            if (_viewModel.Value2Enabled)
            {
                valueTableLayoutPanel.ColumnStyles[0] = new ColumnStyle(SizeType.Percent, 50);
                valueTableLayoutPanel.ColumnStyles[1] = new ColumnStyle(SizeType.Percent, 50);
                value.Margin = new Padding(0, 0, 3, 0);
                value2.Show();
            }
            else
            {
                valueTableLayoutPanel.ColumnStyles[0] = new ColumnStyle(SizeType.Percent, 100);
                valueTableLayoutPanel.ColumnStyles[1] = new ColumnStyle(SizeType.Absolute, 0);
                value.Margin = new Padding(0);
                value2.Hide();
            }

            valueTableLayoutPanel.ResumeLayout();
            ResumeLayout();
        }
    }
}
