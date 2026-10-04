using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class StringMatcherValueControl : UserControl
    {
        private readonly StringMatcherValueViewModel _viewModel;

        public StringMatcherValueControl()
            : this(new StringMatcherValueViewModel()) { }

        public StringMatcherValueControl(StringMatcherValueViewModel viewModel)
        {
            _viewModel = viewModel;

            InitializeComponent();

            stringMatcherValueViewModelBindingSource.DataSource = _viewModel;

            comboBox1.DataSource = StringMatcherValueViewModel.StringOperators;
            comboBox1.DisplayMember = "Value";
            comboBox1.ValueMember = "Key";
        }
    }
}
