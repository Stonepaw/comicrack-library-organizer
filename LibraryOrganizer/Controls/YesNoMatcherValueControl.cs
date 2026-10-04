using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class YesNoMatcherValueControl : UserControl
    {
        private readonly YesNoMatcherValueViewModel _viewModel;

        public YesNoMatcherValueControl()
            : this(new YesNoMatcherValueViewModel()) { }

        public YesNoMatcherValueControl(YesNoMatcherValueViewModel viewModel)
        {
            _viewModel = viewModel;
            InitializeComponent();

            yesNoMatcherValueViewModelBindingSource.DataSource = _viewModel;
            comboBox1.DataSource = YesNoMatcherValueViewModel.Options;
            comboBox1.ValueMember = "Key";
            comboBox1.DisplayMember = "Value";
        }
    }
}
