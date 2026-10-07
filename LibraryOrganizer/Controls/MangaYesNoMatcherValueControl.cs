using System.Windows.Forms;
using LibraryOrganizer.ViewModel;

namespace LibraryOrganizer.Controls
{
    internal partial class MangaYesNoMatcherValueControl : UserControl
    {
        private readonly MangaYesNoMatcherValueViewModel _viewModel;

        public MangaYesNoMatcherValueControl()
            : this(new MangaYesNoMatcherValueViewModel()) { }

        public MangaYesNoMatcherValueControl(MangaYesNoMatcherValueViewModel viewModel)
        {
            _viewModel = viewModel;
            InitializeComponent();

            mangaYesNoMatcherValueViewModelBindingSource.DataSource = _viewModel;
            comboBox1.DataSource = MangaYesNoMatcherValueViewModel.Options;
            comboBox1.ValueMember = "Key";
            comboBox1.DisplayMember = "Value";
        }
    }
}
