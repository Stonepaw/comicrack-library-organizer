namespace LibraryOrganizer.ViewModel
{
    internal enum ConfigFormPage
    {
        Overview,
        Files,
        Folders,
        Rules,
        Options,
    }

    internal class ConfigFormViewModel : ViewModelBase
    {
        private ConfigFormPage _currentPage = ConfigFormPage.Overview;

        public bool OverviewPageEnabled
        {
            get => _currentPage == ConfigFormPage.Overview;
        }

        public bool FilesPageEnabled
        {
            get => _currentPage == ConfigFormPage.Files;
        }

        public bool FoldersPageEnabled
        {
            get => _currentPage == ConfigFormPage.Folders;
        }

        public bool RulesPageEnabled
        {
            get => _currentPage == ConfigFormPage.Rules;
        }

        public bool OptionsPageEnabled
        {
            get => _currentPage == ConfigFormPage.Options;
        }

        public void SetPage(ConfigFormPage page)
        {
            if (_currentPage != page)
            {
                var current = _currentPage;
                _currentPage = page;

                NotifyChangeForm(current);
                NotifyChangeForm(page);
            }
        }

        private void NotifyChangeForm(ConfigFormPage page)
        {
            switch (page)
            {
                case ConfigFormPage.Files:
                    NotifyPropertyChanged(nameof(FilesPageEnabled));
                    break;
                case ConfigFormPage.Folders:
                    NotifyPropertyChanged(nameof(FoldersPageEnabled));
                    break;
                case ConfigFormPage.Rules:
                    NotifyPropertyChanged(nameof(RulesPageEnabled));
                    break;
                case ConfigFormPage.Options:
                    NotifyPropertyChanged(nameof(OptionsPageEnabled));
                    break;
                case ConfigFormPage.Overview:
                    NotifyPropertyChanged(nameof(OverviewPageEnabled));
                    break;
            }
        }
    }
}
