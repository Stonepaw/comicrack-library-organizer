using System;

namespace LibraryOrganizer.ViewModel
{
    public enum ConfigFormPage
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

        public bool OverviewPageEnabled => _currentPage == ConfigFormPage.Overview;

        public bool FilesPageEnabled => _currentPage == ConfigFormPage.Files;

        public bool FoldersPageEnabled => _currentPage == ConfigFormPage.Folders;

        public bool RulesPageEnabled => _currentPage == ConfigFormPage.Rules;

        public bool OptionsPageEnabled => _currentPage == ConfigFormPage.Options;

        public ConfigFormPage CurrentPage
        {
            get => _currentPage;
            set => SetPage(value);
        }

        public void SetPage(ConfigFormPage page)
        {
            if (_currentPage == page)
            {
                return;
            }

            ConfigFormPage current = _currentPage;
            _currentPage = page;

            NotifyPropertyChanged(nameof(CurrentPage));

            NotifyChangedPage(current);
            NotifyChangedPage(page);
        }

        private void NotifyChangedPage(ConfigFormPage page)
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
                default:
                    throw new ArgumentOutOfRangeException(nameof(page), page, null);
            }
        }
    }
}
