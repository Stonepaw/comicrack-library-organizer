using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.ViewModel
{
    internal class MatcherViewModel : ViewModelBase, IMatcherViewModel
    {
        private readonly IGroupMatcherViewModel _parent;

        private bool _negated;

        private IComicBookField _field = ComicBookField.ComicBookField.AlternateCount;

        private IMatcherValueViewModel _value = new NumberMatcherValueViewModel();

        public bool Negated
        {
            get => _negated;
            set => Set(ref _negated, value);
        }

        public IComicBookField Field
        {
            get => _field;
            set
            {
                _field = value;

                if (ChangeValueIfRequired(value))
                {
                    NotifyPropertyChanged();
                    NotifyPropertyChanged(nameof(Value));
                }
                else
                {
                    NotifyPropertyChanged();
                }
            }
        }

        public IMatcherValueViewModel Value
        {
            get => _value;
            private set => Set(ref _value, value);
        }

        public MatcherViewModel(IGroupMatcherViewModel parent)
        {
            this._parent = parent;
        }

        public void Delete()
        {
            _parent.Remove(this);
        }

        private bool ChangeValueIfRequired(IComicBookField field)
        {
            switch (field)
            {
                case ComicBookIntField _ when !(_value is NumberMatcherValueViewModel):
                    _value = new NumberMatcherValueViewModel();
                    return true;
                case ComicBookStringField _ when !(_value is StringMatcherValueViewModel):
                    _value = new StringMatcherValueViewModel();
                    return true;
                case ComicBookYesNoField _ when !(_value is YesNoMatcherValueViewModel):
                    _value = new YesNoMatcherValueViewModel();
                    return true;
            }

            return false;
        }
    }
}
