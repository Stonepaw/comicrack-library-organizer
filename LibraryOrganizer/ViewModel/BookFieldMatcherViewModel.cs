using LibraryOrganizer.ComicBookField;
using LibraryOrganizer.Matcher;

namespace LibraryOrganizer.ViewModel
{
    internal class BookFieldMatcherViewModel : ViewModelBase, IMatcherViewModel
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

        public BookFieldMatcherViewModel(IGroupMatcherViewModel parent)
        {
            this._parent = parent;
        }

        public BookFieldMatcherViewModel(IGroupMatcherViewModel parent, IBookFieldMatcher matcher)
            : this(parent)
        {
            _negated = matcher.Negated;

            switch (matcher)
            {
                case IBookFieldIntMatcher bookFieldIntMatcher:
                {
                    _field = bookFieldIntMatcher.Field;
                    _value = new NumberMatcherValueViewModel(bookFieldIntMatcher);
                    break;
                }
                case IBookFieldMangaYesNoMatcher bookFieldMangaYesNoMatcher:
                {
                    _field = bookFieldMangaYesNoMatcher.Field;
                    _value = new MangaYesNoMatcherValueViewModel(bookFieldMangaYesNoMatcher);
                    break;
                }
                case IBookFieldStringMatcher bookFieldStringMatcher:
                {
                    _field = bookFieldStringMatcher.Field;
                    _value = new StringMatcherValueViewModel(bookFieldStringMatcher);
                    break;
                }
                case IBookFieldYesNoMatcher bookFieldYesNoMatcher:
                {
                    _field = bookFieldYesNoMatcher.Field;
                    _value = new YesNoMatcherValueViewModel(bookFieldYesNoMatcher);
                    break;
                }
            }
        }

        public void AddBookFieldMatcherAfter()
        {
            _parent.AddBookFieldMatcherAfter(this);
        }

        public void AddGroupAfter()
        {
            _parent.AddGroupAfter(this);
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
                case ComicBookMangaYesNoField _ when !(_value is MangaYesNoMatcherValueViewModel):
                    _value = new MangaYesNoMatcherValueViewModel();
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
