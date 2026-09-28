using LibraryOrganizer.ComicBookField;

namespace LibraryOrganizer.ViewModel
{
    internal abstract class BookFieldMatcherViewModel<T> : IBookFieldMatcherViewModel
        where T : IComicBookField
    {
        private readonly IGroupMatcherViewModel _parent;

        private T _field;

        public IComicBookField Field
        {
            get => _field;
            set
            {
                if (value is T expected)
                {
                    _field = expected;
                }
                else
                {
                    _parent.ChangeFieldType(this, value);
                }
            }
        }

        public bool Negated { get; set; } = false;

        protected BookFieldMatcherViewModel(IGroupMatcherViewModel parent, T field)
        {
            _parent = parent;
            _field = field;
        }

        public void Delete()
        {
            _parent.Remove(this);
        }
    }
}
