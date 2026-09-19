using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace LibraryOrganizer.Data
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected void Set<T>(ref T field, T value, [CallerMemberName] string caller = "")
        {
            if (!EqualityComparer<T>.Default.Equals(field, value))
            {
                field = value;
                NotifyPropertyChanged(caller);
            }
        }
    }
}
