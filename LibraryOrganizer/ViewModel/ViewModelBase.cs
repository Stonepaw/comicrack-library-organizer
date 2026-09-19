using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace LibraryOrganizer.Data
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool Set<T>(ref T field, T value, [CallerMemberName] string caller = "")
        {
            if (!EqualityComparer<T>.Default.Equals(field, value))
            {
                field = value;
                NotifyPropertyChanged(caller);
                return true;
            }

            return false;
        }

        protected bool SetProperty<T>(
            T oldValue,
            T newValue,
            Action<T> setter,
            [CallerMemberName] string caller = ""
        )
        {
            if (!EqualityComparer<T>.Default.Equals(oldValue, newValue))
            {
                setter(newValue);
                NotifyPropertyChanged(caller);
                return true;
            }
            return false;
        }
    }
}
