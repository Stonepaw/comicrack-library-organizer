import pyevent
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Input import ICommand, CommandManager

"""notify_property and NotifyPropertyChangedBase are from:
http://gui-at.blogspot.ca/2009/11/inotifypropertychanged-in-ironpython.html
"""
class notify_property(property):

    def __init__(self, getter):
        def newgetter(slf):
            #return None when the property does not exist yet
            try:
                return getter(slf)
            except AttributeError:
                return None
        super(notify_property, self).__init__(newgetter)

    def setter(self, setter):
        def newsetter(slf, newvalue):
            # do not change value if the new value is the same
            # trigger PropertyChanged event when value changes
            oldvalue = self.fget(slf)
            if oldvalue != newvalue:
                setter(slf, newvalue)
                slf.OnPropertyChanged(setter.__name__)
        return property(
            fget=self.fget,
            fset=newsetter,
            fdel=self.fdel,
            doc=self.__doc__)


class NotifyPropertyChangedBase(INotifyPropertyChanged):
    PropertyChanged = None

    def __init__(self):
        self.PropertyChanged, self._propertyChangedCaller = pyevent.make_event()

    def add_PropertyChanged(self, value):
        self.PropertyChanged += value

    def remove_PropertyChanged(self, value):
        self.PropertyChanged -= value

    def OnPropertyChanged(self, propertyName):
        if self.PropertyChanged is not None:
            self._propertyChangedCaller(self, PropertyChangedEventArgs(propertyName))


class Command(ICommand):
    """A helper class to easily create and use wpf ICommands. Supports Parameters and CanExecute
    
    Originally from: http://mark-dot-net.blogspot.ca/2010/10/wpf-and-mvvm-in-ironpython.html
    """    
    def __init__(self, execute, can_execute=None, uses_parameter=False):
        """Initiates a new command.

        Args:
            execute: a fuction to execute when the Command is executed.
            can_execute: a function to call when CanExecute is called.
                by default can_execute will return true.
            uses_parameter: A bool if the command should expect a parameter or not.
                Note that sometime in menus the parameter in CanExecute will be
                None even when it should not be. Make sure the can_execute function
                handles this.
        """
        self._execute = execute
        self._canexecute = can_execute
        self._uses_parameter = uses_parameter
        self._handlers = []

    def Execute(self, parameter):
        """Executes the function passed into the command when it was initiated"""
        if self._uses_parameter:
            self._execute(parameter)
        else:
            self._execute()
         
    def add_CanExecuteChanged(self, handler):
        #Using CommandManager.RequerySuggested seems to fix a bug in WPF menus
        #Where the CanExecute parameter is not passed
        CommandManager.RequerySuggested += handler 
     
    def remove_CanExecuteChanged(self, handler):
        CommandManager.RequerySuggested -= handler

    def canExecuteChanged(self): 
        for handler in self._handlers: 
            handler(self, EventArgs.Empty)
 
    def CanExecute(self, parameter):
        if self._canexecute is None:
            return True
        #There is a possiblity that the parameter will be none even when it's not supposed to be.
        #So instead of checking for None, use a try except.
        if self._uses_parameter:
            try:
                return self._canexecute(parameter)
            except TypeError:
                return self._canexecute()
        else:
            return self._canexecute()            
            