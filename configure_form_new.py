import wpf

from System.Windows import Window

class ConfigureFormNew(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'ConfigureFormNew.xaml')
