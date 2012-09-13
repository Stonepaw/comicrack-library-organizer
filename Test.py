import wpf

from System.Windows import Window

class Test(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'Test.xaml')
