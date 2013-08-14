import System

import clr
clr.AddReference("IronPython.Wpf")
clr.AddReference("PresentationCore")

from IronPython.Modules import Wpf

class Test(Object):
    def __init__(self):
        Wpf.LoadComponent("Test.xaml")
