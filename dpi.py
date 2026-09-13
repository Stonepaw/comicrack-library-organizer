import clr

import System

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

# https://stackoverflow.com/questions/22735174/how-to-write-winforms-code-that-auto-scales-to-system-font-and-dpi-settings
def apply_dpi_container_scaling(container):
    container.AutoScaleDimensions = System.Drawing.Size(96, 96)
    container.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
