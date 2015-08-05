from ComicRack import ComicRack
from System.IO import Path
import i18n
i18n.resourcespath = ".\\resources\\"
import duplicateform
duplicateform.ICONDIRECTORY = Path.Combine(Path.GetDirectoryName(__file__), "GUI")
from duplicateform import DuplicateForm
i18n.setup(ComicRack())
d = DuplicateForm()
d.ShowDialog()