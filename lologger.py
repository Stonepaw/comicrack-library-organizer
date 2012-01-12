"""
lologger.py

Contains a class for logging what the bookmover does.

1.7.12: Fix for unicode characters

Copyright Stonepaw
"""

import clr
import System

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult

class Logger():
    
    def __init__(self):
        self._log = []
        self._profile = ""
        self.header = ""

    def Add(self, action, path, message = "", profile=""):
        if profile:
            self._log.append([profile, unicode(action), unicode(path), unicode(message)])
        else:
            self._log.append([self._profile, unicode(action), unicode(path), unicode(message)])


    def SetProfile(self, profile):
        self._profile = profile


    def ToArray(self):
        return self._log

    def Clear(self):
        del(self._log[:])

    def SaveLog(self):
        try:
            save = SaveFileDialog()
            save.AddExtension = True
            save.Filter = "Text files (*.txt)|*.txt"
            if save.ShowDialog() == DialogResult.OK:
                with open(save.FileName, 'w') as f:
                    f.write("Library Organizer Report:\n\n" + self.header)
                    for array in self._log:
                        f.write("\n\n%s:\n%s: %s" % (array[0], array[1], array[2]))
                        if array[3] != "":
                            f.write("\nMessage: " + array[3])
            save.Dispose()
        except Exception, ex:
            MessageBox.Show("something went wrong saving the file. The error was: " + str(ex))

    def add_header(self, header_text):
        self.header += header_text