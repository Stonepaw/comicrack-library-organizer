"""
lologger.py

Contains a class for logging what the bookmover does.


Copyright 2010-2012 Stonepaw

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

import clr
import System

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult  # @UnresolvedImport

class Logger():
    
    def __init__(self):
        self._log = []
        self._current_profile = ""
        self.header = ""
        #TODO Store book in this instead of passing the path in ADD

    def Add(self, action, path, message = "", profile=""):
        if profile:
            self._log.append([profile, unicode(action), unicode(path), unicode(message)])
        else:
            self._log.append([self._current_profile, unicode(action), unicode(path), unicode(message)])


    def SetProfile(self, profile):
        self._current_profile = profile


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