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

class MoveReporter(object):
    
    def __init__(self):
        self._log = []
        self._current_profile = ""
        self._current_book = ""
        self.header = ""
        #TODO Store book in this instead of passing the path in ADD

    def failed(self, message):
        self._add_log_message(message, "Failed")
    
    def warn(self, message):
        self._add_log_message(message, "Warning")
    
    def success(self, message):
        self._add_log_message(message, "Success")
        
    def skip(self, message):
        self._add_log_message(message, "Skipped")
        
    def success_simulated(self, message):
        self._add_log_message(message, "Success (Simulated)")
    
    def _add_log_message(self, message, action):
        self._log.append(self._current_profile, action, self._current_book, message)

    def Add(self, action, path, message = "", profile=""):
        if profile:
            self._log.append([profile, unicode(action), unicode(path), unicode(message)])
        else:
            self._log.append([self._current_profile, unicode(action), unicode(path), unicode(message)])


    def set_profile(self, profile):
        self._current_profile = profile
        
    @property
    def current_book(self):
        return self._current_book
    
    @current_book.setter
    def current_book(self, book):
        if book.FilePath:
            self._current_book = book.FilePath
        else:
            self._current_book = book.Caption

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