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
        self._profile_reports = {}
        #TODO Store book in this instead of passing the path in ADD

    def fail(self, message):
        self._add_log_message("Failed", message)
        self._profile_reports[self._current_profile].failed()
    
    def warn(self, message):
        self._add_log_message("Warning", message)
    
    def success(self, message):
        self._add_log_message("Success", message)
        self._profile_reports[self._current_profile].success()
        
    def skip(self, message, profile_name=""):
        if not profile_name:
            profile_name = self._current_profile
        self._add_log_message("Skipped", message, profile_name)
        self._profile_reports[profile_name].skipped()
        
    def success_simulated(self, message):
        self._add_log_message("Success (Simulated)", message)
        self._profile_reports[self._current_profile].success()
    
    def _add_log_message(self, action, message, profile_name=""):
        if not profile_name:
            profile_name = self._current_profile
        self._log.append(profile_name, action, self._current_book, message)

    def Add(self, action, path, message = "", profile=""):
        if profile:
            self._log.append([profile, unicode(action), unicode(path), unicode(message)])
        else:
            self._log.append([self._current_profile, unicode(action), unicode(path), unicode(message)])


    def set_profile(self, profile):
        self._current_profile = profile.Name
        if profile not in self._profile_reports.keys():
            self._profile_reports[profile.Name] = ProfileReport(profile.Name, profile.Mode)
        
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
        
class ProfileReport(object):
    """Provides way to report on each profile's results for the move operation
    """
    #TODO: Move into MoveReport functions
    def __init__(self, name, mode):
        self._success = 0
        self._failed = 0
        self._skipped = 0
        #self._total = total
        self._name = name
        self._mode = mode
        
    def failed(self):
        self._failed += 1
        
    def success(self):
        self._success += 1
        
    def skipped(self):
        self._skipped += 1

    def get_report(self, cancelled):
        """Creates the report to display to the user of this profile's move
        operation results

        Args:
            cancelled: A Boolean if the move process was cancelled, a different
                report is created if true.

        Returns:
            A string of the total _success, _failed, and _skipped operations
        """
        if cancelled:
            self._skipped = self._total - self._success - self._failed
        return "%s:\nSuccessfully %s: %s\tSkipped: %s\tFailed: %s" % (self._name, ModeText.get_mode_past(self._mode), self._success,
                                                                      self._skipped, self._failed)
