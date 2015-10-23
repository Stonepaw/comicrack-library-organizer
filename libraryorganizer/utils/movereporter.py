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
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult  # @UnresolvedImport @IgnorePep8


class MoveReporter(object):

    def __init__(self):
        self._log = []
        self._current_profile = ""
        self._current_book = ""
        self.header = ""
        self._profile_reports = {}
        self.failed_or_skipped = False
        # TODO Store book in this instead of passing the path in ADD

    def fail(self, message):
        self.log("Failed", message)
        self._profile_reports[self._current_profile].failed()
        self.failed_or_skipped = True

    def warn(self, message):
        self.log("Warning", message)

    def success(self, message=""):
        if message:
            self.log("Success", message)
        self._profile_reports[self._current_profile].success()

    def skip(self, message, profile_name=""):
        if not profile_name:
            profile_name = self._current_profile
        self.log("Skipped", message, profile_name)
        self._profile_reports[profile_name].skipped()
        self.failed_or_skipped = True

    def success_simulated(self, message, action=""):
        if not action:
            action = "Success (Simulated)"
        self.log(action, message)
        self._profile_reports[self._current_profile].success()

    def log(self, action, message, profile_name=""):
        if not profile_name:
            profile_name = self._current_profile
        self._log.append([profile_name, action, self._current_book, message])

    def Add(self, action, path, message="", profile=""):
        if profile:
            self._log.append([profile, unicode(action), unicode(path),
                              unicode(message)])
        else:
            self._log.append([self._current_profile, unicode(action),
                              unicode(path), unicode(message)])

    def set_profile(self, profile):
        """ Sets the current profile name.

        Args:
            profile: A profile object.
        """
        self._current_profile = profile.Name

    def create_profile_reports(self, profiles, count):
        """ Creates the profile reports. Initialize this before calling
        set_profile.

        Args:
            profiles: A list of Profiles.
            count: The number of books that are being moved.
        """
        self._profile_reports = {profile.Name: ProfileReport(profile, count)
                                 for profile in profiles}

    def get_profile_reports(self, canceled=False):
        """Gets the profile reports as list of strings.

        Pass True to canceled to calculated the skipped count correctly when
        the operation is canceled.
        """
        return [self._profile_reports[profile].get_report()
                for profile in self._profile_reports]

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
    # TODO: Move into MoveReport functions
    def __init__(self, profile, count):
        self._success = 0
        self._failed = 0
        self._skipped = 0
        self._total = count
        self._profile = profile
        self._name = profile.Name
        self._mode = profile.Mode

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
            cancelled: A Boolean if the move process was cancelled. If true
                then the remaining books are counted as skipped.

        Returns:
            A string of the total success, failed, and skipped operations
        """
        if cancelled:
            self._skipped = self._total - self._success - self._failed
        report = "{0}:\nSuccess: {1}\tSkipped: {2}\tFailed: {3}".format(
            self._name, self._success, self._skipped, self._failed)
        return report
