"""
loworkerform.py

Contains the worker form. All the copying is done in the background worker of this form.


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
import System.Drawing
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
from System.Windows.Interop import WindowInteropHelper

import System.IO
from System.IO import StreamWriter, TextWriter

import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

import loduplicate
from loduplicate import DuplicateForm

from locommon import ICON, Mode

from bookmover import BookMover, UndoMover

import loforms

from lologger import Logger

class WorkerForm(Form):


    def __init__(self, books, profiles):
        self.InitializeComponent()
        self.books = books
        self.profiles = profiles
        #The percentage to raise to progress bar of one book.
        self.percentage = 1.0/len(books)/len(profiles)*100
        self.progress = 0.0
        self._DuplicateForm = None
        loduplicate.ComicRack = ComicRack
        

    def InitializeComponent(self):
        self._progress = System.Windows.Forms.ProgressBar()
        self._Worker = System.ComponentModel.BackgroundWorker()
        self._Cancel = System.Windows.Forms.Button()
        self.SuspendLayout()
        # 
        # progress
        # 
        self._progress.Dock = System.Windows.Forms.DockStyle.Top
        self._progress.Location = System.Drawing.Point(0, 0)
        self._progress.Name = "progress"
        self._progress.Size = System.Drawing.Size(350, 23)
        self._progress.TabIndex = 0
        #
        # Cancel
        #
        self._Cancel.Location = System.Drawing.Point(141, 29)
        self._Cancel.Size = System.Drawing.Size(75, 23)
        self._Cancel.TabIndex = 1
        self._Cancel.Text = "Cancel"
        self._Cancel.Click += self.WorkerFormFormClosing
        # 
        # Worker
        # 
        self._Worker.WorkerReportsProgress = True
        self._Worker.WorkerSupportsCancellation = True
        self._Worker.DoWork += self.WorkerDoWork
        self._Worker.ProgressChanged += self.WorkerProgressChanged
        self._Worker.RunWorkerCompleted += self.WorkerRunWorkerCompleted
        # 
        # WorkerForm
        # 
        self.ClientSize = System.Drawing.Size(350, 55)
        self.Controls.Add(self._progress)
        self.Controls.Add(self._Cancel)
        self.CancelButton = self._Cancel
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.Icon = System.Drawing.Icon(ICON)
        self.Name = "WorkerForm"
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.Text = "Processing Files...."
        self.FormClosing += self.WorkerFormFormClosing
        self.Shown += self.WorkerFormLoad
        self.ResumeLayout(False)


    def WorkerDoWork(self, sender, e):
        books = e.Argument[0]
        profiles = e.Argument[1]        
        report = WorkerResult()
        self.logger = Logger()
        mover = BookMover(sender, self, self.logger)

        
        result = mover.process_books(books, profiles)
        if result.failed_or_skipped:
            report.failed_or_skipped = result.failed_or_skipped

        report.report_text = result.report_text

        for profile in profiles:
            if profile.Mode == Mode.Simulate:
                report.show_report = True

        e.Result = report


    def WorkerProgressChanged(self, sender, e):
        self._progress.Value = e.ProgressPercentage


    def WorkerRunWorkerCompleted(self, sender, e):
        #print "Thread completed"
        if not e.Error:
            if e.Result.failed_or_skipped or e.Result.show_report:
                result = MessageBox.Show(e.Result.report_text + "\n\nWould you like to see a detailed report of the failed or skipped files?", "View full report?", MessageBoxButtons.YesNo)
                if result == DialogResult.Yes:
                    self.show_report()
            else:
                MessageBox.Show(e.Result.report_text)

        else:
            MessageBox.Show(str(e.Error))
        self.Close()


    def WorkerFormLoad(self, sender, e):
        self._Worker.RunWorkerAsync([self.books, self.profiles])

    
    def WorkerFormFormClosing(self, sender, e):
        if self._Worker.IsBusy:
            if self._Worker.CancellationPending == False:
                self._Worker.CancelAsync()
            self.DialogResult = DialogResult.None


    def ShowDuplicateForm(self, profile, newbook, oldbook, renamefile, count):

        if self._DuplicateForm == None:
            self._DuplicateForm = DuplicateForm(profile.Mode)
            helper = WindowInteropHelper(self._DuplicateForm.win)
            helper.Owner = self.Handle
        return self._DuplicateForm.ShowForm(newbook, oldbook, renamefile, count)


    def show_report(self):
        """Shows the report form with the available logger."""
        report = loforms.ReportForm()
        report.LoadData(self.logger.ToArray())
        r = report.ShowDialog()
        if r == DialogResult.Yes:
            self.logger.SaveLog()
        report.Dispose()



class ProfileSelector(Form):


    def __init__(self, profile_names, selected_profiles):
        self._profile_names = System.Array[str](profile_names)
        self._selected_profiles = selected_profiles
        self.initialize_component()

    
    def initialize_component(self):
        self._label1 = System.Windows.Forms.Label()
        self._first_profile = System.Windows.Forms.ComboBox()
        self._okay = System.Windows.Forms.Button()
        self._add = System.Windows.Forms.Button()
        self._profile_container = System.Windows.Forms.FlowLayoutPanel()
        self._panel1 = System.Windows.Forms.Panel()
        self._panel2 = System.Windows.Forms.Panel()
        self._panel3 = System.Windows.Forms.Panel()
        self._profile_container.SuspendLayout()
        self._panel1.SuspendLayout()
        self._panel2.SuspendLayout()
        self._panel3.SuspendLayout()
        self.SuspendLayout()
        # 
        # label1
        # 
        self._label1.Location = System.Drawing.Point(3, 9)
        self._label1.Size = System.Drawing.Size(260, 17)
        self._label1.TabIndex = 0
        self._label1.Text = "Select the profile(s) to be used:"
        # 
        # comboBox1
        # 
        self._first_profile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._first_profile.FormattingEnabled = True
        self._first_profile.Location = System.Drawing.Point(3, 3)
        self._first_profile.Size = System.Drawing.Size(190, 21)
        self._first_profile.TabIndex = 1
        self._first_profile.Margin = Padding(6, 0, 0, 6)
        self._first_profile.Items.AddRange(self._profile_names)
        self._first_profile.SelectedIndex = 0
        self._first_profile.Sorted = True
        # 
        # okay
        # 
        self._okay.DialogResult = System.Windows.Forms.DialogResult.OK
        self._okay.Location = System.Drawing.Point(174, 5)
        self._okay.Size = System.Drawing.Size(75, 23)
        self._okay.TabIndex = 2
        self._okay.Text = "OK"
        self._okay.UseVisualStyleBackColor = True
        # 
        # add
        # 
        self._add.Location = System.Drawing.Point(89, 5)
        self._add.Size = System.Drawing.Size(75, 23)
        self._add.TabIndex = 3
        self._add.Text = "Add"
        self._add.UseVisualStyleBackColor = True
        self._add.Click += self.add_click
        # 
        # flowLayoutPanel1
        # 
        self._profile_container.AutoSize = True
        self._profile_container.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self._profile_container.Controls.Add(self._first_profile)
        self._profile_container.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        self._profile_container.Location = System.Drawing.Point(3, 6)
        self._profile_container.Size = System.Drawing.Size(196, 27)
        self._profile_container.TabIndex = 4
        self._profile_container.WrapContents = False
        # 
        # panel1
        # 
        self._panel1.Controls.Add(self._okay)
        self._panel1.Controls.Add(self._add)
        self._panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        self._panel1.Location = System.Drawing.Point(0, 68)
        self._panel1.Size = System.Drawing.Size(264, 34)
        self._panel1.TabIndex = 5
        # 
        # panel2
        # 
        self._panel2.Controls.Add(self._label1)
        self._panel2.Dock = System.Windows.Forms.DockStyle.Top
        self._panel2.Location = System.Drawing.Point(0, 0)
        self._panel2.Size = System.Drawing.Size(264, 31)
        self._panel2.TabIndex = 3
        # 
        # panel3
        # 
        self._panel3.AutoScroll = True
        self._panel3.AutoSize = True
        self._panel3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self._panel3.Controls.Add(self._profile_container)
        self._panel3.Dock = System.Windows.Forms.DockStyle.Fill
        self._panel3.Location = System.Drawing.Point(0, 31)
        self._panel3.Size = System.Drawing.Size(264, 37)
        self._panel3.TabIndex = 7

        if self._selected_profiles:
            self._first_profile.SelectedItem = self._selected_profiles[0]

            for profile in self._selected_profiles[1:]:
                self.add_click(None, None, profile)

        # 
        # ProfileSelector
        # 
        self.AcceptButton = self._okay
        self.AutoSize = True
        self.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        self.ClientSize = System.Drawing.Size(264, 102)
        self.Controls.Add(self._panel3)
        self.Controls.Add(self._panel2)
        self.Controls.Add(self._panel1)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MaximumSize = System.Drawing.Size(270, 268)
        self.MinimizeBox = False
        self.MinimumSize = System.Drawing.Size(270, 130)
        self.StartPosition = FormStartPosition.CenterParent
        self.Text = "Select Profile(s) to use"
        self.Icon = Icon(ICON)
        self._profile_container.ResumeLayout(False)
        self._panel1.ResumeLayout(False)
        self._panel2.ResumeLayout(False)
        self._panel3.ResumeLayout(False)
        self._panel3.PerformLayout()
        self.ResumeLayout(False)
        self.PerformLayout()


    def add_click(self, sender, e, selected=None):
        c = ProfileSelectorControl(self.remove, self._profile_names, selected)
        self._profile_container.Controls.Add(c)
        
        
    def remove(self, sender, e):
        c = sender.Tag
        self._profile_container.Controls.Remove(c)
        c.Dispose()


    def get_profiles_to_use(self):
        return [control.SelectedItem for control in self._profile_container.Controls]
        
        
        
class ProfileSelectorControl(FlowLayoutPanel):
    
    def __init__(self, remove, profile_names, selected):
        self._profile = ComboBox()
        self._profile.Size = System.Drawing.Size(190, 21)
        self._profile.DropDownStyle = ComboBoxStyle.DropDownList
        self._profile.Items.AddRange(profile_names)
        self._profile.Sorted = True

        if selected is not None:
            self._profile.SelectedItem = selected
        else:
            self._profile.SelectedIndex = 0

        self._remove = Button()
        self._remove.Text = "-"
        self._remove.Size = Size(20, 23)
        self._remove.Click += remove
        self._remove.Tag = self
        self.AutoSize = True
        self.Controls.Add(self._profile)
        self.Controls.Add(self._remove)
        
    @property
    def SelectedItem(self):
        return self._profile.SelectedItem



class WorkerResult(object):
    
    def __init__(self):
        self.failed_or_skipped = False
        self.report_text = ""
        self.show_report = False


        
class ReportForm(Form):
    def __init__(self, text):
        self.InitializeComponent()
        self._report.Text = text
    
    def InitializeComponent(self):
        self._button1 = System.Windows.Forms.Button()
        self._report = System.Windows.Forms.RichTextBox()
        self.SuspendLayout()
        # 
        # button1
        # 
        self._button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._button1.DialogResult = System.Windows.Forms.DialogResult.OK
        self._button1.Location = System.Drawing.Point(370, 399)
        self._button1.Name = "button1"
        self._button1.Size = System.Drawing.Size(75, 23)
        self._button1.TabIndex = 1
        self._button1.Text = "OK"
        self._button1.UseVisualStyleBackColor = True
        # 
        # report
        # 
        self._report.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._report.Location = System.Drawing.Point(12, 12)
        self._report.Name = "report"
        self._report.ReadOnly = True
        self._report.Size = System.Drawing.Size(425, 379)
        self._report.TabIndex = 4
        self._report.Text = ""
        # 
        # lomessagebox
        # 
        self.AcceptButton = self._button1
        self.ClientSize = System.Drawing.Size(453, 435)
        self.Controls.Add(self._report)
        self.Controls.Add(self._button1)
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.Icon = System.Drawing.Icon(ICON)
        self.Name = "ReportForm"
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.Text = "Report"
        self.ResumeLayout(False)



class WorkerFormUndo(WorkerForm):
    """
    Can use pretty much all the code from the WorkerForm. Just have to override a couple of things
    """
    
    def __init__(self, undo_collection, profiles):
        super(WorkerFormUndo,self).__init__(undo_collection, profiles)
        self.Text = "Undoing last move..."


    def WorkerDoWork(self, sender, e):
        undo_collection = e.Argument[0]
        profiles = e.Argument[1]
        self.logger = Logger()
        report = WorkerResult()
        
        mover = UndoMover(sender, self, undo_collection, profiles, self.logger)
        result = mover.process_books()        

        if result.failed_or_skipped:
            report.failed_or_skipped = result.failed_or_skipped

        report.report_text += result.report_text

        e.Result = report