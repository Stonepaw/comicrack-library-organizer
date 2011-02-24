import clr
import System
clr.AddReference("System.Windows.Forms")
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

from locommon import ICON

class loMessageBox(Form):
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
		self.Name = "lomessagebox"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Report"
		self.ResumeLayout(False)

