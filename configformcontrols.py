"""
configformcontrols.py

Contains various custom controls used in the config form .

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

import pyevent

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import Appearance, Padding, FlowLayoutPanel, TextBox, Button, Panel, Label, AutoSizeMode, NumericUpDown, CheckBox, ComboBox, ComboBoxStyle, BorderStyle, BindingSource
                                 
from System.Drawing import Size, Point, ContentAlignment

from locommon import ExcludeGroup, ExcludeRule

import lodpi



class InsertControl(FlowLayoutPanel):
    """The basic insert control. Contains a prefix textbox, button and postfix textbox.""" 
    Insert, _insert = pyevent.make_event()

    def __init__(self):
        
        self.SuspendLayout()
        self.Margin = Padding(0)
        self.Template = ""
        self.WrapContents = True
        
        self.AutoSize = True
        self.AutoSizeMode = AutoSizeMode.GrowAndShrink
        
        self.Prefix = TextBox()
        self.InsertButton = Button()
        self.Postfix = TextBox()
        self.LabelPanel = Panel()
        SpacingFix = Panel()
                
        #This fixes a wierd bug where when the panel is set to flowbreak, extra space is added to the right of the control
        #The extra space is equal to the following control. There for a control with size 0,0 can be added to fix the bug
        SpacingFix.Size = Size(0,0)
        SpacingFix.Margin = Padding(0)

        self.Prefix.Size = lodpi.textbox_size()
        self.Prefix.TabIndex = 0
        self.Prefix.Margin = Padding(3, 0, 3, 0)

        self.InsertButton.Size = lodpi.scaled_size(66, 23)
        self.InsertButton.MinimumSize = lodpi.scaled_size(66, 23)
        self.InsertButton.AutoSize = True
        self.InsertButton.Click += self.ButtonClick
        self.InsertButton.TabIndex = 1
        self.InsertButton.Margin = Padding(3, 0, 3, 0)

        self.Postfix.Size = lodpi.textbox_size()
        self.Postfix.TabIndex = 2
        self.Postfix.Margin = Padding(3, 0, 3, 0)
        
        self.LabelPanel.Height = lodpi.scale_int(17)
        self.LabelPanel.Margin = Padding(0, 0, 0, 0)

        self.Controls.Add(self.LabelPanel)
        self.Controls.Add(SpacingFix)
        self.Controls.Add(self.Prefix)
        self.Controls.Add(self.InsertButton)
        self.Controls.Add(self.Postfix)
        self.SetFlowBreak(self.LabelPanel, True)

        self.ResumeLayout()
        

    def SetPrefixText(self, text):
        self.Prefix.Text = text
        

    def SetPostfixText(self, text):
        self.Postfix.Text = text
        

    def SetTemplate(self, template, text):
        """
        sets the template field of the control. template is a string. Text sets the text of the button.
        """
        self.Template = template
        self.InsertButton.Text = text
        

    def SetLabels(self, prefix_text="", button_text="", postfix_text=""):
        """Sets the labels for this control."""
        self.SuspendLayout()
        
        self.PrefixLabel = Label()
        self.PrefixLabel.Text = prefix_text
        self.PrefixLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.PrefixLabel)
        self.PrefixLabel.Location = Point(self.Prefix.Location.X + self.Prefix.Width/2 - self.PrefixLabel.Width/2, 0)
        self.PrefixLabel.Margin = Padding(0)

        self.ButtonLabel = Label()
        self.ButtonLabel.Text = button_text
        self.ButtonLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.ButtonLabel)
        self.ButtonLabel.Location = Point(self.InsertButton.Location.X + self.InsertButton.Width/2 - self.ButtonLabel.Width/2, 0)
        self.ButtonLabel.Margin = Padding(0)
        
        self.PostfixLabel = Label()
        self.PostfixLabel.Text = postfix_text
        self.PostfixLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.PostfixLabel)
        self.PostfixLabel.Location = Point(self.Postfix.Location.X + self.Postfix.Width/2 - self.PostfixLabel.Width/2, 0)
        self.PostfixLabel.Margin = Padding(0)
        
        self.LabelPanel.Height = self.PrefixLabel.Height
        
        self.ResumeLayout()
        self.RefreshLabelLayout()
        

    def RefreshLabelLayout(self):
        """Re-center column labels after control sizes change (HiDPI)."""
        if not hasattr(self, "PrefixLabel"):
            return
        self.SuspendLayout()
        self.PrefixLabel.Location = Point(self.Prefix.Location.X + self.Prefix.Width/2 - self.PrefixLabel.Width/2, 0)
        self.ButtonLabel.Location = Point(self.InsertButton.Location.X + self.InsertButton.Width/2 - self.ButtonLabel.Width/2, 0)
        self.PostfixLabel.Location = Point(self.Postfix.Location.X + self.Postfix.Width/2 - self.PostfixLabel.Width/2, 0)
        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.LabelPanel.PreferredSize.Width
        self.ResumeLayout()

    def apply_hidpi_metrics(self, owner=None, scale=None):
        """Re-apply scaled metrics using the Configure form as DPI source."""
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.SuspendLayout()
        self.Prefix.Size = lodpi.textbox_size(scale=scale, owner=owner)
        self.Postfix.Size = lodpi.textbox_size(scale=scale, owner=owner)
        btn = lodpi.scaled_size(66, 23, scale, owner)
        self.InsertButton.Size = btn
        self.InsertButton.MinimumSize = btn
        self.LabelPanel.Height = lodpi.scale_int(17, scale, owner)
        self.ResumeLayout()
        self.RefreshLabelLayout()
        

    def GetPrefixText(self):
        return self.Prefix.Text
    

    def GetPostfixText(self):
        return self.Postfix.Text
    

    def ButtonClick(self, sender, e):
        self._insert(self, None)
        

    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "
        return "{" + s + self.Prefix.Text + "<" + self.Template + ">" + self.Postfix.Text + "}"

    

class InsertControlNumber(InsertControl):
    """
    Insert Control containing a prefix textbox, button, postfix textbox, and a numeric updown
    """
    def __init__(self):
        
        InsertControl.__init__(self)
        
        self.Pad = NumericUpDown()
        self.Pad.Size = lodpi.scaled_size(lodpi.NUMERIC_WIDTH, lodpi.TEXTBOX_HEIGHT)
        self.Pad.TabIndex = 3
        self.Pad.Margin = Padding(3, 0, 3, 0)
        
        self.Controls.Add(self.Pad)

        self.Width = self.PreferredSize.Width

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.Pad.Size = lodpi.scaled_size(lodpi.NUMERIC_WIDTH, lodpi.TEXTBOX_HEIGHT, scale, owner)
        self.RefreshLabelLayout()
        

    def SetLabels(self, prefix_text, button_text, postfix_text, padding_text):
        
        InsertControl.SetLabels(self, prefix_text, button_text, postfix_text)
        

        self.SuspendLayout()
        
        self.PaddingLabel = Label()
        self.PaddingLabel.Text = padding_text
        self.PaddingLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.PaddingLabel)
        self.PaddingLabel.Location = Point(self.Pad.Location.X + self.Pad.Width/2 - self.PaddingLabel.Width/2, 0)
        self.PaddingLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.LabelPanel.PreferredSize.Width
        
        self.ResumeLayout()
        

    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "
        return "{" + s + self.Prefix.Text + "<" + self.Template + str(self.Pad.Value) + ">" + self.Postfix.Text + "}"

    

class InsertControlStartMonth(InsertControlNumber):
    """
    Insert Control containing a prefix textbox, button, postfix textbox, textbox and a checkbox
    """
    
    def __init__(self, check_text):
        InsertControlNumber.__init__(self)

        self.Check = CheckBox()
        self.Check.AutoSize = True
        self.Check.Text = check_text
        
        self.Controls.Add(self.Check)

        self.Width = self.PreferredSize.Width

    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "

        insert = ""
        if not self.Check.Checked:
            insert = "#" + str(self.Pad.Value)

        return "{" + s + self.Prefix.Text + "<" + self.Template + insert + ">" + self.Postfix.Text + "}"

    

class InsertControlMultipleValue(InsertControl):
    """
    Insert Control containing a prefix textbox, button, postfix textbox, seperator textbox and a checkbox
    """
    def __init__(self):
        InsertControl.__init__(self)

        self.Prefix.Width = lodpi.scale_int(62)
        self.Postfix.Width = lodpi.scale_int(62)

        self.Seperator = TextBox()
        self.Seperator.Size = lodpi.scaled_size(24, lodpi.TEXTBOX_HEIGHT)
        self.Seperator.TabIndex = 3
        self.Seperator.Margin = Padding(3, 0, 3, 0)

        self.Check = CheckBox()
        self.Check.Size = Size(15, 14)
        self.Check.TabIndex = 4
        self.Check.Margin.Top = 5

        self.Controls.Add(self.Seperator)
        self.Controls.Add(self.Check)

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.SuspendLayout()
        self.Prefix.Width = lodpi.scale_int(62, scale, owner)
        self.Postfix.Width = lodpi.scale_int(62, scale, owner)
        self.Seperator.Size = lodpi.scaled_size(24, lodpi.TEXTBOX_HEIGHT, scale, owner)
        self.ResumeLayout()
        self.RefreshLabelLayout()
        
    def SetLabels(self, prefix_label, button_label, postfix_label, sperator_label, checkbox_label):
        
        self.Width = self.PreferredSize.Width
        
        InsertControl.SetLabels(self, prefix_label, button_label, postfix_label)
        
        self.SuspendLayout()
        
        self.SeperatorLabel = Label()
        self.SeperatorLabel.Text = sperator_label
        self.SeperatorLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.SeperatorLabel)
        self.SeperatorLabel.Location = Point(self.Seperator.Location.X + self.Seperator.Width/2 - self.SeperatorLabel.Width/2, 0)
        self.SeperatorLabel.Margin = Padding(0)
        
        self.CheckboxLabel = Label()
        self.CheckboxLabel.Text = checkbox_label
        self.CheckboxLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.CheckboxLabel)
        self.CheckboxLabel.Location = Point(self.Check.Location.X + self.Check.Width/2 - self.CheckboxLabel.Width/2, 0)
        self.CheckboxLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.PreferredSize.Width
        
        self.ResumeLayout()
        self.RefreshLabelLayout()

    def RefreshLabelLayout(self):
        if not hasattr(self, "PrefixLabel"):
            return
        self.SuspendLayout()
        InsertControl.RefreshLabelLayout(self)
        if hasattr(self, "SeperatorLabel"):
            self.SeperatorLabel.Location = Point(self.Seperator.Location.X + self.Seperator.Width/2 - self.SeperatorLabel.Width/2, 0)
        if hasattr(self, "CheckboxLabel"):
            self.CheckboxLabel.Location = Point(self.Check.Location.X + self.Check.Width/2 - self.CheckboxLabel.Width/2, 0)
        self.LabelPanel.Width = max(self.PreferredSize.Width, self.LabelPanel.PreferredSize.Width)
        self.ResumeLayout()

    def GetSeperatorText(self):
        return self.Seperator.Text

    def SetSeperatorText(self, text):
        self.Seperator.Text = text      

    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        if self.Check.Checked:
            checktext = "(series)"
        else:
            checktext = "(issue)"
        sep = self.Seperator.Text
        s = ""
        if space:
            s = " "
            sep = sep.rstrip() + s
        return "{" + s + self.Prefix.Text + "<" + self.Template + "(" + sep + ")" + checktext + ">" + self.Postfix.Text + "}"

    

class InsertControlYesNo(InsertControl):
    """
    Insert Control containing a prefix textbox, button, postfix textbox and a textbox
    """
    def __init__(self):
        InsertControl.__init__(self)
        
        self.Invert = CheckBox()
        self.Invert.Margin = Padding(3, 0, 3, 0)
        self.Invert.Text = "!"
        self.Invert.TabIndex = 4
        self.Invert.Padding = Padding(0)
        self.Invert.Appearance = Appearance.Button
        self.Invert.AutoSize = True
        
        self.TextBox = TextBox()
        self.TextBox.Size = lodpi.textbox_size()
        self.TextBox.TabIndex = 3
        self.TextBox.Margin = Padding(3, 0, 3, 0)
        self.Controls.Add(self.TextBox)
        self.Controls.Add(self.Invert)
        self.Width = self.PreferredSize.Width

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.TextBox.Size = lodpi.textbox_size(scale=scale, owner=owner)
        self.RefreshLabelLayout()
        
        
    def SetTextBoxText(self, text):
        self.TextBox.Text = text
        

    def GetTextBoxText(self):
        return self.TextBox.Text
    
    
    def SetLabels(self, prefix_label, button_label, postfix_label, textbox_label):
        
        self.Width = self.PreferredSize.Width
        
        InsertControl.SetLabels(self, prefix_label, button_label, postfix_label)
        
        self.SuspendLayout()
        
        self.TextBoxLabel = Label()
        self.TextBoxLabel.Text = textbox_label
        self.TextBoxLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.TextBoxLabel)
        self.TextBoxLabel.Location = Point(self.TextBox.Location.X + self.TextBox.Width/2 - self.TextBoxLabel.Width/2, 0)
        self.TextBoxLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.PreferredSize.Width
        
        self.ResumeLayout()
    

    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "
            
        invert = ""
        if self.Invert.Checked:
            invert = "(!)"
            
        return "{" + s + self.Prefix.Text + "<" + self.Template + "(" + self.TextBox.Text + ")" + invert + ">" + self.Postfix.Text + "}"

    

class InsertControlReadPercentage(InsertControl):
    """
    Insert control with a prefix textbox, button, postfix button, textbox, combobox, and numeric updown
    """

    def __init__(self):

        InsertControl.__init__(self)

        self.TextBox = TextBox()
        self.TextBox.Size = lodpi.textbox_size()
        self.TextBox.TabIndex = 3
        self.TextBox.Margin = Padding(3, 0, 3, 0)
        

        self.Operator = ComboBox()
        self.Operator.Items.AddRange(System.Array[System.String](["equal to", "greater than", "less than"]))
        self.Operator.DropDownStyle = ComboBoxStyle.DropDownList
        self.Operator.SelectedItem = "greater than"
        self.Operator.Width = lodpi.scale_int(80)
        self.Operator.Margin = Padding(3, 0, 3, 0)

        self.Percentage = NumericUpDown()
        self.Percentage.Maximum = 100
        self.Percentage.Value = 90
        self.Percentage.AutoSize = True
        self.Percentage.Margin = Padding(3, 0, 3, 0)
        
        self.label = Label()
        self.label.Text = "%"
        self.label.TextAlign = ContentAlignment.MiddleLeft

        self.Controls.Add(self.TextBox)
        self.Controls.Add(self.Operator)
        self.Controls.Add(self.Percentage)
        self.Controls.Add(self.label)

        self.Width = self.PreferredSize.Width

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.TextBox.Size = lodpi.textbox_size(scale=scale, owner=owner)
        self.Operator.Width = lodpi.scale_int(80, scale, owner)
        self.RefreshLabelLayout()


    def SetLabels(self, prefix_label, button_label, postfix_label, text_label, operator_label, percentage_label):
        
        InsertControl.SetLabels(self, prefix_label, button_label, postfix_label)

        self.SuspendLayout()

        self.TextBoxLabel = Label()
        self.TextBoxLabel.Text = text_label
        self.TextBoxLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.TextBoxLabel)
        self.TextBoxLabel.Location = Point(self.TextBox.Location.X + self.TextBox.Width/2 - self.TextBoxLabel.Width/2, 0)
        self.TextBoxLabel.Margin = Padding(0)
        
        self.OperatorLabel = Label()
        self.OperatorLabel.Text = operator_label
        self.OperatorLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.OperatorLabel)
        self.OperatorLabel.Location = Point(self.Operator.Location.X + self.Operator.Width/2 - self.OperatorLabel.Width/2, 0)
        self.OperatorLabel.Margin = Padding(0)

        self.PercentageLabel = Label()
        self.PercentageLabel.Text = percentage_label
        self.PercentageLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.PercentageLabel)
        self.PercentageLabel.Location = Point(self.Percentage.Location.X + self.Percentage.Width/2 - self.PercentageLabel.Width/2, 0)
        self.PercentageLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.PreferredSize.Width
        
        self.ResumeLayout()


    def SetTextBoxText(self, text):
        self.TextBox.Text = text


    def GetTextBoxText(self):
        return self.TextBox.Text

    def GetOperator(self):
        if self.Operator.SelectedItem == "equal to":
            return "="
        elif self.Operator.SelectedItem == "greater than":
            return ">"
        elif self.Operator.SelectedItem == "less than":
            return "<"


    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "
        

        return "{" + s + self.Prefix.Text + "<" + self.Template + "(" + self.TextBox.Text + ")(" + self.GetOperator() + ")(" + self.Percentage.Value.ToString() + ")>" + self.Postfix.Text + "}"

    

class InsertControlFirstLetter(InsertControl):
    """
    Insert control containing a prefix textbox, button, postfix textbox and a combobox
    """
    def __init__(self):
        InsertControl.__init__(self)

        self.ComboBox = ComboBox()
        self.ComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        self.ComboBox.Width = lodpi.scale_int(100)
        self.ComboBox.Margin = Padding(3, 0, 3, 0)
        
        self.Controls.Add(self.ComboBox)
        self.Width = self.PreferredSize.Width

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.ComboBox.Width = lodpi.scale_int(100, scale, owner)
        self.RefreshLabelLayout()


    def SetLabels(self, prefix_label, button_label, postfix_label, combobox_label):
        InsertControl.SetLabels(self,prefix_label, button_label, postfix_label)

        self.SuspendLayout()
        
        self.ComboBoxLabel = Label()
        self.ComboBoxLabel.Text = combobox_label
        self.ComboBoxLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.ComboBoxLabel)
        self.ComboBoxLabel.Location = Point(self.ComboBox.Location.X + self.ComboBox.Width/2 - self.ComboBoxLabel.Width/2, 0)
        self.ComboBoxLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.LabelPanel.PreferredSize.Width
        
        self.ResumeLayout()


    def SetComboBoxItems(self, items):
        """
        Sets the ComboBox items.
        items should be either an array or list or strings
        """

        self.ComboBox.Items.Clear()
        self.ComboBox.Items.AddRange(System.Array[System.String](items))
        self.ComboBox.SelectedIndex = 0 if len(items) > 0 else -1


    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "

        return "{" + s + self.Prefix.Text + "<" + self.Template + "(" + self.ComboBox.SelectedItem + ")>" + self.Postfix.Text + "}"

    

class InsertControlCounter(InsertControl):
    """
    Insert Control containing a prefix textbox, button, postfix textbox, and three numerical updown controls
    """

    def __init__(self):
        InsertControl.__init__(self)

        self.Start = NumericUpDown()
        self.Start.Size = lodpi.scaled_size(lodpi.NUMERIC_WIDTH, lodpi.TEXTBOX_HEIGHT)
        self.Start.TabIndex = 3
        self.Start.Increment = 1
        self.Start.Value = 1
        self.Start.Margin = Padding(3, 0, 3, 0)

        self.Increment = NumericUpDown()
        self.Increment.Size = lodpi.scaled_size(lodpi.NUMERIC_WIDTH, lodpi.TEXTBOX_HEIGHT)
        self.Increment.TabIndex = 4
        self.Increment.Increment = 1
        self.Increment.Value = 1
        self.Increment.Margin = Padding(3, 0, 3, 0)

        self.Pad = NumericUpDown()
        self.Pad.Size = lodpi.scaled_size(lodpi.NUMERIC_WIDTH, lodpi.TEXTBOX_HEIGHT)
        self.Pad.TabIndex = 5
        self.Pad.Increment = 1
        self.Pad.Value = 0
        self.Pad.Margin = Padding(3, 0, 3, 0)

        self.Controls.Add(self.Start)
        self.Controls.Add(self.Increment)
        self.Controls.Add(self.Pad)

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        num = lodpi.scaled_size(lodpi.NUMERIC_WIDTH, lodpi.TEXTBOX_HEIGHT, scale, owner)
        self.Start.Size = num
        self.Increment.Size = num
        self.Pad.Size = num
        self.RefreshLabelLayout()

    def SetLabels(self, prefix_label, button_label, postfix_label, start_label, increment_label, padding_label):
        InsertControl.SetLabels(self, prefix_label, button_label, postfix_label)

        self.Width = self.PreferredSize.Width
        
        self.SuspendLayout()

        self.StartLabel = Label()
        self.StartLabel.Text = start_label
        self.StartLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.StartLabel)
        self.StartLabel.Location = Point(self.Start.Location.X + self.Start.Width/2 - self.StartLabel.Width/2, 0)
        self.StartLabel.Margin = Padding(0)

        
        self.IncrementLabel = Label()
        self.IncrementLabel.Text = increment_label
        self.IncrementLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.IncrementLabel)
        self.IncrementLabel.Location = Point(self.Increment.Location.X + self.Increment.Width/2 - self.IncrementLabel.Width/2, 0)
        self.IncrementLabel.Margin = Padding(0)
        
        self.PaddingLabel = Label()
        self.PaddingLabel.Text = padding_label
        self.PaddingLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.PaddingLabel)
        self.PaddingLabel.Location = Point(self.Pad.Location.X + self.Pad.Width/2 - self.PaddingLabel.Width/2, 0)
        self.PaddingLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.PreferredSize.Width

    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "
        return "{" + s + self.Prefix.Text + "<" + self.Template + "(" + str(self.Start.Value) + ")(" + str(self.Increment.Value) + ")(" + str(self.Pad.Value) + ")>" + self.Postfix.Text + "}"



class InsertControlDateTime(InsertControl):
    """
    Insert control containing a prefix textbox, button, postfix textbox and a combobox
    """
    def __init__(self):
        InsertControl.__init__(self)

        self.ComboBox = ComboBox()
        self.ComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        self.SetComboBoxItems()
        self.ComboBox.Width = lodpi.scale_int(100)
        self.ComboBox.Margin = Padding(3, 0, 3, 0)
        
        
        self.Controls.Add(self.ComboBox)
        self.Width = self.PreferredSize.Width

    def apply_hidpi_metrics(self, owner=None, scale=None):
        InsertControl.apply_hidpi_metrics(self, owner, scale)
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        self.ComboBox.Width = lodpi.scale_int(100, scale, owner)
        self.RefreshLabelLayout()


    def SetLabels(self, prefix_label, button_label, postfix_label, combobox_label):
        InsertControl.SetLabels(self,prefix_label, button_label, postfix_label)

        self.SuspendLayout()
        
        self.ComboBoxLabel = Label()
        self.ComboBoxLabel.Text = combobox_label
        self.ComboBoxLabel.AutoSize = True
        self.LabelPanel.Controls.Add(self.ComboBoxLabel)
        self.ComboBoxLabel.Location = Point(self.ComboBox.Location.X + self.ComboBox.Width/2 - self.ComboBoxLabel.Width/2, 0)
        self.ComboBoxLabel.Margin = Padding(0)

        self.LabelPanel.Height = self.PrefixLabel.Height
        self.LabelPanel.Width = self.LabelPanel.PreferredSize.Width
        
        self.ResumeLayout()


    def SetComboBoxItems(self):
        """
        Sets the ComboBox items.
        items should be either an array or list or strings
        """
        datetime = System.DateTime(2013, 2, 1)

        format_dictionary = System.Collections.Generic.Dictionary[System.String, System.String]()
        format_dictionary.Add("dd/MM/yy", datetime.ToString("dd/MM/yy"))
        format_dictionary.Add("MM/dd/yy", datetime.ToString("MM/dd/yy"))
        format_dictionary.Add("D", datetime.ToString("D"))
        format_dictionary.Add("M", datetime.ToString("M"))
        format_dictionary.Add("Y", datetime.ToString("Y"))
        format_dictionary.Add("MMM d, yyy", datetime.ToString("MMM d, yyy"))
        format_dictionary.Add("MMM dd, yyy", datetime.ToString("MMM dd, yyy"))
        format_dictionary.Add("MMMM d, yyy", datetime.ToString("MMMM d, yyy"))
        format_dictionary.Add("MMMM dd, yyy", datetime.ToString("MMMM dd, yyy"))
        self.ComboBox.DisplayMember = "Value"
        self.ComboBox.ValueMember = "Key"
        self.ComboBox.DataSource = BindingSource(format_dictionary, "")


    def GetTemplateText(self, space):
        """
        Builds the template text and returns it as a string
        space-> Boolean if automatically inserting spaceing
        """
        s = ""
        if space:
            s = " "

        return "{" + s + self.Prefix.Text + "<" + self.Template + "(" + self.ComboBox.SelectedValue + ")>" + self.Postfix.Text + "}"


class MetadataExcludeGroupControl(Panel):
    """A control to display a metadata exclude group"""

    def __init__(self, remove_method, exclude_rule_group=None):

        self.Size = Size(451, 70)
        self.MinimumSize = Size(451, 70)
        self.AutoSize = True

        #Labels
        self._label1 = Label()
        self._label1.Text = "Match"
        self._label1.Size = Size(40, 23)
        self._label1.Location = Point(3, 2)
        self._label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        self._label2 = Label()
        self._label2.Location = Point(96, 4)
        self._label2.Text = "of the following rules"
        self._label2.Size = Size(120, 20)
        self._label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        #Operator combobox
        self._operator = ComboBox()
        self._operator.DropDownStyle = ComboBoxStyle.DropDownList
        self._operator.Location = Point(44, 2)
        self._operator.Items.AddRange(System.Array[System.Object](
            ["Any",
            "All"]))
        self._operator.Size = Size(46, 23)
        self._operator.SelectedIndex = 0

        #Add buttons
        self._add_rule = Button()
        self._add_rule.Size = Size(75, 23)
        self._add_rule.Location = Point(346, 2)
        self._add_rule.Text = "Add Rule"
        self._add_rule.Click += self.add_rule

        self._add_group = Button()
        self._add_group.Location = Point(265, 2)
        self._add_group.Size = Size(75, 23)
        self._add_group.Text = "Add Group"
        self._add_group.Click += self.add_group
        
        #Remove buttons
        self._remove = Button()
        self._remove.Text = "-"
        self._remove.Location = Point(427, 2)
        self._remove.Size = Size(18, 23)
        self._remove.Tag = self
        self._remove.Click += remove_method
    
        #Container
        self._rules_container = FlowLayoutPanel()
        self._rules_container.Size = Size(458, 30)
        self._rules_container.MinimumSize = Size(458, 30)
        self._rules_container.AutoSize = True
        self._rules_container.Location = Point(15, 32)
        self._rules_container.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        
        self.Controls.Add(self._label1)
        self.Controls.Add(self._label2)
        self.Controls.Add(self._operator)
        self.Controls.Add(self._add_group)
        self.Controls.Add(self._add_rule)
        self.Controls.Add(self._remove)
        self.Controls.Add(self._rules_container)

        if exclude_rule_group is not None:

            self._operator.SelectedItem = exclude_rule_group.operator

            for rule in exclude_rule_group.rules:
                if type(rule) is ExcludeGroup:
                    self.add_group(None, None, rule)
                else:
                    self.add_rule(None, None, rule)
        else:
            self.add_rule(None, None, None)


    def add_rule(self, sender, e, rule=None):
        """Adds a new metadata exclude rule control to this group"""
        rule = MetadataExcludeRuleControl(self.remove_rule, rule)

        self._rules_container.Controls.Add(rule)


    def add_group(self, sender, e, rule_group=None):
        """Adds a new metadata exclude group control to this group"""
        group = MetadataExcludeGroupControl(self.remove_rule, rule_group)

        self._rules_container.Controls.Add(group)


    def remove_rule(self, sender, e):
        """Removes a metadata exclude rule from this group"""
        self._rules_container.Controls.Remove(sender.Tag)


    def set_operator(self, operator):
        if operator in ("Any", "All"):
            self._operator.SelectedItem = operator


    def get_rule(self):
        """Returns an ExcludeGroup object containing all the rules in this rule group control"""

        rules = [rule.get_rule() for rule in self._rules_container.Controls]

        return ExcludeGroup(self._operator.SelectedItem, rules)

    def apply_hidpi_metrics(self, owner=None, scale=None):
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        w = lodpi.scale_int(451, scale, owner)
        self.Width = w
        self.MinimumSize = Size(w, lodpi.scale_int(70, scale, owner))
        rc_w = lodpi.scale_int(458, scale, owner)
        self._rules_container.Width = rc_w
        self._rules_container.MinimumSize = Size(rc_w, lodpi.scale_int(30, scale, owner))
        lodpi.layout_row(
            [self._label1, self._operator, self._label2],
            3, 2, owner, 8, scale=scale)
        margin = lodpi.scale_int(3, scale, owner)
        row_y = lodpi.scale_int(2, scale, owner)
        self._remove.Location = Point(w - self._remove.Width - margin, row_y)
        self._add_rule.Location = Point(
            self._remove.Left - self._add_rule.Width - margin, row_y)
        self._add_group.Location = Point(
            self._add_rule.Left - self._add_group.Width - margin, row_y)
        for child in self._rules_container.Controls:
            if hasattr(child, "apply_hidpi_metrics"):
                child.apply_hidpi_metrics(owner, scale)


    
class MetadataExcludeRuleControl(FlowLayoutPanel):
    """This class is the display panel of a metadata exclude rule. It can either be in a rule group or by itself"""
    
    def __init__(self, remove_function, rule=None):
        
        #Flow Layout Panel
        self.Size = Size(451, 30)
        
        #Field selector
        self._field = ComboBox()
        self._field.Items.AddRange(System.Array[System.String](["Age Rating", "Alternate Count", "Alternate Number", "Alternate Series",
                                                               "Black And White", "Characters", "Count", "File Name", "File Path", "File Format",
                                                               "Format", "Genre", "Imprint", "Language", "Locations", "Main Character Or Team", "Manga", "Month", "Number",
                                                               "Notes", "Publisher", "Rating", "Read Percentage", "Review", "Series Complete", "Tags", "Teams", 
                                                               "Title", "Scan Information", "Series", "Series Group", "Start Month", "Start Year", "Story Arc", "Volume", "Web", "Year"]))            
        self._field.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._field.Size = Size(121, 21)
        self._field.MaxDropDownItems = 15
        self._field.IntegralHeight = False
        self._field.Sorted = True
        self._field.SelectedIndex= 0

     
        #Operator selector
        self._operator = ComboBox()
        self._operator.Items.AddRange(System.Array[System.String](
            ["contains",
            "does not contain",
            "greater than",
            "less than",
            "is",
            "is not"]))
        self._operator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._operator.Size = Size(110, 21)
        self._operator.SelectedIndex = 0
        
        #Value Textbox
        self._value_textbox = TextBox()
        self._value_textbox.Size = Size(175, 20)
        
        #Value Combobox
        self._value_combobox = ComboBox()
        self._value_combobox.Size = Size(175, 21)
        self._value_combobox.Enabled = False
        self._value_combobox.Visible = False
        self._value_combobox.DropDownStyle = ComboBoxStyle.DropDownList

        #Remove Button
        self._remove = Button()
        self._remove.Size = Size(18, 23)
        self._remove.Text = "-"
        self._remove.Tag = self
        self._remove.Click += remove_function

        self._field.SelectedIndexChanged += self.field_selection_index_changed

        if rule is not None:
            self.set_fields(rule)

        
        #Add controls
        self.Controls.Add(self._field)
        self.Controls.Add(self._operator)
        self.Controls.Add(self._value_textbox)
        self.Controls.Add(self._value_combobox)
        self.Controls.Add(self._remove)

    def apply_hidpi_metrics(self, owner=None, scale=None):
        if scale is None:
            scale = lodpi.get_scale(owner)
        if scale <= 1.01:
            return
        w = lodpi.scale_int(451, scale, owner)
        self.Width = w
        self._field.Width = lodpi.scale_int(121, scale, owner)
        self._operator.Width = lodpi.scale_int(110, scale, owner)
        val_w = lodpi.scale_int(175, scale, owner)
        self._value_textbox.Width = val_w
        self._value_combobox.Width = val_w
        

    def set_fields(self, rule):

        self._field.SelectedItem = rule.field

        self._operator.SelectedItem = rule.operator

        if self._field.SelectedItem in ("Manga", "Series Complete", "Black And White"):
                self._value_combobox.SelectedItem = rule.value
        else:
            self._value_textbox.Text = rule.value

        
    def field_selection_index_changed(self, sender, e):
        """
        Changes the items in the operator combobox based on what field is selected.
        Also enables or disables the value combobox as required
        """

        if sender.SelectedItem in ("Series Complete", "Black And White"):
            self.set_operator_and_value_items(["is", "is not"], ["Yes", "No", "Unknown"])

        elif sender.SelectedItem is "Manga":
            self.set_operator_and_value_items(["is", "is not"], ["Yes", "Yes (Right to Left)", "No", "Unknown"])

        else:
            if self._value_textbox.Enabled == False:
                self.set_operator_and_value_items(["contains", "does not contain", "greater than", "less than", "is", "is not"])

    
    def set_operator_and_value_items(self, operator_items, value_items=None):
        """
        Sets the items in the operator and value comboboxes and enables and disables the correct controls
        
        If value_items is None then the value textbox will be enabled and the value combobox will be disabled
        """
        if operator_items:
            self._operator.Items.Clear()
            self._operator.Items.AddRange(System.Array[System.String](operator_items))
            self._operator.SelectedIndex = 0

        if value_items is not None:
            self._value_textbox.Enabled = False
            self._value_textbox.Visible = False
            self._value_combobox.Enabled = True
            self._value_combobox.Items.Clear()
            self._value_combobox.Items.AddRange(System.Array[System.String](value_items))
            self._value_combobox.SelectedIndex = 0
            self._value_combobox.Visible = True

        else:
            self._value_textbox.Enabled = True
            self._value_textbox.Visible = True
            self._value_combobox.Enabled = False
            self._value_combobox.Visible = False


    def get_rule(self):
        """Returns an ExcludeRule object containing the rule of this control"""
        if self._value_textbox.Enabled:
            value = self._value_textbox.Text

        else:
            value = self._value_combobox.SelectedItem

        return ExcludeRule(self._field.SelectedItem, self._operator.SelectedItem, value)