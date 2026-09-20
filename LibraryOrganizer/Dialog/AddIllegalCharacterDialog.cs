using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using LibraryOrganizer.Data;

namespace LibraryOrganizer.Dialog
{
    public partial class AddIllegalCharacterDialog : Form
    {
        private readonly HashSet<char> _illegalCharacterReplacements;

        public AddIllegalCharacterDialog(
            List<IllegalCharacterReplacement> illegalCharacterReplacements
        )
        {
            InitializeComponent();

            _illegalCharacterReplacements = new HashSet<char>(
                illegalCharacterReplacements.Select(illegalCharacterReplacement =>
                    illegalCharacterReplacement.Character
                )
            );

            errorProvider.SetIconAlignment(textBox, ErrorIconAlignment.MiddleRight);
            errorProvider.SetIconPadding(textBox, -20);
        }

        public char GetCharacter()
        {
            return textBox.Text[0];
        }

        private void ok_Click(object sender, EventArgs e)
        {
            if (ValidateTextContent())
            {
                DialogResult = DialogResult.OK;
            }
            else
            {
                DialogResult = DialogResult.None;
            }
        }

        private void textBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateTextContent();
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            ValidateTextContent();
        }

        private bool ValidateTextContent()
        {
            if (textBox.Text.Length == 0 || textBox.Text.Length > 1)
            {
                errorProvider.SetError(textBox, "A single character is required");
                return false;
            }
            else if (_illegalCharacterReplacements.Contains(textBox.Text[0]))
            {
                errorProvider.SetError(textBox, "Already created");
                return false;
            }
            else
            {
                errorProvider.SetError(textBox, string.Empty);
                return true;
            }
        }
    }
}
