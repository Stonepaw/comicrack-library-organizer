using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LibraryOrganizer.Controls
{
    internal static class FolderBrowserUtil
    {
        public static string PromptForFolderPath(IWin32Window owner)
        {
            var openFolderDialog = new FolderBrowserDialog();

            if (
                openFolderDialog.ShowDialog(owner) == DialogResult.OK
                && openFolderDialog.SelectedPath != null
            )
            {
                return openFolderDialog.SelectedPath;
            }

            return null;
        }
    }
}
