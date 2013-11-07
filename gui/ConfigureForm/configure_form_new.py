import clr
clr.AddReference("IronPython.Wpf")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("System.Windows.Forms")

from IronPython.Modules import Wpf
from System import Uri
from System.Diagnostics import Process, ProcessStartInfo
from System.IO import FileInfo, Path
from System.Windows import Visibility, Window
from System.Windows.Controls import Grid
from System.Windows.Media import Colors, SolidColorBrush
from System.Windows.Media.Imaging import BitmapImage

from codeboxdecorations import LibraryOrganizerArgsDecoration, LibraryOrganizerNameDecoration, LibraryOrganizerPrefixSuffixDecoration
from config_form_view_models import ConfigureFormViewModel
from insert_view_models import InsertFieldTemplateSelector
from locommon import SCRIPTDIRECTORY
from wpfutils import ComparisonConverter

RUNNER = False

class ConfigureForm(Window):
    def __init__(self, profiles, global_settings):

        # Release version will have a flatted directory structure
        if RUNNER:
            directory = Path.Combine(SCRIPTDIRECTORY, "resources\icons")
        else:
            directory = SCRIPTDIRECTORY

        self.ViewModel = ConfigureFormViewModel(profiles, global_settings)
        self.DataContext = self.ViewModel

        self.Resources.Add("InsertFieldTemplateSelector", 
                           InsertFieldTemplateSelector())

        """ Insert the images required into the resources. Images have to
        be inserted this way because of changing directory structure
        in dev and release """
        self.Resources.Add("HomeImage", BitmapImage(
                Uri(Path.Combine(directory, 'home_32.png'))))
        self.Resources.Add("FolderImage", BitmapImage(
                Uri(Path.Combine(directory, 'folder_32.png'))))
        self.Resources.Add("FileImage", BitmapImage(
                Uri(Path.Combine(directory, 'page_text_32.png'))))
        self.Resources.Add("RulesImage", BitmapImage(
                Uri(Path.Combine(directory, 'chart_32.png'))))
        self.Resources.Add("OptionsImage", BitmapImage(
                Uri(Path.Combine(directory, 'tools_32.png'))))
        self.Resources.Add("ComparisonConverter", ComparisonConverter())
        Wpf.LoadComponent(self, Path.Combine(FileInfo(__file__).DirectoryName, 
                                             'ConfigureFormNew.xaml'))



        self.setup_text_highlighting();

    def setup_text_highlighting(self):
        """
        Setups up the text highlighting for the template text boxes
        """
        names = LibraryOrganizerNameDecoration()
        names.Brush = SolidColorBrush(Colors.Blue)
        self.FileTemplateTextBox.Decorations.Add(names);
        self.FolderTemplateBox.Decorations.Add(names);

        prefix = LibraryOrganizerPrefixSuffixDecoration()
        prefix.Brush = SolidColorBrush(Colors.Teal)
        self.FileTemplateTextBox.Decorations.Add(prefix);
        self.FolderTemplateBox.Decorations.Add(prefix);

        args = LibraryOrganizerArgsDecoration();
        args.Brush = SolidColorBrush(Colors.Red);
        self.FileTemplateTextBox.Decorations.Add(args);
        self.FolderTemplateBox.Decorations.Add(args);

    def new_profile_clicked(self, *args):
        self.ProfileNameInputBox.Text = ""
        self.ProfileNameInput.SetValue(Grid.VisibilityProperty, 
                                       Visibility.Visible)
        self.ProfileNameInputBox.Focus()

    def close_inputbox(self, *args):
        self.ProfileNameInput.SetValue(Grid.VisibilityProperty, 
                                       Visibility.Collapsed)

    def add_illegal_character_clicked(self, *args):
        self.NewIllegalCharacter.Text = ""
        self.NewIllegalCharacter.Focus()

    def navigate_uri(self, sender, e):
        Process.Start(ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True