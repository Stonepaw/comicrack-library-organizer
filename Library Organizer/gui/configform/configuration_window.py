import clr
clr.AddReference("IronPython.Wpf")
clr.AddReference("NLog")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")

from IronPython.Modules import Wpf
from NLog import LogManager
from System import Uri
from System.Diagnostics import Process, ProcessStartInfo
from System.IO import FileInfo, Path
from System.Windows import Visibility, Window
from System.Windows.Controls import Grid
from System.Windows.Media import Colors, SolidColorBrush
from System.Windows.Media.Imaging import BitmapImage

#from codeboxdecorations import LibraryOrganizerArgsDecoration, LibraryOrganizerNameDecoration, LibraryOrganizerPrefixSuffixDecoration
from common import SCRIPTDIRECTORY
from configuration_window_view_models import ConfigurationWindowViewModel
from globals import RUNNER
from insert_view_models import InsertFieldTemplateSelector
from wpfutils import ComparisonConverter

logger = LogManager.GetLogger("ConfigurationWindow")

class ConfigurationWindow(Window):
    """The code behind for the ConfigurationWindow XMAL.

    Most of the window is controled by viewmodels.
    """
    def __init__(self, profiles, global_settings):
        """Creates a new ConfigurationWindow

        Args:
            profiles: The ObservableCollection[Stonepaw.LibraryOrganizer.Profile]
                containing all the profiles.

            global_settings:
                The Stonepaw.LibraryOrganizer.GlobalSettings instance.
        """
        logger.Info("Initializing Configure Form")
        # Release version will have a flatted directory structure
        if RUNNER:
            directory = Path.Combine(SCRIPTDIRECTORY, "resources\icons")
        else:
            directory = SCRIPTDIRECTORY

        self.ViewModel = ConfigurationWindowViewModel(profiles, global_settings)
        self.DataContext = self.ViewModel

        #Due to the fact that some things in wpf are "tricky" in wpf, some
        #things have to be done in the code rather than in the xaml.
        self.Resources.Add("InsertFieldTemplateSelector", 
                           InsertFieldTemplateSelector())
        self.Resources.Add("ComparisonConverter", ComparisonConverter())

        #Insert the images required into the resources. Images have to
        #be inserted this way because of changing directory structure
        #in dev and release.
        logger.Debug("Loading the ConfigurationWindow images")
        try:
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
            b = BitmapImage()
            b.BeginInit()
            b.DecodePixelWidth = 16
            b.UriSource = Uri(Path.Combine(directory, 'tools_32.png'))
            b.EndInit()

            self.Resources.Add("SmallOptionsImage", b)
        except IOError, ex:
            logger.Error(ex)
            raise

        try:
            logger.Info("Loading configure form xaml")

            Wpf.LoadComponent(self, Path.Combine(FileInfo(__file__).DirectoryName, 
                                                 'ConfigurationWindow.xaml'))
        except Exception, ex:
            logger.Error(ex)
            raise

        #self._setup_text_highlighting();

    #def _setup_text_highlighting(self):
    #    """Setups up the text highlighting for the template text boxes
    #    """
    #    logger.Info("Setting up text highlighting")
    #    names = LibraryOrganizerNameDecoration()
    #    names.Brush = SolidColorBrush(Colors.Blue)
    #    self.FileTemplateTextBox.Decorations.Add(names);
    #    self.FolderTemplateBox.Decorations.Add(names);

    #    prefix = LibraryOrganizerPrefixSuffixDecoration()
    #    prefix.Brush = SolidColorBrush(Colors.Teal)
    #    self.FileTemplateTextBox.Decorations.Add(prefix);
    #    self.FolderTemplateBox.Decorations.Add(prefix);

    #    args = LibraryOrganizerArgsDecoration();
    #    args.Brush = SolidColorBrush(Colors.Red);
    #    self.FileTemplateTextBox.Decorations.Add(args);
    #    self.FolderTemplateBox.Decorations.Add(args);

    def _show_profile_inputbox(self, *args):
        """Shows the profile input box when the new profile button is clicked.
        """
        self.ProfileNameInputBox.Text = ""
        self.ProfileNameInput.SetValue(Grid.VisibilityProperty, 
                                       Visibility.Visible)
        self.ProfileNameInputBox.Focus()

    def _close_profile_inputbox(self, *args):
        """Closes the profile input box."""
        self.ProfileNameInput.SetValue(Grid.VisibilityProperty, 
                                       Visibility.Collapsed)

    def _add_illegal_character_clicked(self, *args):
        self.NewIllegalCharacter.Text = ""
        self.NewIllegalCharacter.Focus()

    def _navigate_uri(self, sender, e):
        """Shows a url when a wpf link is clicked."""
        Process.Start(ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True