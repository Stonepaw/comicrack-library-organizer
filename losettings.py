"""
losettings.py

Contains a class for profiles and methods to save and load them from xml files.

Author: Stonepaw

Version 2.0

    Rewrote pretty much everything. Much more modular and requires no maintence when a new attribute is added.
    No longer fully supports profiles from 1.6 and earlier.

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

clr.AddReference("System.Xml")

from System import Convert
from System.IO import File, StreamReader, StreamWriter
from System.Xml import XmlDocument, XmlWriter, XmlWriterSettings

from System.Windows.Forms import MessageBox, MessageBoxIcon, MessageBoxButtons

from locommon import ExcludeRule, ExcludeGroup, Mode, PROFILEFILE, VERSION



class Profile:
    """This class contains all the variables for a profile.
    Use save_to_xml to save the profile to a xml file.
    Use load_from_xml to load the profile from the file.

    Anytime a new variable is added it will automatically be save and loaded.
    """
    def __init__(self):
        
        self.Version = 0

        self.FolderTemplate = ""
        self.BaseFolder = ""
        self.FileTemplate = ""
        self.Name = ""
        self.EmptyFolder = ""

        self.EmptyData = {}
        
        self.Postfix = {}

        self.Prefix = {}

        self.Seperator = {}

        self.IllegalCharacters = {"?" : "", "/" : "", "\\" : "", "*" : "", ":" : " - ", "<" : "[", ">" : "]", "|" : "!", "\"" : "'"}

        self.Months = {1 : "January", 2 : "February", 3 : "March", 4 : "April", 5 : "May", 6 : "June", 7 : "July", 8 :"August", 9 : "September", 10 : "October",
                       11 : "November", 12 : "December", 13 : "Spring", 14 : "Summer", 15 : "Fall", 16 : "Winter"}

        self.TextBox = {}
                
        self.UseFolder = True
        
        self.UseFileName = True
        
        self.ExcludeFolders = []
    
        self.DontAskWhenMultiOne = True
        
        self.ExcludeRules = []
    
        self.ExcludeOperator = "Any"
        
        self.RemoveEmptyFolder = True
        self.ExcludedEmptyFolder = []
        
        self.MoveFileless = False       
        self.FilelessFormat = ".jpg"
        
        self.ExcludeMode = "Do not"
        

        self.FailEmptyValues = False
        self.MoveFailed = False
        self.FailedFolder = ""
        self.FailedFields = []

        self.Mode = Mode.Move

        self.CopyMode = True

        self.AutoSpaceFields = True

        self.ReplaceMultipleSpaces = True

        self.CopyReadPercentage = True


    def duplicate(self):
        """Returns a duplicate of the profile instance."""
        duplicate = Profile()
        
        for i in self.__dict__:
            if type(getattr(self, i)) is dict:
                setattr(duplicate, i, getattr(self, i).copy())
            else:
                setattr(duplicate, i, getattr(self, i))

        return duplicate


    def update(self):
        if self.Version < 2.0:
            if self.Mode is "Test":
                self.Mode = "Simulate"

            replacements = {"Language" : "LanguageISO", "Format" : "ShadowFormat", "Count" : "ShadowCount", "Number" : "ShadowNumber", "Series" : "ShadowSeries",
                            "Title" : "ShadowTitle", "Volume" : "ShadowVolume", "Year" : "ShadowYear"}

            for key in self.EmptyData.keys():
                if key in replacements:
                    self.EmptyData[replacements[key]] = self.EmptyData[key]
                    del(self.EmptyData[key])

            insert_control_replacements = {"SeriesComplete" : "Series Complete", "Read" : "Read Percentage", "FirstLetter" : "First Letter", "AgeRating" : "Age Rating",
                                    "AlternateSeriesMulti" : "Alternate Series Multi", "MonthNumber" : "Month Number", "AlternateNumber" : "Alternate Number",
                                    "StartMonth" : "Start Month", "AlternateSeries" : "Alternate Series", "ScanInformation" : "Scan Information", "StartYear" : "Start Year",
                                    "AlternateCount" : "Alternate Count"}
            for key in insert_control_replacements :
                if key in self.TextBox.keys():
                    self.TextBox[insert_control_replacements[key]] = self.TextBox[key]
                    del(self.TextBox[key])

                if key in self.Prefix.keys():
                    self.Prefix[insert_control_replacements[key]] = self.Prefix[key]
                    del(self.Prefix[key])

                if key in self.Postfix.keys():
                    self.Postfix[insert_control_replacements[key]] = self.Postfix[key]
                    del(self.Postfix[key])

                if key in self.Seperator.keys():
                    self.Seperator[insert_control_replacements[key]] = self.Seperator[key]
                    del(self.Seperator[key])

        self.Version = VERSION

                
    def save_to_xml(self, xwriter):
        """
        To save this profile intance to xml file using a XmlWriter.
        xwriter->should be a XmlWriter instance.
        """

        xwriter.WriteStartElement("Profile")
        xwriter.WriteAttributeString("Name", self.Name)
        xwriter.WriteStartAttribute("Version")
        xwriter.WriteValue(self.Version)
        xwriter.WriteEndAttribute()

        for var_name in self.__dict__:
            var_type = type(getattr(self, var_name))

            if var_type is str and var_name != "Name":
                self.write_string_to_xml(var_name, xwriter)

            elif var_type is bool:
                self.write_bool_to_xml(var_name, xwriter)

            elif var_type is dict:
                self.write_dict_to_xml(var_name, xwriter)

            elif var_type is list and var_name != "ExcludeRules":
                self.write_list_to_xml(var_name, xwriter)

        xwriter.WriteStartElement("ExcludeRules")
        xwriter.WriteAttributeString("Operator", self.ExcludeOperator)
        xwriter.WriteAttributeString("ExcludeMode", self.ExcludeMode)
        for rule in self.ExcludeRules:
            if rule:
                rule.save_xml(xwriter)
        xwriter.WriteEndElement()
        
        xwriter.WriteEndElement()

    
    def write_dict_to_xml(self, attribute_name, xmlwriter, write_empty=False):
        """Writes a dictionary to an xml file in the form of
        <attribute_name>
            <Item Name="attribute_name key" Value="attribute_name value" />
            <Item Name="attribute_name key" Value="attribute_name value" />
            etc.
        </attribute_name>

        attribute_name->The name of the dictonary attribute to write.
        xmlwriter->The xml writer to write with.
        write_empty->A bool of whether to write empty values to the xml file. Default is don't write them.
        """
        if attribute_name in ("IllegalCharacters", "Months"):
            write_empty = True
        dictionary = getattr(self, attribute_name)
        xmlwriter.WriteStartElement(attribute_name)
        for key in dictionary:
            if dictionary[key] or write_empty:
                xmlwriter.WriteStartElement("Item")
                xmlwriter.WriteStartAttribute("Name")
                xmlwriter.WriteValue(key)
                xmlwriter.WriteEndAttribute()
                xmlwriter.WriteStartAttribute("Value")
                xmlwriter.WriteValue(dictionary[key])
                xmlwriter.WriteEndAttribute()
                xmlwriter.WriteEndElement()
        xmlwriter.WriteEndElement()


    def write_list_to_xml(self, attribute_name, xmlwriter, write_empty=False):
        """Writes a list to an xml file in the form of
        <attribute_name>
            <Item>value</Item>
            <Item>value</Item>
            etc.
        </attribute_name>

        attribute_name->The name of the list attribute to write.
        xmlwriter->The xml writer to write with.
        write_empty->A bool of whether to write empty values to the xml file. Default is don't write them.
        """
        attribute_list = getattr(self, attribute_name)
        xmlwriter.WriteStartElement(attribute_name)
        for item in attribute_list:
            if item or write_empty:
                xmlwriter.WriteElementString("Item", item)
        xmlwriter.WriteEndElement()


    def write_string_to_xml(self, attribute_name, xmlwriter, write_empty=True):
        """Writes a string to an xml file in the form of
        <attribute_name>string</attribute_name>

        attribute_name->The name of the string attribute to write.
        xmlwriter->The xml writer to write with.
        write_empty->A bool of whether to write empty strings to the xml file. Default is write empty strings.
        """
        string = getattr(self, attribute_name)
        if string or write_empty:
            xmlwriter.WriteElementString(attribute_name, string)


    def write_bool_to_xml(self, attribute_name, xmlwriter):
        """Writes a boolean to an xml file in the form of
        <attribute_name>true/false</attribute_name>

        attribute_name->The name of the attribute to write.
        xmlwriter->The xml writer to write with.
        """
        xmlwriter.WriteStartElement(attribute_name)
        xmlwriter.WriteValue(getattr(self, attribute_name))
        xmlwriter.WriteEndElement()


    def load_from_xml(self, Xml):
        """Loads the profile instance from the Xml.
        
        Xml->should be a XmlNode/XmlDocument containing a profile node.
        """
        try:
            #Text vars
            self.Name = Xml.Attributes["Name"].Value

            if "Version" in Xml.Attributes:
                self.Version = float(Xml.Attributes["Version"].Value)

            for var_name in self.__dict__:
                if type(getattr(self,var_name)) is str:
                    self.load_text_from_xml(Xml, var_name)


                elif type(getattr(self,var_name)) is bool:
                    self.load_bool_from_xml(Xml, var_name)


                elif type(getattr(self, var_name)) is list and var_name != "ExcludeRules":
                    self.load_list_from_xml(Xml, var_name)

                elif type(getattr(self, var_name)) is dict:
                    self.load_dict_from_xml(Xml, var_name)

            #Exclude Rules
            exclude_rules_node = Xml.SelectSingleNode("ExcludeRules")
            if exclude_rules_node is not None:
                self.ExcludeOperator = exclude_rules_node.Attributes["Operator"].Value

                self.ExcludeMode = exclude_rules_node.Attributes["ExcludeMode"].Value

                for node in exclude_rules_node.ChildNodes:
                    if node.Name == "ExcludeRule":
                        try:
                            rule = ExcludeRule(node.Attributes["Field"].Value, node.Attributes["Operator"].Value, node.Attributes["Value"].Value)
                        except AttributeError:
                            rule = ExcludeRule(node.Attributes["Field"].Value, node.Attributes["Operator"].Value, node.Attributes["Text"].Value)

                        self.ExcludeRules.append(rule)
    
                    elif node.Name == "ExcludeGroup":
                        group = ExcludeGroup(node.Attributes["Operator"].Value)
                        group.load_from_xml(node)
                        self.ExcludeRules.append(group)

            self.update()

        except Exception, ex:
            print ex
            return False


    def load_text_from_xml(self, xmldoc, name):
        """Loads a string with a specified node name from an XmlDocument and saves it to the attribute. The string should be saved as:
        <name>string</name>

        xmldoc->The XmlDocment to load from.
        name->The attribute to save to and the root node name to load the string from."""
        if xmldoc.SelectSingleNode(name) is not None:
            setattr(self, name, xmldoc.SelectSingleNode(name).InnerText)


    def load_bool_from_xml(self, xmldoc, name):
        """Loads a bool with a specified node name from an XmlDocument and saves it to the attribute. The bool should be saved as:
        <name>true/false</name>

        xmldoc->The XmlDocment to load from.
        name->The attribute to save to and the root node name to load the bool from."""
        if xmldoc.SelectSingleNode(name) is not None:
            setattr(self, name, Convert.ToBoolean(xmldoc.SelectSingleNode(name).InnerText))


    def load_list_from_xml(self, xmldoc, name):
        """Loads a list with a specified node name from an XmlDocument and saves it to the attribute. The list should be saved as:
        <name>
            <Item>list value</Item>
        </name>

        xmldoc->The XmlDocment to load from.
        name->The attribute to save to and the root node name to load the list from."""
        nodes = xmldoc.SelectNodes(name + "/Item")
        if nodes.Count > 0:
            setattr(self, name, [item.InnerText for item in nodes])


    def load_dict_from_xml(self, xmldoc, name):
        """Loads a dict with a specified node name from an XmlDocument and saves it to the attribute. The dict should be saved as:
        <name>
            <Item Name="key" Value="value" />
        </name>

        xmldoc->The XmlDocment to load from.
        name->The attribute to save to and the root node name to load the dict from."""
        nodes = xmldoc.SelectNodes(name + "/Item")
        if nodes.Count > 0:
            dictionary = getattr(self, name)
            for node in nodes:
                if node.Attributes.Count == 2:
                    if name == "Months":
                        dictionary[int(node.Attributes["Name"].Value)] = node.Attributes["Value"].Value
                    else:
                        dictionary[node.Attributes["Name"].Value] = node.Attributes["Value"].Value



def load_profiles(file_path):
    """
    Load profiles from a xml file. If no profiles are found it creates a blank profile.
    file_path->The absolute path to the profile file

    Returns a dict of the found profiles and a list of the lastused profile(s)
    """
    profiles, lastused = load_profiles_from_file(file_path)

    if len(profiles) == 0:
        #Just in case
        profiles["Default"] = Profile()
        profiles["Default"].Name = "Default"
        #Some default templates
        profiles["Default"].FileTemplate = "{<series>}{ Vol.<volume>}{ #<number2>}{ (of <count2>)}{ ({<month>, }<year>)}"
        profiles["Default"].FolderTemplate = "{<publisher>}\{<imprint>}\{<series>}{ (<startyear>{ <format>})}"
        
    if not lastused:
        lastused = [profiles.keys()[0]]
       
    return profiles, lastused     


def load_profiles_from_file(file_path):
    """
    Loads profiles from a file.
    
    file_path->The absolute path the xml file

    Returns a dict of the profiles
    """
    profiles = {}

    lastused = ""

    if File.Exists(file_path):
        try:
            with StreamReader(file_path) as xmlfile:
                xmldoc = XmlDocument()
                xmldoc.Load(xmlfile)

            if xmldoc.DocumentElement.Name == "Profiles":
                nodes = xmldoc.SelectNodes("Profiles/Profile")
            #Individual exported profiles are saved with the document element as Profile
            elif xmldoc.DocumentElement.Name == "Profile":
                nodes = xmldoc.SelectNodes("Profile")

            #Changed from 1.7 to 2.0 to use Profiles/Profile instead of Settings/Setting
            elif xmldoc.DocumentElement.Name == "Settings":
                nodes = xmldoc.SelectNodes("Settings/Setting")
            elif xmldoc.DocumentElement.Name == "Setting":
                nodes = xmldoc.SelectNodes("Setting")

            #No valid root elements
            else:
                MessageBox.Show(file_path + " is not a valid Library Organizer profile file.", "Not a valid profile file", MessageBoxButtons.OK, MessageBoxIcon.Error)
                return profiles, lastused

            if nodes.Count > 0:
                for node in nodes:                    
                    profile = Profile()
                    profile.Name = node.Attributes["Name"].Value
                    result = profile.load_from_xml(node)

                    #Error loading the profile
                    if result == False:
                        MessageBox.Show("An error occured loading the profile " + profile.Name + ". That profile has been skipped.")

                    else:
                        profiles[profile.Name] = profile


            #Load the last used profile
            rootnode = xmldoc.DocumentElement
            if rootnode.HasAttribute("LastUsed"):
                lastused = rootnode.Attributes["LastUsed"].Value.split(",")

        except Exception, ex:
            MessageBox.Show("Something seems to have gone wrong loading the xml file.\n\nThe error was:\n" + str(ex), "Error loading file", MessageBoxButtons.OK, MessageBoxIcon.Error)

    return profiles, lastused


def import_profiles(file_path):
    """
    Load profiles from a xml file. If no profiles are found it returns an empty dict.
    file_path->The absolute path to the profile file

    Returns a dict of the found profiles.
    """
    profiles, lastused = load_profiles_from_file(file_path)

    return profiles


def save_profiles(file_path, profiles, lastused=""):
    """
    Saves the profiles to an xml file.

    settings_file: The complete file path of the file to save to.
    profiles: a dict of profile objects.
    lastused: a string containing the last used profile.
    """
    try:
        xSettings = XmlWriterSettings()
        xSettings.Indent = True
        with XmlWriter.Create(file_path, xSettings) as writer:
            writer.WriteStartElement("Profiles")
            if lastused:
                writer.WriteAttributeString("LastUsed", ",".join(lastused))
            for profile in profiles:
                profiles[profile].save_to_xml(writer)
            writer.WriteEndElement()
    except Exception, ex:
        MessageBox.Show("An error occured writing the settings file. The error was:\n\n" + ex.message, "Error saving settings file", MessageBoxButtons.OK, MessageBoxIcon.Error)


def save_profile(file_path, profile):
    """
    Saves a single profile to an xml file.

    settings_file: The complete file path of the file to save to.
    profile: a Profile object.
    """
    try:
        xSettings = XmlWriterSettings()
        xSettings.Indent = True
        with XmlWriter.Create(file_path, xSettings) as writer:
            profile.save_to_xml(writer)
    except Exception, ex:
        MessageBox.Show("An error occured writing the settings file. The error was:\n\n" + ex.message, "Error saving settings file", MessageBoxButtons.OK, MessageBoxIcon.Error)


def save_last_used(file_path, lastused):
    "Saves the lastused profiles to the xml file."""
    x = XmlDocument()
    x.Load(file_path)
    x.DocumentElement.SetAttribute("LastUsed", ",".join(lastused))
    x.Save(file_path)