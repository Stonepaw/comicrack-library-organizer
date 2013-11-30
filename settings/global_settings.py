
from System.IO import File, FileMode
from System.Xml import XmlWriter, XmlWriterSettings

import xml2py

from locommon import VERSION
import xmlserializer


class GlobalSettings(object):

    def __init__(self):
        self.version = VERSION
        self.last_used_profiles = []
        self.generate_localized_months = True
        self.remove_empty_folders = True
        self.excluded_empty_folders = []

        self.month_names = {"1" : "January", 
                           "2" : "February", 
                           "3" : "March", 
                           "4" : "April",
                           "5" : "May", 
                           "6" : "June", 
                           "7" : "July", 
                           "8" :"August", 
                           "9" : "September",
                           "10" : "October", 
                           "11" : "November", 
                           "12" : "December"}

        self.illegal_character_replacements = {"?" : "", 
                                               "/" : "", 
                                               "\\" : "", 
                                               "*" : "", 
                                               ":" : " - ",
                                               "<" : "[", 
                                               ">" : "]", 
                                               "|" : "!", 
                                               "\"" : "'"}
    def save(self, file_path):
        serializer = XmlSerializer()
        xSettings = XmlWriterSettings()
        xSettings.Indent = True
        with XmlWriter.Create(file_path, xSettings) as writer:
            writer.WriteStartElement("GlobalSettings")
            for i in self.__dict__:
                value = getattr(self, i)
                if type(value) is bool:
                    serializer.serialize_bool(value, i, writer)
                elif type(value) is list:
                    serializer.serialize_list(value, i, writer)
                elif type(value) is dict:
                    serializer.serialize_dict(value, i, writer, True)
                elif type(value) is str:
                    writer.WriteElementString(i, value)
            writer.WriteEndElement()

    def load(self, file_path):
        if not File.Exists(file_path):
            return
        with File.Open(file_path, FileMode.Open) as f:
            xml = xml2py.parse(f)
            xmlserializer.object_from_xml2py(self, xml)