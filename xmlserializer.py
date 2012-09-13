
# Rather annoyingly I couldn't get any built in Xml serializer, JSON serializer or pickle to work correctly.
# None seem to be compatible with both .net and python types
# So I had to write my own.
import clr
clr.AddReference("System.Xml")
from System.IO import File, IOException
from System.Collections.ObjectModel import ObservableCollection
from System.Xml import XmlWriter, XmlReader, XmlWriterSettings, XmlDocument, XmlException, XmlConvert
from excluderules import ExcludeRuleCollection, ExcludeGroup, ExcludeRule
from losettings import Profile
from locommon import VERSION

class XmlSerializer(object):
    """Serializes a Profile object to xml."""

    collections_to_write_empty_items = ("IllegalCharacters", "Months")

    def serialize_profiles(self, profiles, last_used, file_path):
        """Serializes a dictionary of Profiles to an xml file.

        Args:
            profiles: A dictionary of Profile objects.
            last_used: A list of names of the last used profiles.
            file_path: The full path of the xml file to write to.
        """
        xSettings = XmlWriterSettings()
        xSettings.Indent = True
        with XmlWriter.Create(file_path, xSettings) as writer:
            writer.WriteStartElement("Profiles")
            writer.WriteAttributeString("LastUsed", ",".join(last_used))
            for profile in profiles:
                self.serialize_profile(profiles[profile], writer)

    def serialize_profile(self, profile, xmlwriter):
        """Serializes a profile to an xml file.

        Args:
            profile: The profile to serialize.
            xmlwriter: An open instance of and XmlWriter.
        """
        xmlwriter.WriteStartElement("Profile")
        xmlwriter.WriteAttributeString("Name", profile.Name)
        xmlwriter.WriteStartAttribute("Version")
        xmlwriter.WriteValue(profile.Version)
        xmlwriter.WriteEndAttribute()
        for attribute_name in profile.__dict__:
            attribute = getattr(profile, attribute_name)
            attribute_type = type(attribute)
            if attribute_type is str and attribute_name != "Name":
                xmlwriter.WriteElementString(attribute_name, attribute)
            elif attribute_type is bool:
                self.serialize_bool(attribute, attribute_name, xmlwriter)
            elif attribute_type is dict:
                self.serialize_dict(attribute, attribute_name, xmlwriter)
            elif attribute_type is list:
                self.serialize_list(attribute, attribute_name, xmlwriter)
            elif attribute_type is ObservableCollection[str]:
                self.serialize_observable_collection(attribute, attribute_name, xmlwriter)
            elif attribute_type is ExcludeRuleCollection:
                self.serialize_exclude_rule_collection(attribute, attribute_name, xmlwriter)
        xmlwriter.WriteEndElement()

    def serialize_bool(self, attribute, attribute_name, xmlwriter):
        """Serializes a bool to xml in the format <AttributeName>true/false</AttriubteName>
        
        Args:
            attribute: the bool to serialize
            attribute_name: the name of the attribute. This will be used as the element name
            xmlwriter: An open XmlWriter instance to write with
        """
        xmlwriter.WriteStartElement(attribute_name)
        xmlwriter.WriteValue(attribute)
        xmlwriter.WriteEndElement()

    def serialize_list(self, attribute, attribute_name, xmlwriter):
        """Writes a list to an xml file. Does not support nested lists.

        Writes the list in the format
        <attribute_name>
            <Item>value</Item>
            <Item>value</Item>
            etc.
        </attribute_name>

        Args:
            attribute: The list to serialize
            attribute_name: The name of the list attribute to write. This will be the root element.
            xmlwriter: An open XmlWriter instance to write with.
        """
        xmlwriter.WriteStartElement(attribute_name)
        for item in attribute:
            if item:
                xmlwriter.WriteElementString("Item", item)
        xmlwriter.WriteEndElement()

    def serialize_dict(self, attribute, attribute_name, xmlwriter):
        """Writes a dictionary to an xml file. Does not support nested Dictionaries.

        Writes in the dictonary in the format:
        <attribute_name>
            <Item Name="attribute_name key" Value="attribute_name value" />
            <Item Name="attribute_name key" Value="attribute_name value" />
            etc.
        </attribute_name>

        Args:
            attribute: The dict to serialize
            attribute_name: The name of the dict attribute to write. This will be the root element.
            xmlwriter: An open XmlWriter instance to write with.
        """
        write_empty = False
        if attribute_name in self.collections_to_write_empty_items:
            write_empty = True
        xmlwriter.WriteStartElement(attribute_name)
        for key in attribute:
            if attribute[key] or write_empty:
                xmlwriter.WriteStartElement("Item")
                xmlwriter.WriteStartAttribute("Name")
                xmlwriter.WriteValue(key)
                xmlwriter.WriteEndAttribute()
                xmlwriter.WriteStartAttribute("Value")
                xmlwriter.WriteValue(attribute[key])
                xmlwriter.WriteEndAttribute()
                xmlwriter.WriteEndElement()
        xmlwriter.WriteEndElement()

    def serialize_observable_collection(self, attribute, attribute_name, xmlwriter):
        """Writes an observable collection to an xml file. Does not support nested observable collections

        Writes in the format:
        <attribute_name>
            <Item>Value</Item>
            <Item>Value</Item>
            etc.
        </attribute_name>

        Args:
            attribute: The observable collection to serialize
            attribute_name: The name of the observable collection attribute to write. 
                            This will be the root element.
            xmlwriter: An open XmlWriter instance to write with.
        """
        write_empty = False
        if attribute_name in self.collections_to_write_empty_items:
            write_empty = True
        xmlwriter.WriteStartElement(attribute_name)
        for item in attribute:
            if item or write_empty:
                xmlwriter.WriteElementString("Item", item)
        xmlwriter.WriteEndElement()

    def serialize_exclude_rule_collection(self, exclude_rule_collection, attribute_name, xmlwriter):
        """Serializes an ExcludeRuleCollection to xml.

        Args:
            exclude_rule_collection: The ExcludeRuleCollection to serialize.
            attribute_name: The name of the ExcludeRuleCollection. This will be the root element.
            xmlwriter: An open XmlWriter instance to write with.        
        """
        xmlwriter.WriteStartElement(attribute_name)
        xmlwriter.WriteAttributeString("Operator", exclude_rule_collection.operator)
        xmlwriter.WriteAttributeString("ExcludeMode", exclude_rule_collection.mode)
        for rule in exclude_rule_collection.rules:
            if type(rule) is ExcludeRule:
                self.serialize_exclude_rule(rule, xmlwriter)
            elif type(rule) is ExcludeGroup:
                self.serialize_exclude_rule_group(rule, xmlwriter)
        xmlwriter.WriteEndElement()

    def serialize_exclude_rule_group(self, exclude_rule_group, xmlwriter):
        """Serializes an ExcludeGroup to xml.

        Args:
            exclude_rule_group: The ExcludeGroup to serialize.
            xmlwriter: An open XmlWriter instance to write with.        
        """
        xmlwriter.WriteStartElement("ExcludeGroup")
        xmlwriter.WriteAttributeString("Operator", exclude_rule_group.operator)
        xmlwriter.WriteStartAttribute("Invert")
        xmlwriter.WriteValue(exclude_rule_group.invert)
        xmlwriter.WriteEndAttribute()
        for rule in exclude_rule_group.rules:
            if type(rule) is ExcludeRule:
                self.serialize_exclude_rule(rule, xmlwriter)
            elif type(rule) is ExcludeGroup:
                self.serialize_exclude_rule_group(rule, xmlwriter)
        xmlwriter.WriteEndElement()

    def serialize_exclude_rule(self, exclude_rule, xmlwriter):
        """Serializes an ExcludeRule to xml.

        Args:
            exclude_rule: The ExcludeRule to serialize.
            xmlwriter: An open XmlWriter instance to write with.        
        """
        xmlwriter.WriteStartElement("ExcludeRule")
        xmlwriter.WriteAttributeString("Field", exclude_rule.field)
        xmlwriter.WriteAttributeString("Operator", exclude_rule.operator)
        xmlwriter.WriteAttributeString("Value", exclude_rule.value)
        xmlwriter.WriteStartAttribute("Invert")
        xmlwriter.WriteValue(exclude_rule.invert)
        xmlwriter.WriteEndAttribute()
        xmlwriter.WriteEndElement()


class XmlDeserializer(object):
    """Deserializes a profile or group of profiles from an xml file"""

    def deserialize_profiles(self, file_path):
        """Deserializes a set of profiles or single profile from a file.

        Args:
            file_path: The full file path to the xml file from with to load profiles

        Returns:
            A dict containing the returned profiles as the values with keys of the profile names.
            and a list containing the last used profiles.

            Or None when the profile(s) could not be loaded.

        Raises:
            XmlDeserializerException: when something goes wrong loading or parsing the file
        """
        if not File.Exists(file_path):
            raise XmlDeserializerException("Unable to deserialize profile(s) from the path: %s because the file does not exist" % (file_path))

        profiles_xml_document = XmlDocument()
        try:
            profiles_xml_document.Load(file_path)
        except (XmlException, IOException) as e:
            raise XmlDeserializerException(e.Message)
        
        if profiles_xml_document.DocumentElement.Name == "Profiles":
            profile_nodes = profiles_xml_document.SelectNodes("Profiles/Profile")
        elif profiles_xml_document.DocumentElement.Name == "Profile":
            profile_nodes = profiles_xml_document.SelectNodes("Profile")
        else:
            raise XmlDeserializerException("%s is not a vaid Library Organizer file." % file_path)
        
        profiles = {}
        for profile_node in profile_nodes:
            profile = self.deserialize_profile(profile_node)
            profiles[profile.Name] = profile

        rootnode = profiles_xml_document.DocumentElement
        if rootnode.HasAttribute("LastUsed"):
            last_used = rootnode.Attributes["LastUsed"].Value.split(",")
        else:
            last_used = []

        return profiles, last_used
        

    def deserialize_profile(self, profile_node):
        """Loads attributes from a profiles xml node.

        Args: 
            profile_node: The xml to load the profile from.

        Returns:
            A Profile instance that was loaded from the xml
        """
        #Text vars
        profile = Profile()
        profile.Name = profile_node.Attributes["Name"].Value

        if "Version" in profile_node.Attributes:
            profile.Version = float(profile_node.Attributes["Version"].Value)
        else:
            profile.Version = VERSION

        for attribute_name in profile.__dict__:

            if attribute_name in ("Version", "Name"):
                continue

            attribute_type = type(getattr(profile, attribute_name))

            if attribute_type is str:
                attribute = self.deserialize_string(profile_node, attribute_name)
                if attribute is not None:
                    setattr(profile, attribute_name, attribute)
            elif attribute_type is bool:
                attribute = self.deserialize_bool(profile_node, attribute_name)
                if attribute is not None:
                    setattr(profile, attribute_name, attribute)
            elif attribute_type is list:
                attribute = self.deserialize_list(profile_node, attribute_name)
                if attribute is not None:
                    getattr(profile, attribute_name).extend(attribute)
            elif attribute_type is dict:
                attribute = self.deserialize_dict(profile_node, attribute_name)
                if attribute is not None:
                    d = getattr(profile, attribute_name)
                    for i in attribute:
                        d[i] = attribute[i]
            elif attribute_type is ObservableCollection[str]:
                attribute = self.deserialize_observable_collection(profile_node, attribute_name)
                if attribute is not None:
                    collection = getattr(profile, attribute_name)
                    for i in attribute:
                        collection.Add(i)
            elif attribute_type is ExcludeRuleCollection:
                attribute = self.deserialize_exclude_rule_collection(profile_node, attribute_name)
                if attribute is not None:
                    setattr(profile, attribute_name, attribute)

        profile.update()
        return profile

    def deserialize_string(self, profile_node, attribute_name):
        """Deserializes a string from a profile xml node.

        Args:
            profile_node: The profile node to deserialize the string from.
            attribute_name: The name of the attribute attribute. This should also 
                            be the element name to deserialize from.

        Returns:
            The deserialized string or None if the element could not be found
        """
        if profile_node.SelectSingleNode(attribute_name) is not None:
            return profile_node.SelectSingleNode(attribute_name).InnerText
        return None

    def deserialize_bool(self, profile_node, attribute_name):
        """Deserializes a boolean from a profile xml node

        Args:
            profile_node: The profile node to deserialize the list from.
            attribute_name: The name of the boolean attribute. This should 
                            be the element name to deserialize from.

        Returns:
            The deserialized boolean or None if the element could not be found.
        """
        if profile_node.SelectSingleNode(attribute_name) is not None:
            return XmlConvert.ToBoolean(profile_node.SelectSingleNode(attribute_name).InnerText)
        return None

    def deserialize_list(self, profile_node, attribute_name):
        """Deserializes a list from a profile xml node

        Args:
            profile_node: The profile node to deserialize the list from.
            attribute_name: The name of the list attribute. This should
                            be the element name to deserialize from.

        Returns:
            The deserialized list or None if the element could not be found.
        """
        list_item_nodes = profile_node.SelectNodes(attribute_name + "/Item")
        if list_item_nodes.Count > 0:
            return [item.InnerText for item in list_item_nodes]
        return None

    def deserialize_dict(self, profile_node, attribute_name):
        """Deserializes a dict from a profile xml node

        Args:
            profile_node: The profile node to deserialize the dict from.
            attribute_name: The name of the dict attribute. This should
                            be the element name to deserialize from.

        Returns:
            The deserialized dict or None if the element could not be found.
        """
        dict_nodes = profile_node.SelectNodes(attribute_name + "/Item")
        if dict_nodes.Count > 0:
            return {node.Attributes["Name"].Value : node.Attributes["Value"].Value 
                    for node in dict_nodes if node.Attributes.Count == 2}
        return None

    def deserialize_observable_collection(self, profile_node, attribute_name):
        """Deserializes an observable collection[str] from a profile xml node

        Args:
            profile_node: The profile node to deserialize the observable collection from.
            attribute_name: The name of the observable collection attribute. This should
                            be the element name to deserialize from.

        Returns:
            The deserialized observable collection or None if the element could not be found.
        """
        collection_item_nodes = profile_node.SelectNodes(attribute_name + "/Item")
        if collection_item_nodes.Count > 0:
            return ObservableCollection[str]([item.InnerText for item in collection_item_nodes])
        return None

    def deserialize_exclude_rule_collection(self, profile_node, attribute_name):
        """Deserializes an ExcludeRuleCollection from a profile xml node.

        Args:
            profile_node: The profile node to deserialize the exclude rule collection from.
            attribute_name: The name of the exclude rule collection attribute. This should
                            be the element name to deserialize from.

        Returns:
            The deserialized ExcludeRuleCollection or None if the element could not be found.
        """
        exclude_rules_node = profile_node.SelectSingleNode(attribute_name)
        if exclude_rules_node is not None:
            exclude_rules = ExcludeRuleCollection(operator=exclude_rules_node.Attributes["Operator"].Value, 
                                                 mode=exclude_rules_node.Attributes["ExcludeMode"].Value)
            for node in exclude_rules_node.ChildNodes:
                if node.Name == "ExcludeRule":
                    exclude_rules.add_rule(self.deserialize_exclude_rule(node)) 
                elif node.Name == "ExcludeGroup":
                    exclude_rules.add_rule(self.deserialize_exclude_group(node))
            return exclude_rules
        return None

    def deserialize_exclude_group(self, exclude_group_node):
        """Deserializes an ExcludeGroup from a exclude group xml node.

        Args:
            exclude_group_node: The exclude group node to deserialize the exclude group from.

        Returns:
            The deserialized ExcludeGroup.
        """
        invert = exclude_group_node.Attributes["Invert"].Value
        if invert is None: invert = "false"
        invert = XmlConvert.ToBoolean(invert)
        group = ExcludeGroup(operator=exclude_group_node.Attributes["Operator"].Value,
                             invert=invert)
        for node in exclude_group_node.ChildNodes:
            if node.Name == "ExcludeRule":
                group.add_rule(self.deserialize_exclude_rule(node)) 
            elif node.Name == "ExcludeGroup":
                group.add_rule(self.deserialize_exclude_group(node))
        return group

    def deserialize_exclude_rule(self, exclude_rule_node):
        """Deserializes an ExcludeRule from a exclude rule xml node.

        Args:
            exclude_rule_node: The exclude rule node to deserialize the exclude rule from.

        Returns:
            The deserialized ExcludeRule.
        """
        invert = exclude_rule_node.Attributes["Invert"].Value
        if invert is None: invert = "false"
        invert = XmlConvert.ToBoolean(invert)
        return ExcludeRule(field=exclude_rule_node.Attributes["Field"].Value, 
                           operator=exclude_rule_node.Attributes["Operator"].Value, 
                           value=exclude_rule_node.Attributes["Value"].Value, invert=invert)


class XmlDeserializerException(Exception):
    pass
