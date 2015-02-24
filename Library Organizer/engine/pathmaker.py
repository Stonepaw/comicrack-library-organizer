import clr
import re

clr.AddReference("ComicRack.Engine")
from cYo.Projects.ComicRack.Engine import ComicBook, MangaYesNo, YesNo

from fields import FIELDS, FieldType

class PathMaker(object):
    

    def __init__(self):
        self.book = None

    """This regular expression is built to match all fields formatted in the form
    {prefix<Field(arg)(arg2)>suffix}. It allows for nesting of fields into
    The prefix and suffix parts of the regex. Args are optional depending on
    the field."""
    template_regex = re.compile("{(?P<prefix>[^{}<]*)<(?P<name>[^\d\s(>]*)(?P<args>\d*|(?:\([^)]*\))*)>(?P<suffix>[^{}]*)}")

    def make_folder_path(self, book, folder_template):
        """Creates a folder path with the given folder template by substituting 
           field values into the folder template
        
            Args:
                book: The ComicBook object to make the path for.
                folder_template: the folder template string to use.

            Returns:
                The string of the final folder path.

            Raises:
        """
        self.book = book
        folder_path = ""
        folder_template = folder_template.strip().strip("\\")
        if folder_template:
            pass

    def make_file_name(self, book, file_template, profile):
        """Creates a file name with the given file template by substituting 
           field values into the template
        
            Args:
                book: The ComicBook object to make the name for.
                file_template: the file template string to use.
                profile: The Stonepaw.LibraryOrganizer.Profile object 
                    containing the settings to use to make this file name.

            Returns:
                The string of the final file name.

            Raises:
        """
        #if book is None or not type(book) == ComicBook:
        #    raise ValueError("book has to be an instance of ComicBook")
        #self.book = book
        #file_path = self.book.FileNameWithExtension
        #file_template = file_template.strip().strip("\\")
        #if file_template:
        #    return self.insert_fields_into_template(file_template)
        #return file_name
        self.book = book
        file_name = self.insert_fields_into_template(file_template)
        file_name = file_name.strip()
        file_name = self.replace_illegal_characters(file_name)
        if not file_name:
            return ""
        if self.book.FilePath:
            extension = FileInfo(self.book.FilePath).Extension
        
        return file_name + extension
            
    def insert_fields_into_template(self, template):
        """Replaces fields in the template with the correct field text.
        
        Args:
            template: The string template to use

        Returns:
            The substituted template string

        """
        self.invalid = 0
        #By using a while loop we can recursively insert nested fields. 
        while self.template_regex.search(template):
            if self.invalid == len(self.template_regex.findall(template)):
                break
            else:
                self.invalid = 0
                template = self.template_regex.sub(self.insert_field, template)
        return template

    def insert_field(self, match):
        """Replaces a regex match with the correct text. Returns the original match if something is not valid."""

        match_groups = match.groupdict()
        field_name = match_groups["name"]
        field_args = match_groups["args"]

        #Checking template names
        try:
            field = FIELDS.get_by_template(field_name)
        except KeyError:
            #Template field is not a valid field
            self.invalid +=1
            return match.group(0)
        
        #Get the field text
        text = self.get_field_text(field, field_args)

    def get_field_text(self, field, args_match):
        """Get the text of a field from the book
        
        Args:
            field: The fields.Field object to use
            args_match: The args matched from the regex.

        Returns:
            A string or None if an error occurs or the field isn't valid.
        """
        args = re.findall("\(([^)]*)\)", args_match)

        try:
            if field.type == FieldType.String:
                return self.insert_string_field(self, field)
            if field.type == FieldType.YesNo:
                return self.insert_yes_no_field(self, field, args)
            if field.type == FieldType.ReadPercentage:
                return self.insert_read_percentage(self, args)
            if field.type == FieldType.FirstLetter:
                return self.insert_first_letter(args[0])
        except (AttributeError, ValueError), ex:
            print ex
            return None
        
    def insert_string_field(self, field):
        """Retrieves a string value from the class instance global book and 
            formats it.

        Args:
            field: A string containing the ComicRack internal field name.

        Returns: The formatted unicode string value. If the field is empty
            then it returns an empty string.
        """
        text = getattr(self.book, field)

        if not text or text == -1:
            return ""

        return self.replace_illegal_characters(unicode(text))

    def replace_illegal_characters(self, text):
        """Replaces illegal path characters in a string using the profile's list
            of illegal characters.

        Args:
            text: The unicode string to remove illegal character from.
        """

        #Sort the illegal characters from longest to smallest. This will prevent
        #problems where a phrase containing another illegal character exists.
        #TODO: Try Caching this!
        for illegal_character in sorted(self.profile.IllegalCharacters.keys(), 
                                        key=len, reverse=True):
            text = text.replace(illegal_character, 
                                self.profile.IllegalCharacters[illegal_character])

        return text

    def insert_yes_no_field(self, field, args):
        """Checks a yes no field and returns the args text when the field is the
            correct value.

        Args:
            field: The ComicRack internal name of the field as a string.
            args: A list with one or two items. 
                    First item should be a string of what text to insert when 
                    the value is Yes.
                    The second item (optional) should be a ! or null. ! means 
                    the text is only to be inserted when the value is No.
        
        Returns: The cleaned up user text from the first argument if the field 
            value matches the user's choice or and empty string otherwise. 
        """
        field_value = getattr(self.book, field)
        result = ""
        invert = len(args) == 2 and args[1] == "!"
        if not invert and field_value in (MangaYesNo.Yes,
                                      MangaYesNo.YesAndRightToLeft, 
                                      YesNo.Yes):
            result = args[0]
        elif invert and field_value in (MangaYesNo.No, YesNo.No):
            result = args[0]
        return self.replace_illegal_characters(result)

    def insert_read_percentage(self, args):
        """Returns a user defined string based on minimum or maximum of the 
            book's read percentage.

        Args: 
            args: Required to be a list containing 3 items:
                1. Should be a string insert if the ReadPercentage matches the calculations.
                2. Should be a string containing one of: < > =
                3. Should be the string containing the percentage to match.

        Returns: The formatted user desired text or an empty string.
        """
        text = args[0]
        operator = args[1]
        percent = args[2]
        result = ""
        if operator == "=":
            if self.book.ReadPercentage == int(percent):
                result = text
        elif operator == ">":
            if self.book.ReadPercentage > int(percent):
                result = text
        elif operator == "<":
            if self.book.ReadPercentage < int(percent):
                result = text
        return self.replace_illegal_characters(result)

    def insert_first_letter(self, field):
        """Retrieves the selected field and returns the first letter while ignoring
            pre-programmed initial articles.

            TODO: Let the user choose initial articles to exclude
        Args:
            field: A string containing the ComicRack internal field name.
        
        Returns:
            A single character string or an empty string if the field is empty.
        """
        result = ""
        try:
            f = FIELDS.get_by_CR_field(field)
        except KeyError:
            try:
                field = FIELDS.get_by_template(field).field
            except KeyError:
                try:
                    field = FIELDS.get_by_default_name(field).field
                except KeyError:
                    return "Unable to parse First Letter field"
        property = getattr(self.book, field)
        match_result = re.match(r"(?:(?:the|a|an|de|het|een|die|der|das|des|dem|der|ein|eines|einer|einen|la|le|l'|les|un|une|el|las|los|las|un|una|unos|unas|o|os|um|uma|uns|umas|en|et|il|lo|uno|gli)\s+)?(?P<letter>.).+", property, re.I)
        if match_result:
            result = match_result.group("letter").capitalize()
        print result
        return self.replace_illegal_characters(result)