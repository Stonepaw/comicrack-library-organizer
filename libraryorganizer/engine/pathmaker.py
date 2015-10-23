"""
pathmaker.py

This contains the class and methods to create the path from the template

Copyright 2010 Stonepaw

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
import re

clr.AddReferenceByPartialName('ComicRack.Engine')

import System
from System import Func
from System.IO import Path, FileInfo

from cYo.Projects.ComicRack.Engine import MangaYesNo, YesNo

from loforms import MultiValueSelectionFormArgs, MultiValueSelectionFormResult, MultiValueSelectionForm
from locommon import get_earliest_book, name_to_field, field_to_name, get_last_book


class PathMaker(object):
    """A class to create directory and file paths from the passed book.
    
    Some of the functions are based on functions in wadegiles's guided rename script. (c) wadegiles. Most have been heavily modified.
    """

    template_to_field = {"series" : "ShadowSeries", "number" : "ShadowNumber", "count" : "ShadowCount", "Day" : "Day", "ReleasedDate": "ReleasedTime",
                         "AddedDate" : "AddedTime", "EndYear" : "EndYear", "EndMonth" : "EndMonth", "EndMonth#" : "EndMonth",
                         "month" : "Month", "month#" : "Month", "year" : "ShadowYear", "imprint" : "Imprint", "publisher" : "Publisher",
                         "altSeries" : "AlternateSeries", "altNumber" : "AlternateNumber", "altCount" : "AlternateCount",
                         "volume" : "ShadowVolume", "title" : "ShadowTitle", "ageRating" : "AgeRating", "language" : "LanguageAsText",
                         "format" : "ShadowFormat", "startyear" : "StartYear", "writer" : "Writer", "tags" : "Tags", "genre" : "Genre",
                         "characters" : "Characters", "altSeries" : "AlternateSeries", "teams" : "Teams", "scaninfo" : "ScanInformation",
                         "manga" : "Manga", "seriesComplete" : "SeriesComplete", "first" : "FirstLetter", "read" : "ReadPercentage",
                         "counter" : "Counter", "startmonth" : "StartMonth", "startmonth#" : "StartMonth", "colorist" : "Colorist", "coverartist" : "CoverArtist",
                         "editor" : "Editor", "inker" : "Inker", "letterer" : "Letterer", "locations" : "Locations", "penciller" : "Penciller", "storyarc" : "StoryArc",
                         "seriesgroup" : "SeriesGroup", "maincharacter" : "MainCharacterOrTeam", "firstissuenumber" : "FirstIssueNumber", "lastissuenumber" : "LastIssueNumber", "Rating" : "Rating", 'CommunityRating': 'CommunityRating', "Custom" : 'Custom'}

    template_regex = re.compile("{(?P<prefix>[^{}<]*)<(?P<name>[^\d\s(>]*)(?P<args>\d*|(?:\([^)]*\))*)>(?P<postfix>[^{}]*)}")

    yes_no_fields = ["Manga", "SeriesComplete"]


    def __init__(self, parentform, profile):

        self._counter = None
        self.profile = profile
        self.failed_fields = []
        self._failed = False

        #Need to store the parent form so it can use the muilt select form
        self.form = parentform


    def make_path(self, book, folder_template, file_template):
        """Creates a path from a book with the given folder and file templates.
        Returns folder path, file name, and a bool if any values were empty.
        """
        self.book = book

        self._failed = False
        self.failed_fields = []


        #if self.profile.FailEmptyValues:
        #    for field in self.profile.FailedFields:
        #        if getattr(self.book, field) in ("", -1, MangaYesNo.Unknown, YesNo.Unknown):
        #            self.failed_fields.append(field)
        #            self._failed = True
            #
            #if self._failed and not self.profile.MoveFailed:
            #    return book.FileDirectory, book.FileNameWithExtension, True

        #Do filename first so that if MoveFailed is true the base folder is used correctly.
        file_path = self.book.FileNameWithExtension
        if self.profile.UseFileName:
            file_path = self.make_file_name(file_template)

        folder_path = book.FileDirectory
        if self.profile.UseFolder:

            folder_path = self.make_folder_path(folder_template)

        if self._failed and not self.profile.MoveFailed:
            return book.FileDirectory, book.FileNameWithExtension, True
            
        return folder_path, file_path, self._failed
        
 
    def make_folder_path(self, template):
        
        folder_path = ""

        template = template.strip()
        template = template.strip("\\")

        if template:
 
            rough_path = self.insert_fields_into_template(template)
            
            #Split into seperate directories for fixing empty paths and other problems.
            lines = rough_path.split("\\")
            
            for line in lines:
                if not line.strip():
                    line = self.profile.EmptyFolder
                line = self.replace_illegal_characters(line)
                #Fix for illegal periods at the end of folder names
                line = line.strip(".")
                folder_path = Path.Combine(folder_path, line.strip())
        
        if self._failed and self.profile.MoveFailed:
            folder_path = Path.Combine(self.profile.FailedFolder, folder_path)
        else:
            folder_path = Path.Combine(self.profile.BaseFolder, folder_path)


        if self.profile.ReplaceMultipleSpaces:
            folder_path = re.sub("\s\s+", " ", folder_path)

        return folder_path

    
    def make_file_name(self, template):
        """Creates file name with the template.

        template->The template to use.
        Returns->The created file name with extension and a bool if any values were empty.
        """
     
        file_name = self.insert_fields_into_template(template)
        file_name = file_name.strip()
        file_name = self.replace_illegal_characters(file_name)

        if not file_name:
            return ""


        extension = self.profile.FilelessFormat

        if self.book.FilePath:
            extension = FileInfo(self.book.FilePath).Extension

        #replace occurences of multiple spaces with a single space.
        if self.profile.ReplaceMultipleSpaces:
            file_name = re.sub("\s\s+", " ", file_name)


        return file_name + extension

    
    def insert_fields_into_template(self, template):
        """Replaces fields in the template with the correct field text."""
        self.invalid = 0
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
        result = ""
        conditional = False
        inversion = False
        inversion_args = ""

        name = match_groups["name"]
        args = match_groups["args"]

        #Inversions
        if name.startswith("!"):
            inversion = True
            name = name.lstrip("!")
            #Inversions can optionally have args.
            if args:
                #Get the last arg and removing from the other args
                r = re.search("(\([^(]*\))$", args)
                if r is not None:
                    args = args[:-len(r.group(0))]
                    inversion_args = r.group(0)[1:-1]

        #Conditionals
        if name.startswith("?"):
            conditional = True
            name = name.lstrip("?")
            if args:
                #Get the last arg and removing from the other args
                r = re.search("(\([^(]*\))$", args)
                if r is None:
                    self.invalid += 1
                    return match.group(0)
                args = args[:-len(r.group(0))]
                conditional_args = r.group(0)[1:-1]
            else:
                self.invalid += 1
                return match.group(0)

        #Checking template names
        if name in self.template_to_field:
            field = self.template_to_field[name]
        else:
            self.invalid += 1
            return match.group(0)



        #Get the fields
        result = self.get_field_text(field, name, args)
        
        #Invalid field result (possibly wrong number of arguments)
        if result is None:
            self.invalid += 1
            return match.group(0)

        #Conditionals
        if conditional:
            #Regex conditional
            if conditional_args.startswith("!"):
                if conditional_args[1:] == "":
                    return ""
                #Insert prefix and suffix if there is a match.
                if re.match(conditional_args[1:], result) is not None:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""
            else:
                #Text argument. Insert if matching the result
                if result == conditional_args:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""

        #Inversions
        if inversion:
            if not inversion_args:
                #No args so only insert the prefix and suffix if the result is empty
                if not result:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""

            elif inversion_args.startswith("!"):
                #Regex so only insert the prefix and suffix if there is no matches
                if re.match(inversion_args[1:], result) is None:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""
            else:
                #Text to match to. Only insert if the result doesn't match the arg
                if result != inversion_args:
                    return match_groups["prefix"] + match_groups["postfix"]
                else:
                    return ""



        #Empty results
        if not result:
            if self.profile.FailEmptyValues and field in self.profile.FailedFields:
                if field not in self.failed_fields:
                    self.failed_fields.append(field)
                self._failed = True

            if field in self.profile.EmptyData and self.profile.EmptyData[field]:
                return self.profile.EmptyData[field]
            else:
                return ""


        return match_groups["prefix"] + result + match_groups["postfix"]
    

    def get_field_text(self, field, template_name, args_match):
        """Gets the text of a field or operation.
        
        field->The name of the field to use.
        template_name->The name of the field as entered in the template. This is for keeping month and month# seperate.
        args_match->The args matched from the regex.

        Returns a string or None if an error occurs or the field isn't valid.
        """
        args = re.findall("\(([^)]*)\)", args_match)

        try:
            if field in ("StartYear", "StartMonth", "EndYear", "EndMonth"):
                return self.insert_start_value(field, args_match, template_name)

            elif field == "ReadPercentage" and len(args) == 3:
                return self.insert_read_percentage(args)

            elif field == "FirstLetter":
                if len(args) == 0:
                    return None
                elif args[0] in name_to_field or args[0] in field_to_name:
                    return self.insert_first_letter(args[0])

            elif field == "Counter" and len(args) == 3:
                return self.insert_counter(args)

            elif field == "FirstIssueNumber":
                return self.insert_first_issue_number(args_match)

            elif field == "LastIssueNumber":
                return self.insert_last_issue_number(args_match)

            elif field == "Custom":
                return self.insert_custom_value(args[0])

            elif type(getattr(self.book, field)) is System.DateTime:
                if args:
                    return self.insert_formated_datetime(args[0], field)
                else:
                    return self.insert_formated_datetime("", field)

            #Yes/no fields can have 1 or 2 args
            elif field in self.yes_no_fields and 0 < len(args) < 3:
                return self.insert_yes_no_field(field, args)

            elif len(args) == 2:
                return self.insert_multi_value_field(field, args)

            elif not args_match:
                return self.insert_text_field(field, template_name)

            elif args_match.isdigit():
                return self.insert_number_field(field, int(args_match))

            else:
                return None
        except (AttributeError, ValueError), ex:
            print ex
            return None


    def insert_text_field(self, field, template_name):
        """Get the string of any field and stips it of any illegal characters.

        field->The name of the field to get.
        template_name->The name of the field as entered in the template. This is for keeping month and month# seperate.
        Returns: The unicode string of the field."""

        if field == "Month" and not template_name.endswith("#"):
            return self.insert_month_as_name()

        text = getattr(self.book, field)

        if not text or text == -1:
            return ""

        return self.replace_illegal_characters(unicode(text))

    
    def insert_number_field(self, field, padding):
        """Get the padded value of a number field. Replaces illegal character in the number field.

        field->The name of the field to use.
        padding->integer of the ammount of padding to use.
        Returns-> Unicode string with the padded number value or empty string if the field is empty.
        """
        number = getattr(self.book, field)

        if number == -1:
            return ""

        if number == "":
            return ""

        if padding == 0:
            value = getattr(get_last_book(self.book), field)
            padding = len(str(value))

        return self.replace_illegal_characters(self.pad(number, padding))


    def insert_yes_no_field(self, field, args):
        """Gets a value using a yes/no field.

        field->The name of the field to check.
        args->A list with one or two items. 
              First item should be a string of what text to insert when the value is Yes.
              Second item (optional) should be a ! or null. ! means the text is inserted when the value is No.
        Returns the user text or an empty string."""
              

        text = args[0]

        no = False

        result = ""

        if len(args) == 2 and args[1] == "!":
            no = True

        field_value = getattr(self.book, field)

        if not no and field_value in (MangaYesNo.Yes, MangaYesNo.YesAndRightToLeft, YesNo.Yes):
            result = text

        elif no and field_value in (MangaYesNo.No, YesNo.No):
            result = text

        return self.replace_illegal_characters(result)


    def insert_read_percentage(self, args):
        """Get a value from the book's readpercentage.

        args should be a list with 3 items:
            1. Should be a string of the text to insert if the readpercentage matches the caclulations.
            2. Should be an operator: < > =
            3. Should be the percentage to match.
        Returns the user text or and empty string.
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
        """Gets the first letter of a field not counting articles.

        field->The name of the field to find the first letter of.
        Returns a single character string.
        """
        if field in name_to_field:
            field = name_to_field[field]

        field_text = unicode(getattr(self.book, field))

        result = ""

        match_result = re.match(r"(?:(?:the|a|an|de|het|een|die|der|das|des|dem|der|ein|eines|einer|einen|la|le|l'|les|un|une|el|las|los|las|un|una|unos|unas|o|os|um|uma|uns|umas|en|et|il|lo|uno|gli)\s+)?(?P<letter>.).+", field_text, re.I)

        if match_result:
            result = match_result.group("letter").capitalize()

        return self.replace_illegal_characters(result)

    def insert_counter(self, args):
        """Get the (padded) next number in a counter given a start and increment.

        args should be a list with 3 items:
            1. An integer with the start number.
            2. An integer of the increment.
            3. The amount of padding to use.
        Returns a string of the number.
        """
        start = int(args[0])
        increment = int(args[1])
        pad = args[2]

        result = ""

        if pad == "":
            pad = 0

        else:
            pad = int(pad)

        if self._counter is None:
            self._counter = start
            result = self.pad(self._counter, pad)

        else:
            self._counter = self._counter + increment
            result = self.pad(self._counter, pad)

        return result


    def insert_month_as_name(self):
        """Get the month name from a month number."""
        month_number = self.book.Month

        if month_number in self.profile.Months:
            return self.replace_illegal_characters(self.profile.Months[month_number])

        return ""


    def insert_start_value(self, field, args, template_name):
        """Get the value for StartYear or StartMonth.
        
        field->The name of the field to get.
        args->args is padding for the startmonth as number. Can be null.
        template_name->The name of the field as entered in the template. This is for keeping startmonth and startmonth# seperate.

        returns the string of the field or an empty string.
        """

        if field in ("StartYear", "EndYear"):
            return self.insert_start_year(field == "EndYear")

        elif field in ("StartMonth", "EndMonth"):
            return self.insert_start_month(args, template_name, field == "EndMonth")


    def insert_start_year(self, end=False):
        """Gets the start year of the earliest book in the series of the current issue."""
        if end:
            year = get_last_book(self.book).ShadowYear
        else:
            year = get_earliest_book(self.book).ShadowYear

        if year == -1:
            return ""

        return self.replace_illegal_characters(unicode(year))


    def insert_start_month(self, args, template_name, end=False):
        """Gets the start month of from the earliest issues in the series.
        Depending on the template_name the month as name or month as a number is inserted.
        args->can be either null or a number.
        template_name->The name of the field as entered in the template. This is for keeping month and month# seperate.
        Returns a string."""
        if end:
            month = get_last_book(self.book).Month
        else:
            month = get_earliest_book(self.book).Month

        if month == -1:
            return ""
        if template_name.endswith('#'):
            if args and args.isdigit():
                month = self.pad(month, int(args))
            return self.replace_illegal_characters(unicode(month))

        else:
            if month in self.profile.Months:
                return self.replace_illegal_characters(self.profile.Months[month])
            return ""

    def insert_formated_datetime(self, time_format, field):

        date_time = getattr(self.book, field)
        if time_format:
            return date_time.ToString(time_format)
        return date_time.ToString()

    def insert_custom_value(self, key):
        r = self.book.GetCustomValue(key)
        if r is None:
            return ""
        return r

    def insert_first_issue_number(self, padding):
        """
        padding is the padding used, can be none.
        """
        number = get_earliest_book(self.book).ShadowNumber
        
        if padding is not None and padding.isdigit():
            return self.replace_illegal_characters(self.pad(number, int(padding)))
        else:
            return self.replace_illegal_characters(number)


    def insert_last_issue_number(self, padding):
        """
        padding is the padding used, can be none.
        """
        number = get_last_book(self.book).ShadowNumber
        
        if padding is not None and padding.isdigit():
            return self.replace_illegal_characters(self.pad(number, int(padding)))
        else:
            return self.replace_illegal_characters(number)


    def insert_multi_value_field(self, field, args):
        """Gets the value from a multiple value field.

        args should be a list of two items:
            1. Should be the seperator between the items.
            2. A string of either issues or series.
        Returns a string containing the values.
        """
        seperator = args[0]
        mode = args[1]

        if mode == "series":
            return self.insert_multi_value_series(field, seperator)

        elif mode == "issue":
            return self.insert_multi_value_issue(field, seperator)

        else:
            return None


    def insert_multi_value_issue(self, field, seperator):
        """
        Finds which values to use in a multiple value field per issue. Asks the user which values to use and offers
        options to save that selection for following issues. Checks the stored selections and only asks the user if required.

        field->The string name of the field to use.
        seperator->The string to seperate every value with.
        Returns a string of the values or an empty string if no values are found or selected.
        """

        #field_dict stores the alwaysused collections and also the selected values for each issue used.
        try:
            field_dict = getattr(self, field)
        except AttributeError:
            field_dict = {}
            setattr(self, field, field_dict)


        index = self.book.Publisher + self.book.ShadowSeries + str(self.book.ShadowVolume) + self.book.ShadowNumber
        booktext = self.book.ShadowSeries + " vol. " + str(self.book.ShadowVolume) + " #" + self.book.ShadowNumber


        if index in field_dict:
            #This particular issue has already been done.
            result = field_dict[index]
            return self.make_multi_value_issue_string(result.Selection, seperator, result.Folder)


        if not getattr(self.book, field).strip():
            field_dict[index] = MultiValueSelectionFormResult([])
            return ""


        try:
            always_used_values = getattr(self, field + "AlwaysUse")
        except AttributeError:
            always_used_values = []
            setattr(self, field + "AlwaysUse", always_used_values)


        values = [item.strip() for item in getattr(self.book, field).split(",")]


        if len(values) == 1 and self.profile.DontAskWhenMultiOne:
            return self.replace_illegal_characters(values[0])


        selected_values = []

        #Find which items are set to always use
        for list_of_always_used_values in sorted(always_used_values, key=len, reverse=True):
            count = 0
            for value in list_of_always_used_values:
                if value in values:
                    count +=1

            if count == len(list_of_always_used_values):

                if list_of_always_used_values.do_not_ask:
                    return self.make_multi_value_issue_string(list_of_always_used_values, seperator, list_of_always_used_values.use_folder_seperator)

                selected_values = list_of_always_used_values[:]
                break

        if self.form.InvokeRequired:
            result = self.form.Invoke(Func[MultiValueSelectionFormArgs, MultiValueSelectionFormResult](self.show_multi_value_selection_form), System.Array[object]([MultiValueSelectionFormArgs(values, selected_values, field, booktext, False)]))

        else:
            result = self.show_multi_value_selection_form(MultiValueSelectionFormArgs(values, selected_values, field, booktext, False))

        field_dict[index] = result

        if len(result.Selection) == 0:
            return ""

        if result.AlwaysUse:
            always_used_values.append(MultiValueAlwaysUsedValues(result.AlwaysUseDontAsk, result.Folder, result.Selection))
        
        return self.make_multi_value_issue_string(result.Selection, seperator, result.Folder)


    def make_multi_value_issue_string(self, values, seperator, usefolder):
        #When folders are being used we need to make sure there are no "\" characters in any of the values or it will mess up the folders. 
        if usefolder:
            seperator = self.replace_illegal_characters(seperator)
            seperator += "\\"
            values = [self.replace_illegal_characters(value) for value in values]

        
            return seperator.join(values)

        return self.replace_illegal_characters(seperator.join(values))


    def insert_multi_value_series(self, field, seperator):
        """
        Finding a multiple value field via a series operation will find all the possible multiple values of the field
        from every issue of the series in the library. The user can then pick which values to use and those chosen values will
        be used for every issue encountered.

        The user can choose to use every chossen value for the every issue in the series, even if the issue doesn't have that value.
        Or the user can choose to only use the value if it is in the issue.
        
        field->The string name of the field to use.
        seperator->The string to seperate every value with
        """

        #Chosen values for series are stored in the field_dict.
        try:
            field_dict = getattr(self, field)
        except AttributeError:
            field_dict = {}
            setattr(self, field, field_dict)


        index = self.book.Publisher + self.book.ShadowSeries + str(self.book.ShadowVolume)
        booktext = self.book.ShadowSeries + " vol. " + str(self.book.ShadowVolume) 


        #See if this series has been done before:
        if index in field_dict:
            return self.make_multi_value_series_string(field_dict[index], seperator, field)



        values = self.get_all_multi_values_from_series(field)


        if not values:
            field_dict[index] = MultiValueSelectionFormResult([])
            return ""


        #Since this can be shown from the configform...
        if self.form.InvokeRequired:
            result = self.form.Invoke(Func[MultiValueSelectionFormArgs, MultiValueSelectionFormResult](self.show_multi_value_selection_form), System.Array[object]([MultiValueSelectionFormArgs(values, [], field, booktext, True)]))

        else:
            result = self.show_multi_value_selection_form(MultiValueSelectionFormArgs(values, [], field, booktext, True))

        field_dict[index] = result

        return self.make_multi_value_series_string(result, seperator, field)


    def make_multi_value_series_string(self, selection_result, seperator, field):
        """
        Makes the correctly formated inserted string based on the input values:

        selection_result is a SelectionFormResult object
        seperator is the string seperator
        field is the correct name of the field to find from the book
        """

        #If using folder sperator we need to make sure that there are no "\" characters in any of the fields.
        if selection_result.Folder:
            seperator = self.replace_illegal_characters(seperator)
            seperator += "\\"
            values = [self.replace_illegal_characters(value.strip()) for value in getattr(self.book, field).split(",")]
            selection = [self.replace_illegal_characters(value) for value in selection_result.Selection]
        else:
            values = [value.strip() for value in getattr(self.book, field).split(",")]
            selection = selection_result.Selection

        #Using every issue
        if selection_result.EveryIssue:
            result = seperator.join(selection)

        #Not using every issue so just use the ones the particular issue has.
        else:
            items_to_use = [item for item in selection if item in values]
            result = seperator.join(items_to_use)
        
        if selection_result.Folder:
            return result

        return self.replace_illegal_characters(result)


    def replace_illegal_characters(self, text):
        """Replaces illegal path characters in a string and retures the cleaned string."""
        for illegal_character in sorted(self.profile.IllegalCharacters.keys(), key=len, reverse=True):
            text = text.replace(illegal_character, self.profile.IllegalCharacters[illegal_character])

        return text

    
    def pad(self, value, padding):
        #value as string, padding as int
        remainder = ""

        try:
            numberValue = float(value)
        except ValueError:
            return value

        if type(value) in (int, System.Single):
            value = str(value)

        if numberValue >= 0:
            #To make sure that the item is padded correctly when a decimal such as 7.1
            if value.Contains("."):
                value, remainder = value.split(".")
                remainder = "." + remainder
            return value.PadLeft(padding, '0') + remainder
        else:
            value = value[1:]
            #To make sure that the item is padded correctly when a decimal such as 7.1
            if value.Contains("."):
                value, remainder = value.split(".")
                remainder = "." + remainder
            value = value.PadLeft(padding, '0')
            return '-' + value + remainder

    
    def show_multi_value_selection_form(self, args):
        form = MultiValueSelectionForm(args)
        form.ShowDialog()
        result = form.GetResults()
        form.Dispose()
        return result


    def get_all_multi_values_from_series(self, field):
        """
        Helper to get all the unique values from a multivale field in an entire series.

        field is the name of the field to find values from.
        Returns a list of the values.
        """
        allbooks = ComicRack.App.GetLibraryBooks()
        #using a set to avoid duplicate entries
        results = set()
        for book in allbooks:
            if book.ShadowSeries == self.book.ShadowSeries and book.ShadowVolume == self.book.ShadowVolume and book.Publisher == self.book.Publisher:
                results.update([value.strip() for value in getattr(book, field).split(",") if value.strip()])

        return list(results)


class MultiValueAlwaysUsedValues(list):

    def __init__(self, do_not_ask=False, folder=False, values=None):
        self.do_not_ask = do_not_ask
        self.use_folder_seperator = folder
        if values is None:
            values = []
        list.__init__(self, values)
