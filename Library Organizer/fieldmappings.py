import localizer

class FieldType(object):
    """ An enum for field types """
    String = 'String'
    Number = 'Number'
    MultipleValue = 'MultipleValue'
    YesNo = 'YesNo'
    MangaYesNo = 'MangaYesNo'
    DateTime = 'DateTime'
    Boolean = 'Boolean'
    Month = 'Month'
    Year = 'Year'
    CustomValue = 'CustomValue'
    Conditional = 'Conditional'
    ReadPercentage = 'ReadPercentage'
    Text = 'Text'
    FirstLetter = 'FirstLetter'
    Counter = 'Counter'

#I tried several ways to pull all this information in automatically but nothing worked quite right. This is simple enough to keep up though.
class Field(object):

    def __init__(self, backup_name, field, type, template=None, 
                 exclude=True, conditional=True):
        """Creates a new field.

        Args:
            backup_name: The name used in the translation if no
                translation is available.
            field: The string representation of the property on the
                ComicBook.
            type: The FieldType of this field.
            template: The string to use for the template builder. Pass
                False if this field should not be used in the template
                builder. Pass None if the template name is the same as
                the field string.
            exclude: A boolean if this field should be usable in
                exclude rules.
            conditional: A boolean if this field should be usable in a
                conditional field.
        """
        if template is None:
            template = field
        self.name = backup_name
        self.backup_name = backup_name
        self.field = field
        self.template = template
        self.type = type
        self.exclude = exclude
        self.conditional = conditional

class FieldList(list):

    def __init__(self, *args, **kwargs):
        self._exclude_rule_fields = None
        self._template_fields = None
        return super(FieldList, self).__init__(*args, **kwargs)

    def get_by_field(self, field):
        for item in self.__iter__():
            if item.field == field:
                return item
        raise KeyError

    @property
    def exclude_rule_fields(self):
        if self._exclude_rule_fields is None:
            l = [f for f in self.__iter__() if f.exclude]
            self._exclude_rule_fields = l
        return self._exclude_rule_fields

    def add(self, backup_name, field, type, template=None, exclude=True,
            conditional=True):
        """Adds a field to this field list.
        Args:
            backup_name: The name used in the translation if no
                translation is available.
            field: The string representation of the property on the
                ComicBook.
            type: The FieldType of this field.
            template: The string to use for the template builder. Pass
                False if this field should not be used in the template
                builder. Pass None if the template name is the same as
                the field string. Defaults to None.
            exclude: A boolean if this field should be usable in
                exclude rules. Defaults to True
            conditional: A boolean if this field should be usable in a
                conditional field. Defaults to True.
        """
        self.append(Field(backup_name, field, type, template, exclude,
                          conditional))

    def get_template_fields(self):
        return [i for i in self if i.template != False]

    def get_conditional_fields(self):
        return [i for i in self if i.conditional != False and i.field != "Text"]

    def get_conditional_then_else_fields(self):
        return [i for i in self if i.conditional != False]

    def get_exclude_fields(self):
        return [i for i in self if i.exclude == True]


#Build all the fields
def create_fields():
    fields = FieldList()
    field_list = []
    with open('fields.csv', 'r') as f:
        field_list = f.readlines()
    for line in field_list[1:]:
        values = line.strip().split(',')
        fields.append(Field(values[0], values[1], values[2], values[3], 
                            values[4] == 'True', values[5] == 'True'))
    return fields
FIELDS = create_fields()


##############################   Date fields   #################################
#FIELDS.add("Added", "AddedTime", FieldType.DateTime, "Added")
#FIELDS.add("Published", "Published", FieldType.DateTime)
#FIELDS.add("Released", "ReleasedTime", FieldType.DateTime, "Released")

############################    String fields ##################################
#FIELDS.add("Age Rating", "AgeRating", FieldType.String)
#FIELDS.add("Alternate Series", "AlternateSeries", FieldType.String)
#FIELDS.add("Book Age", "BookAge", FieldType.String, "Age")
#FIELDS.add("Book Collection Status", "BookCollectionStatus", FieldType.String, 
#           "CollectionStatus")
#FIELDS.add("Book Condition", "BookCondition", FieldType.String, "Condition")
#FIELDS.add("Book Location", "BookLocation", FieldType.String, "Location")
#FIELDS.add("Book Notes", "BookNotes", FieldType.String, False, conditional=False)
#FIELDS.add("Book Owner", "BookOwner", FieldType.String, "Owner")
#FIELDS.add("Book Store", "BookStore", FieldType.String, "Store")
#FIELDS.add("File Directory", "FileDirectory", FieldType.String, False, conditional=False)
#FIELDS.add("File Format", "FileFormat", FieldType.String, False, conditional=False)
#FIELDS.add("File Name", "FileName", FieldType.String, False, conditional=False)
#FIELDS.add("File Path", "FilePath", FieldType.String, False, conditional=False)
#FIELDS.add("Format", "ShadowFormat", FieldType.String, "Format")
#FIELDS.add("Imprint", "Imprint", FieldType.String)
#FIELDS.add("ISBN", "ISBN", FieldType.String)
#FIELDS.add("Language", "LanguageAsText", FieldType.String, "Language")
#FIELDS.add("Notes", "Notes", FieldType.String, False, conditional=False)
#FIELDS.add("Publisher", "Publisher", FieldType.String)
#FIELDS.add("Review", "Review", FieldType.String, False, conditional=False)
#FIELDS.add("Series", "ShadowSeries", FieldType.String, "Series")
#FIELDS.add("Series Group", "SeriesGroup", FieldType.String)
#FIELDS.add("Story Arc", "StoryArc", FieldType.String)
#FIELDS.add("Summary", "Summary", FieldType.String, False, conditional=False)
#FIELDS.add("Title", "ShadowTitle", FieldType.String, "Title")
#FIELDS.add("Web", "Web", FieldType.String, False, conditional=False)

########################  MultipleValue fields  ################################
#FIELDS.add("Characters", "Characters", FieldType.MultipleValue)
#FIELDS.add("Colorist", "Colorist", FieldType.MultipleValue)
#FIELDS.add("Cover Artist", "CoverArtist", FieldType.MultipleValue)
#FIELDS.add("Editor", "Editor", FieldType.MultipleValue)
#FIELDS.add("Genre", "Genre", FieldType.MultipleValue)
#FIELDS.add("Inker", "Inker", FieldType.MultipleValue)
#FIELDS.add("Letterer", "Letterer", FieldType.MultipleValue)
#FIELDS.add("Locations", "Locations", FieldType.MultipleValue)
#FIELDS.add("Main Character Or Team", "MainCharacterOrTeam", 
#           FieldType.MultipleValue)
#FIELDS.add("Penciller", "Penciller", FieldType.MultipleValue)
#FIELDS.add("Scan Information", "ScanInformation", FieldType.MultipleValue)
#FIELDS.add("Tags", "Tags", FieldType.MultipleValue)
#FIELDS.add("Teams", "Teams", FieldType.MultipleValue)
#FIELDS.add("Writer", "Writer", FieldType.MultipleValue)

##########################   Number fields   ###################################
#FIELDS.add("Alternate Count", "AlternateCount", FieldType.Number)
#FIELDS.add("Alternate Number", "AlternateNumber", FieldType.Number)
#FIELDS.add("Book Price", "BookPrice", FieldType.Number, "Price")
#FIELDS.add("Community Rating", "CommunityRating", FieldType.Number)
#FIELDS.add("Count", "ShadowCount", FieldType.Number, "Count")
#FIELDS.add("Day", "Day", FieldType.Number)
#FIELDS.add("Number", "ShadowNumber", FieldType.Number, "Number")
#FIELDS.add("Rating", "Rating", FieldType.Number)
#FIELDS.add("Volume", "ShadowVolume", FieldType.Number, "Volume")
#FIELDS.add("Week", "Week", FieldType.Number)
#FIELDS.add("Series: First Number", "FirstNumber", FieldType.Number)
#FIELDS.add("Series: Last Number", "LastNumber", FieldType.Number)
#FIELDS.add("Series: Running Time Years", "SeriesRunningTimeYears", 
#           FieldType.Number, False, conditional=False)

###########################    YesNo fields   ##################################
#FIELDS.add("Black And White", "BlackAndWhite", FieldType.YesNo)
#FIELDS.add("Series Complete", "SeriesComplete", FieldType.YesNo)

########################   MangaYesNo fields ###################################
#FIELDS.add("Manga", "Manga", FieldType.MangaYesNo)

###############################  Month fields  #################################
#FIELDS.add("Month", "Month", FieldType.Month)
#FIELDS.add("Series: First Month", "FirstMonth", FieldType.Month)
#FIELDS.add("Series: Last Month", "LastMonth", FieldType.Month)

###########################    Year fields #####################################
#FIELDS.add("Year", "ShadowYear", FieldType.Year, "Year")
#FIELDS.add("Series: First Year", "FirstYear", FieldType.Year)
#FIELDS.add("Series: Last Year", "LastYear", FieldType.Year)

############################  Boolean Fields  ##################################
#FIELDS.add("Checked", "Checked", FieldType.Boolean, False, conditional=False)
#FIELDS.add("File Is Missing", "FileIsMissing", FieldType.Boolean, False, conditional=False)
#FIELDS.add("Has Been Opened", "HasBeenOpened", FieldType.Boolean, False, conditional=False)
#FIELDS.add("Has Been Read", "HasBeenRead", FieldType.Boolean, False, conditional=False)
#FIELDS.add("Has Custom Values", "HasCustomValues", FieldType.Boolean, False, conditional=False)

############################  Special Fields  ##################################
#FIELDS.add("Conditional", "Conditional", FieldType.Conditional, exclude=False, 
#           conditional=False)
#FIELDS.add("Custom Value", "CustomValue", FieldType.CustomValue)
#FIELDS.add("Read Percentage", "ReadPercentage", FieldType.ReadPercentage)
#FIELDS.add("Series: Percentage Read", "SeriesReadPercentage", 
#           FieldType.ReadPercentage)
#FIELDS.add("Counter", "Counter", FieldType.Counter, exclude=False)
#FIELDS.add("First Letter", "FirstLetter", FieldType.FirstLetter, exclude=False)
#FIELDS.add("Text", "Text", FieldType.Text, False, exclude=False, conditional=True)






















#FIELDS.sort(key=lambda x: x.name)

# This contains the fields that are available to add into the template. 
# Used for building the field selector list
template_fields = ['AddedTime',
                   'AgeRating', 
                   'AlternateCount', 
                   'AlternateNumber', 
                   'AlternateSeries', 
                   'BlackAndWhite', 
                   'BookAge', 
                   'BookCollectionStatus', 
                   'BookCondition', 
                   'BookLocation', 
                   'BookOwner', 
                   'BookPrice', 
                   'BookStore', 
                   'Characters', 
                   'Colorist', 
                   'CommunityRating', 
                   'Conditional', 
                   'ShadowCount', 
                   'Counter', 
                   'CoverArtist',
                   'CustomValue',
                   'Day', 
                   'Editor', 
                   'EndMonth',
                   'EndYear',
                   'FirstIssueNumber',
                   'FirstLetter', 
                   'ShadowFormat', 
                   'Genre', 
                   'ISBN', 
                   'Imprint', 
                   'Inker', 
                   'LanguageAsText', 
                   'LastIssueNumber',
                   'Letterer', 
                   'Locations', 
                   'MainCharacterOrTeam', 
                   'Manga', 
                   'Month', 
                   'ShadowNumber', 
                   'Penciller', 
                   'Publisher', 
                   'Rating', 
                   'ReadPercentage', 
                   'ReleasedTime',
                   'ScanInformation', 
                   'ShadowSeries', 
                   'SeriesComplete', 
                   'SeriesGroup', 
                   'StartMonth', 
                   'StartYear', 
                   'StoryArc', 
                   'Tags', 
                   'Teams', 
                   'ShadowTitle', 
                   'ShadowVolume', 
                   'Writer', 
                   'ShadowYear']

# This defines fields that are usable in the first letter selector. 
first_letter_fields = ['AlternateSeries', 
                       'Imprint', 
                       'Publisher', 
                       'ShadowSeries']

# This contains the fields usable in the selectors for conditional. 
# Should basically contain everything but Conditional with the addition of Text
conditional_fields = ['AddedTime',
                   'AgeRating', 
                   'AlternateCount', 
                   'AlternateNumber', 
                   'AlternateSeries', 
                   'BlackAndWhite', 
                   'BookAge', 
                   'BookCollectionStatus', 
                   'BookCondition', 
                   'BookLocation', 
                   'BookNotes',
                   'BookOwner', 
                   'BookPrice', 
                   'BookStore', 
                   'Characters', 
                   'Colorist', 
                   'CommunityRating', 
                   'ShadowCount', 
                   'CoverArtist',
                   'Day', 
                   'Editor', 
                   'EndMonth',
                   'EndYear',
                   'FirstIssueNumber',
                   'ShadowFormat', 
                   'Genre', 
                   'ISBN', 
                   'Imprint', 
                   'Inker', 
                   'LanguageAsText', 
                   'LastIssueNumber',
                   'Letterer', 
                   'Locations', 
                   'MainCharacterOrTeam', 
                   'Manga', 
                   'Month', 
                   'ShadowNumber', 
                   'Penciller', 
                   'Publisher', 
                   'Rating', 
                   'ReadPercentage', 
                   'ReleasedTime',
                   'ScanInformation', 
                   'ShadowSeries', 
                   'SeriesComplete', 
                   'SeriesGroup', 
                   'StartMonth', 
                   'StartYear', 
                   'StoryArc', 
                   'Tags', 
                   'Teams', 
                   'ShadowTitle', 
                   'ShadowVolume', 
                   'Writer', 
                   'ShadowYear']

conditional_then_else_fields = ['AddedTime',
                   'AgeRating', 
                   'AlternateCount', 
                   'AlternateNumber', 
                   'AlternateSeries', 
                   'BlackAndWhite', 
                   'BookAge', 
                   'BookCollectionStatus', 
                   'BookCondition', 
                   'BookLocation', 
                   'BookOwner', 
                   'BookPrice', 
                   'BookStore', 
                   'Characters', 
                   'Colorist', 
                   'CommunityRating', 
                   'ShadowCount', 
                   'Counter', 
                   'CoverArtist',
                   'Day', 
                   'Editor', 
                   'EndMonth',
                   'EndYear',
                   'FirstIssueNumber',
                   'FirstLetter', 
                   'ShadowFormat', 
                   'Genre', 
                   'ISBN', 
                   'Imprint', 
                   'Inker', 
                   'LanguageAsText', 
                   'LastIssueNumber',
                   'Letterer', 
                   'Locations', 
                   'MainCharacterOrTeam', 
                   'Manga', 
                   'Month', 
                   'ShadowNumber', 
                   'Penciller', 
                   'Publisher', 
                   'Rating', 
                   'ReadPercentage', 
                   'ReleasedTime',
                   'ScanInformation', 
                   'ShadowSeries', 
                   'SeriesComplete', 
                   'SeriesGroup', 
                   'StartMonth', 
                   'StartYear', 
                   'StoryArc', 
                   'Tags', 
                   'Teams', 
                   'Text',
                   'ShadowTitle', 
                   'ShadowVolume', 
                   'Writer', 
                   'ShadowYear']

#These are special fields that are created in library organizer. Needs to be separate for ease of getting translations
library_organizer_fields = ["Counter", "FirstLetter", "Conditional", 
                            "StartMonth", "StartYear", "FirstIssueNumber"]

#These are the fields usable in the exclude rules.
exclude_rule_fields = ['AgeRating', 
                       'AlternateCount', 
                       'AlternateNumber', 
                       'AlternateSeries', 
                       'BlackAndWhite', 
                       'BookAge', 
                       'BookCollectionStatus', 
                       'BookCondition', 
                       'BookLocation', 
                       'BookNotes', 
                       'BookOwner', 
                       'BookPrice', 
                       'BookStore', 
                       'Characters', 
                       #'Checked', 
                       'Colorist', 
                       'CommunityRating', 
                       'ShadowCount', 
                       'CoverArtist', 
                       'Editor', 
                       #'FileDirectory', 
                       'FileFormat', 
                       'FileName', 
                       #'FileIsMissing', 
                       #'FileNameWithExtension', 
                       'FilePath', 
                       #'FileSize', 
                       'ShadowFormat', 
                       'Genre', 
                       #'HasBeenOpened', 
                       #'HasBeenRead', 
                       'ISBN', 
                       'Imprint', 
                       'Inker', 
                       'LanguageAsText', 
                       'Letterer', 
                       'Locations', 
                       'MainCharacterOrTeam', 
                       'Manga', 
                       'Month', 
                       'Notes', 
                       'ShadowNumber', 
                       'Penciller', 
                       'Publisher', 
                       'Rating', 
                       'ReadPercentage', 
                       'Review', 
                       'ScanInformation', 
                       'ShadowSeries', 
                       'SeriesComplete', 
                       'SeriesGroup', 
                       'StoryArc', 
                       'Summary', 
                       'Tags', 
                       'Teams', 
                       'ShadowTitle', 
                       'ShadowVolume', 
                       'Web', 
                       'Writer', 
                       'ShadowYear']

rules_fields = ['BookNotes']



