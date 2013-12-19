import localizer

class FieldType(object):
    """ An enum to compare field types """
    String = 1
    Number = 2
    MultipleValue = 3
    YesNo = 4
    ManagaYesNo = 5
    DateTime = 6

#I tried several ways to pull all this information in automatically but nothing worked quite right. This is simple enough to keep up though.
class TemplateItem(object):

    def __init__(self, backup_name, field, type, template=None, 
                 exclude=True, conditional=True):
        if template is None:
            template = field
        self.name = localizer.get_field_name(field, backup_name)
        self.field = field
        self.template = template
        self.type = type
        self.exclude = exclude
        self.conditional = conditional

class TemplateItemCollection(list):

    def __init__(self, *args, **kwargs):
        self._exclude_rule_fields = None
        return super(TemplateItemCollection, self).__init__(*args, **kwargs)
        
    def get_name_from_template(self, template):
        for item in self.__iter__():
            if item.template == template:
                return item.name
        raise KeyError

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

    


#Build all the fields
FIELDS = TemplateItemCollection()

FIELDS.append(TemplateItem("File Format", "FileFormat", "String", ""))
FIELDS.append(TemplateItem("File Name", "FileName", "String", ""))
FIELDS.append(TemplateItem("File Path", "FilePath", "String", ""))

FIELDS.append(TemplateItem("Added/Purchased", "AddedTime", FieldType.DateTime, "Added"))
FIELDS.append(TemplateItem("Age Rating", "AgeRating", "String"))
FIELDS.append(TemplateItem("Alternate Count", "AlternateCount", "Number"))
FIELDS.append(TemplateItem("Alternate Number", "AlternateNumber", "Number"))
FIELDS.append(TemplateItem("Alternate Series", "AlternateSeries", "MultipleValue"))
FIELDS.append(TemplateItem("Black And White", "BlackAndWhite", "YesNo"))
FIELDS.append(TemplateItem("Book Age", "BookAge", "String", "Age"))
FIELDS.append(TemplateItem("Book Collection Status", "BookCollectionStatus", "String", "CollectionStatus"))
FIELDS.append(TemplateItem("Book Condition", "BookCondition", "String", "Condition"))
FIELDS.append(TemplateItem("Book Location", "BookLocation", "String"))
FIELDS.append(TemplateItem("Book Notes", "BookNotes", "String", ""))
FIELDS.append(TemplateItem("Book Owner", "BookOwner", "String", "Owner"))
FIELDS.append(TemplateItem("Book Price", "BookPrice", "String", "Price"))
FIELDS.append(TemplateItem("Book Store", "BookStore", "String", "Store"))
FIELDS.append(TemplateItem("Characters", "Characters", "MultipleValue"))
FIELDS.append(TemplateItem("Checked", "Checked", "Bool"))
FIELDS.append(TemplateItem("Colorist", "Colorist", "MultipleValue"))
FIELDS.append(TemplateItem("Community Rating", "CommunityRating", "Number"))
FIELDS.append(TemplateItem("Conditional", "Conditional", "Conditional", exclude=False, conditional=False))
FIELDS.append(TemplateItem("Count", "ShadowCount", "Number", "Count"))
FIELDS.append(TemplateItem("Counter", "Counter", "Counter", exclude=False))
FIELDS.append(TemplateItem("Custom Value", "CustomValue", "CustomValue"))
FIELDS.append(TemplateItem("Cover Artist", "CoverArtist", "MultipleValue"))
FIELDS.append(TemplateItem("Day", "Day", "Number"))
FIELDS.append(TemplateItem("Editor", "Editor", "MultipleValue"))
FIELDS.append(TemplateItem("First Letter", "FirstLetter", "FirstLetter", exclude=False))
FIELDS.append(TemplateItem("Format", "ShadowFormat", "String", "Format"))
FIELDS.append(TemplateItem("Genre", "Genre", "MultipleValue"))
FIELDS.append(TemplateItem("Has Custom Values", "HasCustomValues", "Bool", ""))
FIELDS.append(TemplateItem("Imprint", "Imprint", "String"))
FIELDS.append(TemplateItem("Inker", "Inker", "MultipleValue"))
FIELDS.append(TemplateItem("ISBN", "ISBN", "String"))
FIELDS.append(TemplateItem("Language", "LanguageAsText", "String", "Language"))
FIELDS.append(TemplateItem("Letterer", "Letterer", "MultipleValue"))
FIELDS.append(TemplateItem("Locations", "Locations", "MultipleValue"))
FIELDS.append(TemplateItem("Main Character Or Team", "MainCharacterOrTeam", "String"))
FIELDS.append(TemplateItem("Manga", "Manga", "MangaYesNo"))
FIELDS.append(TemplateItem("Month", "Month", "Month"))
FIELDS.append(TemplateItem("Notes", "Notes", "String", ""))
FIELDS.append(TemplateItem("Number", "ShadowNumber", "Number", "Number"))
FIELDS.append(TemplateItem("Penciller", "Penciller", "MultipleValue"))
FIELDS.append(TemplateItem("Published", "Published", FieldType.DateTime, ""))
FIELDS.append(TemplateItem("Publisher", "Publisher", "String"))
FIELDS.append(TemplateItem("Rating", "Rating", "Number"))
FIELDS.append(TemplateItem("Read Percentage", "ReadPercentage", "ReadPercentage"))
FIELDS.append(TemplateItem("Released", "ReleasedTime", FieldType.DateTime, "Released"))
FIELDS.append(TemplateItem("Review", "Review", "String", ""))
FIELDS.append(TemplateItem("Scan Information", "ScanInformation", "MultipleValue"))
FIELDS.append(TemplateItem("Series", "ShadowSeries", "String", "Series"))
FIELDS.append(TemplateItem("Series Complete", "SeriesComplete", "YesNo"))
FIELDS.append(TemplateItem("Series Group", "SeriesGroup", "String"))
FIELDS.append(TemplateItem("Series: First Number", "FirstNumber", "Number"))
FIELDS.append(TemplateItem("Series: First Month", "FirstMonth", "Month"))
FIELDS.append(TemplateItem("Series: First Year", "FirstYear", "Year"))
FIELDS.append(TemplateItem("Series: Last Month", "LastMonth", "Month"))
FIELDS.append(TemplateItem("Series: Last Number", "LastNumber", "Number"))
FIELDS.append(TemplateItem("Series: Last Year", "LastYear", "Year"))
FIELDS.append(TemplateItem("Series: Percentage Read", "SeriesReadPercentage", "ReadPercentage"))
FIELDS.append(TemplateItem("Series: Running Time Years", "SeriesRunningTimeYears", "Number", ""))
FIELDS.append(TemplateItem("Story Arc", "StoryArc", "String"))
FIELDS.append(TemplateItem("Summary", "Summary", "String", ""))
FIELDS.append(TemplateItem("Tags", "Tags", "MultipleValue"))
FIELDS.append(TemplateItem("Teams", "Teams", "MultipleValue"))
FIELDS.append(TemplateItem("Text", "Text", "Text", "", exclude=False, conditional=True))
FIELDS.append(TemplateItem("Title", "ShadowTitle", "String", "Title"))
FIELDS.append(TemplateItem("Volume", "ShadowVolume", "Number", "Volume"))
FIELDS.append(TemplateItem("Web", "Web", "String", "Web"))
FIELDS.append(TemplateItem("Week", "Week", "Number", "Week"))
FIELDS.append(TemplateItem("Writer", "Writer", "MultipleValue"))
FIELDS.append(TemplateItem("Year", "ShadowYear", "Year", "Year"))

FIELDS.sort(key=lambda x: x.name)

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



