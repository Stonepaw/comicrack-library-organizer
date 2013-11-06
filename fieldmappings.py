import localizer


#I tried serveral ways to pull all this information in automatically but nothing worked quite right. This is simple enough to keep up though.
class TemplateItem(object):

    def __init__(self, name, field, template, type):
        self.name = name
        self.field = field
        self.template = template
        self.type = type

class TemplateItemCollection(list):

    def get_name_from_template(self, template):
        for item in self.__iter__():
            if item.template == template:
                return item.name
        raise KeyError

    def get_item_by_field(self, field):
        for item in self.__iter__():
            if item.field == field:
                return item
        raise KeyError


#translatedfields = localizer.get_comic_field_from_comicbook_dialogs();

FIELDS = TemplateItemCollection()

FIELDS.append(TemplateItem("File Format", "FileFormat", "", ""))
FIELDS.append(TemplateItem("File Name", "FileName", "", ""))
FIELDS.append(TemplateItem("File Path", "FilePath", "", ""))


FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("AddedTime", "Added/Released"), "AddedTime", "Added", "DateTime"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("AgeRating", "Age Rating"), "AgeRating", "AgeRating", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_columns("AlternateCount", "Alternate Count"), "AlternateCount", "AlternateCount", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("AlternateNumber", "Alternate Number"), "AlternateNumber", "AlternateNumber", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("AlternateSeries", "Alternate Series"), "AlternateSeries", "AlternateSeries", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BlackAndWhite", "Black And White"), "BlackAndWhite", "BlackAndWhite", "YesNo"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookAge", "Age"), "BookAge", "Age", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookCollectionStatus", "Collection Status"), "BookCollectionStatus", "CollectionStatus", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookCondition", "Book Condition"), "BookCondition", "Condition", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookLocation", "Book Location"), "BookLocation", "BookLocation", "String"))
FIELDS.append(TemplateItem("Book Notes", "BookNotes", "", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookOwner", "Owner"), "BookOwner", "Owner", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookPrice", "Book Price"), "BookPrice", "Price", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("BookStore", "Store"), "BookStore", "Store", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Characters", "Characters"), "Characters", "Characters", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Colorist", "Colorist"), "Colorist", "Colorist", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("CommunityRating", "Community Rating"), "CommunityRating", "CommunityRating", "Number"))
FIELDS.append(TemplateItem("Conditional", "Conditional", "", "Conditional"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_columns("Count", "Count"), "ShadowCount", "Count", "Number"))
FIELDS.append(TemplateItem("Counter", "Counter", "Counter", "Counter"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("CoverArtist", "Cover Artist"), "CoverArtist", "CoverArtist", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Day", "Day"), "Day", "Day", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Editor", "Editor"), "Editor", "Editor", "MultipleValue"))
FIELDS.append(TemplateItem("End Month", "EndMonth", "EndMonth", "Month"))
FIELDS.append(TemplateItem("End Year", "EndYear", "EndYear", "Year"))
FIELDS.append(TemplateItem("First Issue Number", "FirstIssueNumber", "FirstIssueNumber", "Number"))
FIELDS.append(TemplateItem("First Letter", "FirstLetter", "FirstLetter", "FirstLetter"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Format", "Format"), "ShadowFormat", "Format", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Genre", "Genre"), "Genre", "Genre", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Imprint", "Imprint"), "Imprint", "Imprint", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Inker", "Inker"), "Inker", "Inker", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("ISBN", "ISBN"), "ISBN", "ISBN", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Language", "Language"), "LanguageAsText", "Language", "String"))
FIELDS.append(TemplateItem("Last Issue Number", "LastIssueNumber", "LastIssueNumber", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Letterer", "Letterer"), "Letterer", "Letterer", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Locations", "Locations"), "Locations", "Locations", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("MainCharacterOrTeam", "Main Character Or Team"), "MainCharacterOrTeam", "MainCharacterOrTeam", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Manga", "Manga"), "Manga", "Manga", "MangaYesNo"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Month", "Month"), "Month", "Month", "Month"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Notes", "Notes"), "Notes", "", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Number", "Number"), "ShadowNumber", "Number", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Penciller", "Penciller"), "Penciller", "Penciller", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Publisher", "Publisher"), "Publisher", "Publisher", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Rating", "Rating"), "Rating", "Rating", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("ReadPercentage", "Read Percentage"), "ReadPercentage", "ReadPercentage", "ReadPercentage"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("ReleasedTime", "Released"), "ReleasedTime", "Released", "DateTime"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Review", "Review"), "Review", "", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("ScanInformation", "Scan Information"), "ScanInformation", "ScanInformation", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Series", "Series"), "ShadowSeries", "Series", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("SeriesComplete", "Series Complete"), "SeriesComplete", "SeriesComplete", "YesNo"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("SeriesGroup", "Series Group"), "SeriesGroup", "SeriesGroup", "String"))
FIELDS.append(TemplateItem("Start Month", "StartMonth", "StartMonth", "Month"))
FIELDS.append(TemplateItem("Start Year", "StartYear", "StartYear", "Year"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("StoryArc", "Story Arc"), "StoryArc", "StoryArc", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Summary", "Summary"), "Summary", "", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Tags", "Tags"), "Tags", "Tags", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Teams", "Teams"), "Teams", "Teams", "MultipleValue"))
FIELDS.append(TemplateItem("Text", "Text", "", "Text"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Title", "Title"), "ShadowTitle", "Title", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("VolumeAsText", "Volume"), "ShadowVolume", "Volume", "Number"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Web", "Web"), "Web", "Web", "String"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("Writer", "Writer"), "Writer", "Writer", "MultipleValue"))
FIELDS.append(TemplateItem(localizer.get_comic_field_from_comicbook_dialog("YearAsText", "Year"), "ShadowYear", "Year", "Year"))

FIELDS.sort(key=lambda x: x.name)

#This contains the fields that are available to add into the template. Used for building the correct list of things later.
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

#This defines fields that are usable in the first letter selector. Typically string fields.
first_letter_fields = ['AlternateSeries', 
                       'Imprint', 
                       'Publisher', 
                       'ShadowSeries']

#This contains the fields usable in the selectors for condtitional. Should basically contain everything but Conditional with the additon of Text\
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

#These are the fields useable in the exclude rules.
exclude_rule_fields = ['AgeRating', 'AlternateCount', 'AlternateNumber', 
                       'AlternateSeries', 'BlackAndWhite', 'BookAge', 
                       'BookCollectionStatus', 'BookCondition', 
                       'BookLocation', 'BookNotes', 'BookOwner', 'BookPrice', 
                       'BookStore', 'Characters', 'Checked', 'Colorist', 
                       'CommunityRating', 'Count', 'CoverArtist', 'Editor', 
                       'FileDirectory', 'FileFormat', 'FileName', 
                       'FileIsMissing', 'FileNameWithExtension', 'FilePath', 
                       'FileSize', 'Format', 'Genre', 'HasBeenOpened', 'HasBeenRead', 'ISBN', 'Imprint', 'Inker', 'LanguageAsText', 'Letterer', 'Locations', 'MainCharacterOrTeam', 'Manga', 'Month', 'Notes', 'Number', 'Penciller', 'Publisher', 'Rating', 'ReadPercentage', 'Review', 'ScanInformation', 'Series', 'SeriesComplete', 'SeriesGroup', 'StoryArc', 'Summary', 'Tags', 'Teams', 'Title', 'Volume', 'Web', 'Writer', 'Year']

#These are the fields that require the multiple value treatment.
multiple_value_fields = ["AlternateSeries", "Characters", "Colorist", "CoverArtist", "Editor", "Genre", "Inker", "Letterer", "Locations", "Penciller", "ScanInformation", "Tags", "Teams", "Writer"]







rules_fields = ['BookNotes']



