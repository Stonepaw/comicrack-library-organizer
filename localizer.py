from locommon import comic_fields, library_organizer_fields
import clr


#Caching because these functions could be called in different modules.
TRANSLATED_YES_NO_OPERATORS = None
TRANSLATED_MANGA_YES_NO_OPERATORS = None
TRANSLATED_COMIC_FIELDS = None
TRANSLATED_STRING_OPERATORS = None
TRANSLATED_CONDITIONAL_STRING_OPERATORS = None
TRANSLATED_DATE_OPERATORS = None
TRANSLATED_NUMERIC_OPERATORS = None
TRANSLATED_BOOL_OPERATORS = None


def get_comic_fields():
    """Gets a dict (key=propertyname, values=translations) of all the defined comic fields and library organizer fields"""
    #ComicBook field names are contained in the comicrack translation file named columns
    #All string fields are stored with the Key as the name.
    #Most non-string fields are stored with the key NameAsText
    #Certain non-string fields are stored with the Key as the name.
    #For most non string names it has to be retrived with the key NameAsText
    #ComicRack.Localize(Resourcekey, namekey, fail text)
    global TRANSLATED_COMIC_FIELDS
    if TRANSLATED_COMIC_FIELDS is not None:
        return TRANSLATED_COMIC_FIELDS

    translated_fields = {}
        
    #Get the comic fields
    for name in comic_fields:
        if name.startswith("Shadow"):
            name = name[6:]

        s = ComicRack.Localize("Columns", name, "FAILED")
        if not s == "FAILED":
            translated_fields[name] = s
            continue

        s = ComicRack.Localize("Columns", name + "AsText", "FAILED")
        if not s == "FAILED":
            translated_fields[name] = s
            continue

        translated_fields[name] = camel_case_to_spaced(name)

    #Get the library organizer fields
    for name in library_organizer_fields:
        translated_fields[name] = ComicRack.Localize("Script.LibraryOrganizer", name, camel_case_to_spaced(name))

    TRANSLATED_COMIC_FIELDS = translated_fields

    return translated_fields


def get_comic_field_from_columns(field, backup_text):
    """Returns a localized comic field string using ComicRack's built-in translations.
        
    Args:
        field: The field to localize.
        backup_text: The text to return if no localization is available.

    Returns:
        A string of either the localized field name or the backup_text.
    """
    s = ComicRack.Localize("Columns", field, "")
    if s != "":
        return s
    return ComicRack.Localize("Columns", field + "AsText", backup_text)


def get_comic_field_from_comicbook_dialog(field, backup_text):
    """Returns a localized comic field string using ComicRack's built-in translations.
        
    Args:
        field: The field to localize.
        backup_text: The text to return if no localization is available.

    Returns:
        A string of either the localized field name or the backup_text.
    """
    return ComicRack.Localize("ComicBookDialog", "label" + field, backup_text).strip(":")
    


def get_string(key, backup_text):
    """Returns a localized string from Library Organizers translation file.
        
    Args:
        key: The key to localize.
        backup_text: The text to return if no localization is available.

    Returns:
        A string of either the localized text or the backup_text.
    """
    return ComicRack.Localize("Script.LibraryOrganizer", key, backup_text)


def get_exclude_rule_string_operators():

    global TRANSLATED_STRING_OPERATORS

    if TRANSLATED_STRING_OPERATORS is not None:
        return TRANSLATED_STRING_OPERATORS

    operators = ["is", "contains", "contains any of", "contains all of", "starts with", "ends with", "list contains", "regular expression"]
    s = ComicRack.Localize("Matchers", "StringOperators", "is|contains|contains any of|contains all of|starts with|ends with|list contains|regular expression").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators)) if operators[i] != "list contains"}
    TRANSLATED_STRING_OPERATORS = translated_operators
    return translated_operators
    

def get_exclude_rule_numeric_operators():

    global TRANSLATED_NUMERIC_OPERATORS

    if TRANSLATED_NUMERIC_OPERATORS is not None:
        return TRANSLATED_NUMERIC_OPERATORS

    operators = ["is", "greater", "smaller", "in the range"]
    s = ComicRack.Localize("Matchers", "NumericOperators", "is|is greater|is smaller|is in the range").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {operators[i] : s[i] for i in range(len(operators))}
    TRANSLATED_NUMERIC_OPERATORS = translated_operators
    return translated_operators


def get_exclude_rule_bool_operators():

    global TRANSLATED_BOOL_OPERATORS

    if TRANSLATED_BOOL_OPERATORS is not None:
        return TRANSLATED_BOOL_OPERATORS

    operators = ["is True", "is False"]
    s = ComicRack.Localize("Matchers", "TrueFalseOperators", "is True|is False").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators))}
    TRANSLATED_BOOL_OPERATORS = translated_operators
    return translated_operators


def get_exclude_rule_yes_no_operators():

    global TRANSLATED_YES_NO_OPERATORS

    if TRANSLATED_YES_NO_OPERATORS is not None:
        return TRANSLATED_YES_NO_OPERATORS

    operators = ["Yes", "No", "Unknown"]
    s = ComicRack.Localize("Matchers", "YesNoOperators", "is Yes|is No|is Unknown").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {operators[i] : s[i] for i in range(len(operators))}
    TRANSLATED_YES_NO_OPERATORS = translated_operators
    return translated_operators


def get_yes_no_operators():
    """
    Retrieves the yes no operators from ComicRack's built in translations.

    Returns:
        A dict with the key as the translations in the form of "is yes" 
        and the value as the yes no value.
    """
    global TRANSLATED_YES_NO_OPERATORS

    if TRANSLATED_YES_NO_OPERATORS is not None:
        return TRANSLATED_YES_NO_OPERATORS

    operators = ["Yes", "No", "Unknown"]
    s = ComicRack.Localize("Matchers", "YesNoOperators", "is Yes|is No|is Unknown").split("|")

    translated_operators = {operators[i] : s[i] for i in range(len(operators))}
    TRANSLATED_YES_NO_OPERATORS = translated_operators
    return translated_operators


def get_manga_yes_no_operators():
    """
    Retrieves the manga yes no operators from ComicRack's built in translations.

    Returns:
        A dict with the key as the translations and the value as the english translation
    """
    global TRANSLATED_MANGA_YES_NO_OPERATORS

    if TRANSLATED_MANGA_YES_NO_OPERATORS is not None:
        return TRANSLATED_MANGA_YES_NO_OPERATORS

    operators = ["Yes", "Yes (Right to Left)", "No", "Unknown"]
    s = ComicRack.Localize("Matchers", "MangaYesNoOperators", "is Yes|is Yes (Right to Left)|is No|is Unknown").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {operators[i] : s[i] for i in range(len(operators))}
    TRANSLATED_MANGA_YES_NO_OPERATORS = translated_operators
    return translated_operators

def get_conditional_string_operators():

    global TRANSLATED_CONDITIONAL_STRING_OPERATORS

    if TRANSLATED_CONDITIONAL_STRING_OPERATORS is not None:
        return TRANSLATED_CONDITIONAL_STRING_OPERATORS

    operators = ["is", "contains", "any", "all", "starts", "ends", "list contains", "regex"]
    s = ComicRack.Localize("Matchers", "StringOperators", "is|contains|contains any of|contains all of|starts with|ends with|list contains|regular expression").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {operators[i] : s[i] for i in range(len(operators)) if operators[i] != "list contains"}
    TRANSLATED_CONDITIONAL_STRING_OPERATORS = translated_operators
    return translated_operators

def get_date_operators():

    global TRANSLATED_DATE_OPERATORS

    if TRANSLATED_DATE_OPERATORS is not None:
        return TRANSLATED_DATE_OPERATORS

    operators = ["is", "after", "before"]
    s = ComicRack.Localize("Matchers", "DateOperators", "is|is after|is before|is in the last|is in the range").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {operators[i] : s[i] for i in range(3)}
    TRANSLATED_DATE_OPERATORS = translated_operators
    return translated_operators


def camel_case_to_spaced(s):
    """Converts a CamelCaseString to a Camel Case String"""
    s2 = s[0]
    l = len(s)
    for i in range(1, len(s)):
        #If the previous character is an upper than it is most likely an abbreviation
        if s[i].isupper() and not s[i-1].isupper():
            s2 = s2 + " " + s[i]
        else:
            s2 = s2 + s[i]
    return s2