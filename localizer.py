from locommon import comic_fields, library_organizer_fields

translated_yes_no_operators = None
translated_manga_yes_no_operators = None
translated_comic_fields = None
translated_string_operators = None
translated_numeric_operators = None
translated_bool_operators = None


def get_comic_fields():
    """Gets a dict (key=translations, values=propertyname) of all the defined comic fields and library organizer fields"""
    #ComicBook field names are contained in the comicrack translation file named columns
    #All string fields are stored with the Key as the name.
    #Most non-string fields are stored with the key NameAsText
    #Certain non-string fields are stored with the Key as the name.
    #For most non string names it has to be retrived with the key NameAsText
    #ComicRack.Localize(Resourcekey, namekey, fail text)
    global translated_comic_fields
    if translated_comic_fields is not None:
        return translated_comic_fields

    translated_fields = {}
        
    #Get the comic fields
    for name in comic_fields:
        s = ComicRack.Localize("Columns", name, "FAILED")
        if not s == "FAILED":
            translated_fields[s] = name
            continue

        s = ComicRack.Localize("Columns", name + "AsText", "FAILED")
        if not s == "FAILED":
            translated_fields[s] = name
            continue
        translated_fields[camel_case_to_spaced(name)] = name

    #Get the library organizer fields
    for name in library_organizer_fields:
        translated_fields[ComicRack.Localize("Script.LibraryOrganizer", name, camel_case_to_spaced(name))] = name

    translated_comic_fields = translated_fields

    return translated_fields


def get_string(key, text):
    """Returns a localized string using the key parameter. If the key does not exist the text is returned"""
    return ComicRack.Localize("Script.LibraryOrganizer", key, text)


def get_exclude_rule_string_operators():

    global translated_string_operators

    if translated_string_operators is not None:
        return translated_string_operators

    operators = ["is", "contains", "contains any of", "contains all of", "starts with", "ends with", "list contains", "regular expression"]
    s = ComicRack.Localize("Matchers", "StringOperators", "is|contains|contains any of|contains all of|starts with|ends with|list contains|regular expression").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators)) if operators[i] is not "list contains"}
    translated_string_operators = translated_operators
    return translated_operators
    

def get_exclude_rule_numeric_operators():

    global translated_numeric_operators

    if translated_numeric_operators is not None:
        return translated_numeric_operators

    operators = ["is", "is greater", "is smaller", "is in the range"]
    s = ComicRack.Localize("Matchers", "NumericOperators", "is|is greater|is smaller|is in the range").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators))}
    translated_numeric_operators = translated_operators
    return translated_operators


def get_exclude_rule_bool_operators():

    global translated_bool_operators

    if translated_bool_operators is not None:
        return translated_bool_operators

    operators = ["is True", "is False"]
    s = ComicRack.Localize("Matchers", "TrueFalseOperators", "is True|is False").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators))}
    translated_bool_operators = translated_operators
    return translated_operators


def get_exclude_rule_yes_no_operators():

    global translated_yes_no_operators

    if translated_yes_no_operators is not None:
        return translated_yes_no_operators

    operators = ["is Yes", "is No", "is Unknown"]
    s = ComicRack.Localize("Matchers", "YesNoOperators", "is Yes|is No|is Unknown").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators))}
    translated_yes_no_operators = translated_operators
    return translated_operators


def get_exclude_rule_manga_yes_no_operators():

    global translated_manga_yes_no_operators

    if translated_manga_yes_no_operators is not None:
        return translated_manga_yes_no_operators

    operators = ["is Yes", "is Yes (Right to Left)", "is No", "is Unknown"]
    s = ComicRack.Localize("Matchers", "MangaYesNoOperators", "is Yes|is Yes (Right to Left)|is No|is Unknown").split("|")

    #Can't do list contains so make sure we don't use that operator
    translated_operators = {s[i] : operators[i] for i in range(len(operators))}
    translated_manga_yes_no_operators = translated_operators
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