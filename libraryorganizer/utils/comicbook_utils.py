def convert_tags_to_list(tag_string):
    """ Converts the ComicBook comma seperated tag string into a list

    Args:
        tag_string: The string of the tags

    Returns: The list of tags

    """
    if tag_string.strip():
        return [tag.strip() for tag in tag_string.split(",")]
    else:
        return []

def copy_data_to_new_book(book, new_book):
    """This helper function copies all relevant metadata from a book to another book.

    Args:
        book (ComicBook): The ComicBook to copy from.
        new_book (ComicBook): The ComicBook to copy to.
    """
    new_book.SetInfo(book, False, False)  # This copies most fields
    new_book.CustomValuesStore = book.CustomValuesStore
    new_book.SeriesComplete = book.SeriesComplete  # Not copied by SetInfo
    new_book.Rating = book.Rating  # Not copied by SetInfo
    new_book.CustomThumbnailKey = book.CustomThumbnailKey