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