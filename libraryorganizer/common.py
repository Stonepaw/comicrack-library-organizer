class Mode(object):
    """ An Enum object to compare the Mode of the script"""
    Move = "Move"
    Copy = "Copy"
    Simulate = "Simulate"

class BookMoverException(Exception):
    def __init__(self, msg):
        self.message = msg

class MoveSkippedException(BookMoverException):
    pass


class MoveFailedException(BookMoverException):
    pass

class DuplicateExistsException(BookMoverException):
    pass


class BookToMove(object):
    """A wrapper class for a ComicBook object to provide more information
    during a move operation
    """

    def __init__(self, book, path, profile_index, failed_fields):
        self.book = book
        self.path = path
        self.profile_index = profile_index
        self.failed_fields = failed_fields
        self.duplicate_different_extension = False
        self.duplicate_ext_files = []


class DuplicateAction(object):
    Overwrite = 1
    Cancel = 2
    Rename = 3


class DuplicateResult(object):
    def __init__(self, action, always_do_action):
        self.action = action
        self.always_do_action = always_do_action