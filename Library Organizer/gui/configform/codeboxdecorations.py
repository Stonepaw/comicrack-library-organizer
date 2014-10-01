import clr
import re

clr.AddReference("System")
clr.AddReference("CodeBoxControl")

from System.Collections.Generic import List

from CodeBoxControl import Pair
from CodeBoxControl.Decorations import Decoration

from fieldmappings import FIELDS
from locommon import SUBSTITUTION_REGEX

class LibraryOrganizerNameDecoration(Decoration):
    """description of class"""

    fields = [t.template for t in FIELDS if t.template]

    def Ranges(self, text):
        pairs = List[Pair]()
        while True:
            m = SUBSTITUTION_REGEX.search(text)
            if m is None:
                break
            group = m.group('name')
            if group and (group in self.fields or group.lstrip('!?') in self.fields):
                pairs.Add(Pair(m.start('name'), len(group)))

            text = "".join((text[:m.start()], " " * len(m.group()),
                            text[m.end():]))
        return pairs

class LibraryOrganizerPrefixSuffixDecoration(Decoration):
    """description of class"""

    fields = [t.template for t in FIELDS if t.template]
    group_names = ('prefix', 'suffix')

    def Ranges(self, text):
        pairs = List[Pair]()
        while True:
            m = SUBSTITUTION_REGEX.search(text)
            if m is None:
                break
            for group_name in self.group_names:
                group = m.group(group_name)
                if group and m.group('name') in self.fields:
                    start = m.start(group_name)
                    for i in re.finditer("[^\\s]+", group):
                        pairs.Add(Pair(start + i.start(), len(i.group())))

            text = "".join((text[:m.start()], " " * len(m.group()),
                            text[m.end():]))
        return pairs

class LibraryOrganizerArgsDecoration(Decoration):
    """description of class"""

    fields = [t.template for t in FIELDS if t.template]

    def Ranges(self, text):
        pairs = List[Pair]()
        while True:
            m = SUBSTITUTION_REGEX.search(text)
            if m is None:
                break
            args = m.group('args')
            if args and (m.group('name') in self.fields or m.group('name').lstrip("!?") in self.fields):
                if args.startswith('('):
                    for arg in re.finditer("\(([^)]*)\)", args):
                        if arg:
                            pairs.Add(Pair(m.start('args') + arg.start(1), 
                                           len(arg.group(1))))
                else:
                    pairs.Add(Pair(m.start('args'), len(args)))
            text = "".join((text[:m.start()], " " * len(m.group()),
                            text[m.end():]))
        return pairs