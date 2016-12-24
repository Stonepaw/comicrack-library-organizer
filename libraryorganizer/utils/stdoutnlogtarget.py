""" Creates and adds a custom target to Nlog to support python stdout.

This file *must* be imported before nlog starts logging.

Copyright 2016 Stonepaw

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
import sys

clr.AddReference("NLog")
from System import Exception, Array
from NLog import LogLevel, LogManager, LogEventInfo
from NLog.Targets import TargetWithLayout
from NLog.Config import LoggingRule
from NLog.Common import AsyncLogEventInfo

class ComicRackNLogTarget(TargetWithLayout):
    """ A custom NLog target that outputs to the python stdout."""

    #Annoyingly we can't override overloaded methods correctly so this code is a lot more complex then needed.
    def Write(self, logevent):
        if type(logevent) is AsyncLogEventInfo:
            try:
                self.MergeEventProperties(logevent.LogEvent)
                self.Write(logevent.LogEvent)
                logevent.Continuation(None)
            except Exception as e:
                if (e.MustBeRethrown()):
                    raise
                logevent.Continuation(e)
        elif type(logevent) is LogEventInfo:
            sys.stdout.write(self.Layout.Render(logevent) + "\n")
        else:
            print "is array"
            for event in logevent:
                self.Write(logevent)


target = ComicRackNLogTarget()
target.Layout = "${message}"
LogManager.Configuration.AddTarget("StdOut", target)
LogManager.Configuration.LoggingRules.Add(LoggingRule("*", LogLevel.Info, target))
LogManager.Configuration.Reload()
