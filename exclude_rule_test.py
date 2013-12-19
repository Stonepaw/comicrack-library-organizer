import wpf
import clr
clr.AddReferenceByPartialName("Xceed.Wpf.Toolkit")
from System.Windows import Window
from System.Collections.Generic import List
import ComicRack
c = ComicRack.ComicRack()
import localizer
localizer.ComicRack = c
import fieldmappings
import locommon
locommon.ComicRack = c

import Xceed.Wpf.Toolkit


from exclude_rule_view_models import ExcludeRuleViewModel, ExcludeRuleCollectionViewModel, RulesTemplateSelector
from exclude_rules import ExcludeRuleBase, ExcludeRuleCollection

class ExcludeRuleTest(Window):
    def __init__(self):
        self.DataContext = self
        e = ExcludeRuleCollection()
        e.append(ExcludeRuleBase())
        self.Resources.Add("RulesTemplateSelector", RulesTemplateSelector())

        self.exclude_rules_view_model = ExcludeRuleCollectionViewModel(e)
        wpf.LoadComponent(self, 'exclude_rule_test.xaml')
    
    def Button_Click(self, sender, e):
        sender.ContextMenu.IsEnabled = True
        sender.ContextMenu.PlacementTarget = sender
        sender.ContextMenu.IsOpen = True


if __name__ == '__main__':
   
    e = ExcludeRuleTest()
    e.ShowDialog()




