using System.Collections.Generic;

namespace LibraryOrganizer.Data
{
    public enum MetadataRuleGroupOperator
    {
        Any,
        All,
    }

    public class MetadataRuleGroup
    {
        public MetadataRuleGroupOperator Operator { get; set; } = MetadataRuleGroupOperator.Any;

        public List<IMetadataRule> Rules { get; } = new List<IMetadataRule>();
    }
}
