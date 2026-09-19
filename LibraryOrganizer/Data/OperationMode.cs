using System;
using System.Xml.Serialization;

namespace LibraryOrganizer.Data
{
    public enum OperationMode
    {
        [XmlEnum(Name = "Move")]
        Move,

        [XmlEnum(Name = "Copy")]
        Copy,

        [XmlEnum(Name = "Simulate")]
        Simulate,

        [Obsolete("Legacy value from pre version 2.0")]
        [XmlEnum(Name = "Test")]
        Test,
    }
}
