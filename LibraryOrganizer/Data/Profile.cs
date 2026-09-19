using System.Collections.Generic;

namespace LibraryOrganizer.Data
{
    public class Profile
    {
        /// <summary>
        /// If the copied book should be added to the library when in the operation is copy mode.
        /// </summary>
        public bool AddCopyToLibrary { get; set; } = true;

        /// <summary>
        /// Uses the single value from a multiple value field without asking when true.
        /// </summary>
        public bool AutoSelectSingleMultiValueField { get; set; } = true;

        /// <summary>
        /// Automatically space fields in the template when inserting into the template.
        /// </summary>
        public bool AutoSpaceFields { get; set; } = true;

        /// <summary>
        /// The base folder destination for this profile.
        ///
        /// Only needed when folder organization is enabled.
        /// </summary>
        public string BaseFolder { get; set; } = string.Empty;

        /// <summary>
        /// Copies the fileless book thumbnail to the file/folder name.
        ///
        /// Fileless books will be ignored when false.
        /// </summary>
        public bool CopyFilelessBookThumbnail { get; set; }

        /// <summary>
        /// The image format to use for fileless book thumbnails.
        ///
        /// <see cref="CopyFilelessBookThumbnail"/>
        /// </summary>
        public string CopyFilelessBookThumbnailFormat { get; set; } = ".jpg";

        /// <summary>
        /// When overwriting an existing file, copy the read percentage from the old file to the
        /// replacement file.
        /// </summary>
        public bool CopyReadPercentageToReplacement { get; set; } = true;

        /// <summary>
        /// Replacement values for empty field data.
        /// </summary>
        public List<EmptyFieldReplacement> EmptyFieldReplacements { get; } =
            new List<EmptyFieldReplacement>();

        /// <summary>
        /// Replace empty folder names in the template with this value. Empty folders are removed
        /// in the path when this is an empty string.
        /// </summary>
        public string EmptyFolderNameReplacement { get; set; } = string.Empty;

        /// <summary>
        /// When enabled fail an operation when a configured empty field value is encountered.
        ///
        /// <see cref="FailOperationOnEmptyValueFields"/>
        /// <see cref="FailOperationOnEmptyValueDestinationFolder"/>
        /// </summary>
        public bool FailOperationOnEmptyValue { get; set; }

        /// <summary>
        /// The destination folder to move/copy a field to when a field is empty with the failed empty
        /// configuration is enabled.
        ///
        /// <see cref="FailOperationOnEmptyValue"/>
        /// <see cref="FailOperationOnEmptyValueFields"/>
        /// </summary>
        public string FailOperationOnEmptyValueDestinationFolder { get; set; } = string.Empty;

        /// <summary>
        /// The configuration for which empty fields contribute to a failed operation.
        ///
        /// Only applicable when <see cref="FailOperationOnEmptyValue"/> is enabled.
        ///
        /// <see cref="FailOperationOnEmptyValue"/>
        /// <see cref="FailOperationOnEmptyValueDestinationFolder"/>
        /// </summary>
        public List<FailedEmptyField> FailOperationOnEmptyValueFields { get; } =
            FailedEmptyField.DefaultList();

        /// <summary>
        /// Enables moving/copying failed empty fields to a specific folder.
        /// </summary>
        public bool FailOperationOnEmptyValueUseDestinationFolder { get; set; }

        /// <summary>
        /// The file template to use during the operation if file naming is enabled.
        ///
        /// <see cref="UseFileNaming"/>
        /// </summary>
        public string FileTemplate { get; set; } = string.Empty;

        /// <summary>
        /// The folder template to use during the operation if folder organization is enabled.
        ///
        /// <see cref="UseFolderOrganization"/>
        /// </summary>
        public string FolderTemplate { get; set; } = string.Empty;

        /// <summary>
        /// Configurable replacements to use for illegal characters when inserted into the template.
        /// </summary>
        public List<IllegalCharacterReplacement> IllegalCharacterReplacements { get; } =
            new List<IllegalCharacterReplacement>
            {
                new IllegalCharacterReplacement("?", ""),
                new IllegalCharacterReplacement("/", ""),
                new IllegalCharacterReplacement("\\", ""),
                new IllegalCharacterReplacement("*", ""),
                new IllegalCharacterReplacement(":", " -"),
                new IllegalCharacterReplacement("<", "["),
                new IllegalCharacterReplacement(">", "]"),
                new IllegalCharacterReplacement("|", "!"),
                new IllegalCharacterReplacement("\"", "'"),
            };

        /// <summary>
        /// Replacements to use for month numbers when inserted into the template.
        /// </summary>
        public List<MonthReplacement> MonthReplacements { get; } =
            new List<MonthReplacement>
            {
                new MonthReplacement(1, "January"),
                new MonthReplacement(2, "February"),
                new MonthReplacement(3, "March"),
                new MonthReplacement(4, "April"),
                new MonthReplacement(5, "May"),
                new MonthReplacement(6, "June"),
                new MonthReplacement(7, "July"),
                new MonthReplacement(8, "August"),
                new MonthReplacement(9, "September"),
                new MonthReplacement(10, "October"),
                new MonthReplacement(11, "November"),
                new MonthReplacement(12, "December"),
                new MonthReplacement(13, "Spring"),
                new MonthReplacement(14, "Summer"),
                new MonthReplacement(15, "Fall"),
                new MonthReplacement(16, "Winter"),
            };

        /// <summary>
        /// The name of this profile
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Replaces multiple spaces in the generated file and folders with a single space when true.
        /// </summary>
        public bool NormalizeMultipleSpaces { get; set; } = true;

        /// <summary>
        /// The operation mode to use when this profile is run.
        /// </summary>
        public OperationMode OperationMode { get; set; } = OperationMode.Move;

        /// <summary>
        /// Remove empty folders from the source locations when moving the last file from it.
        /// </summary>
        public bool RemoveEmptyFolders { get; set; } = true;

        /// <summary>
        /// The list of folders to exclude when removing empty source folders
        /// </summary>
        public List<string> RemoveEmptyFoldersExclusions { get; } = new List<string>();

        /// <summary>
        /// If file naming should be used during the operation
        /// </summary>
        public bool UseFileNaming { get; set; } = true;

        /// <summary>
        /// If folder organization should be used during the operation
        /// </summary>
        public bool UseFolderOrganization { get; set; } = true;

        /// <summary>
        /// The version of this profile.
        ///
        /// This is used during deserialization to determine if the profile needs to be updated to
        /// the most recent version.
        /// </summary>
        public int Version { get; set; } = 0;
    }
}
