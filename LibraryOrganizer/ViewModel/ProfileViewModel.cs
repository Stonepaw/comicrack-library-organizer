using System.Collections.Generic;
using System.ComponentModel;
using LibraryOrganizer.Data;

namespace LibraryOrganizer.ViewModel
{
    public class ProfileViewModel : ViewModelBase
    {
        private readonly Profile _profile;

        public ProfileViewModel(Profile profile)
        {
            _profile = profile;
        }

        /// <summary>
        /// If the copied book should be added to the library when in the operation is copy mode.
        /// </summary>
        public bool AddCopyToLibrary
        {
            get => _profile.AddCopyToLibrary;
            set =>
                SetProperty(_profile.AddCopyToLibrary, value, (v) => _profile.AddCopyToLibrary = v);
        }

        /// <summary>
        /// Uses the single value from a multiple value field without asking when true.
        /// </summary>
        public bool AutoSelectSingleMultiValueField
        {
            get => _profile.AutoSelectSingleMultiValueField;
            set =>
                SetProperty(
                    _profile.AutoSelectSingleMultiValueField,
                    value,
                    (v) => _profile.AutoSelectSingleMultiValueField = v
                );
        }

        /// <summary>
        /// Automatically space fields in the template when inserting into the template.
        /// </summary>
        public bool AutoSpaceFields
        {
            get => _profile.AutoSpaceFields;
            set =>
                SetProperty(_profile.AutoSpaceFields, value, (v) => _profile.AutoSpaceFields = v);
        }

        /// <summary>
        /// The base folder destination for this profile.
        ///
        /// Only needed when folder organization is enabled.
        /// </summary>
        public string BaseFolder
        {
            get => _profile.BaseFolder;
            set => SetProperty(_profile.BaseFolder, value, (v) => _profile.BaseFolder = v);
        }

        /// <summary>
        /// Copies the fileless book thumbnail to the file/folder name.
        ///
        /// Fileless books will be ignored when false.
        /// </summary>
        public bool CopyFilelessBookThumbnail
        {
            get => _profile.CopyFilelessBookThumbnail;
            set =>
                SetProperty(
                    _profile.CopyFilelessBookThumbnail,
                    value,
                    (v) => _profile.CopyFilelessBookThumbnail = v
                );
        }

        /// <summary>
        /// The image format to use for fileless book thumbnails.
        ///
        /// <see cref="CopyFilelessBookThumbnail"/>
        /// </summary>
        public string CopyFilelessBookThumbnailFormat
        {
            get => _profile.CopyFilelessBookThumbnailFormat;
            set =>
                SetProperty(
                    _profile.CopyFilelessBookThumbnailFormat,
                    value,
                    (v) => _profile.CopyFilelessBookThumbnailFormat = v
                );
        }

        /// <summary>
        /// When overwriting an existing file, copy the read percentage from the old file to the new file.
        /// </summary>
        public bool CopyReadPercentageToReplacement
        {
            get => _profile.CopyReadPercentageToReplacement;
            set =>
                SetProperty(
                    _profile.CopyReadPercentageToReplacement,
                    value,
                    (v) => _profile.CopyReadPercentageToReplacement = v
                );
        }

        /// <summary>
        /// Replacement values for empty field data.
        /// </summary>
        public List<EmptyFieldReplacement> EmptyFieldReplacements =>
            _profile.EmptyFieldReplacements;

        /// <summary>
        /// Replace empty folder names in the template with this value. Empty folders are removed
        /// when an empty string.
        /// </summary>
        public string EmptyFolderNameReplacement
        {
            get => _profile.EmptyFolderNameReplacement;
            set =>
                SetProperty(
                    _profile.EmptyFolderNameReplacement,
                    value,
                    (v) => _profile.EmptyFolderNameReplacement = v
                );
        }

        /// <summary>
        /// When enabled fail an operation when a configured empty field value is encountered.
        ///
        /// <see cref="FailOperationOnEmptyValueFields"/>
        /// <see cref="FailOperationOnEmptyValueDestinationFolder"/>
        /// <see cref="FailOperationOnEmptyValueDestinationFolderEnabled"/>
        /// </summary>
        public bool FailOperationOnEmptyValue
        {
            get => _profile.FailOperationOnEmptyValue;
            set
            {
                if (
                    SetProperty(
                        _profile.FailOperationOnEmptyValue,
                        value,
                        (v) => _profile.FailOperationOnEmptyValue = v
                    )
                )
                {
                    NotifyPropertyChanged(
                        nameof(FailOperationOnEmptyValueDestinationFolderEnabled)
                    );
                    NotifyPropertyChanged(nameof(FailOperationOnEmptyValueFieldsReadonly));
                }
            }
        }

        /// <summary>
        /// If the fail operation on empty value fields should be editable or readonly
        /// </summary>
        public bool FailOperationOnEmptyValueFieldsReadonly
        {
            get => !_profile.FailOperationOnEmptyValue;
        }

        /// <summary>
        /// The destination folder to move/copy a field to when a field is empty with the failed empty
        /// configuration is enabled.
        ///
        /// <see cref="FailOperationOnEmptyValue"/>
        /// <see cref="FailOperationOnEmptyValueFields"/>
        /// </summary>
        public string FailOperationOnEmptyValueDestinationFolder
        {
            get => _profile.FailOperationOnEmptyValueDestinationFolder;
            set =>
                SetProperty(
                    _profile.FailOperationOnEmptyValueDestinationFolder,
                    value,
                    (v) => _profile.FailOperationOnEmptyValueDestinationFolder = v
                );
        }

        /// <summary>
        /// If the failed empty folder controls are enabled.
        ///
        /// This is a combined boolean as it is only available when both FailEmptyValues and MoveFailed are enabled.
        /// </summary>
        public bool FailOperationOnEmptyValueDestinationFolderEnabled =>
            _profile.FailOperationOnEmptyValueUseDestinationFolder
            && _profile.FailOperationOnEmptyValue;

        /// <summary>
        /// The configuration for which empty fields contribute to a failed operation.
        ///
        /// Only applicable when <see cref="FailOperationOnEmptyValue"/> is enabled.
        ///
        /// <see cref="FailOperationOnEmptyValue"/>
        /// <see cref="FailOperationOnEmptyValueDestinationFolder"/>
        /// </summary>
        public List<FailedEmptyField> FailOperationOnEmptyValueFields =>
            _profile.FailOperationOnEmptyValueFields;

        /// <summary>
        /// Enables moving failed empty fields to a specific folder.
        /// </summary>
        public bool FailOperationOnEmptyValuesUseDestinationFolder
        {
            get => _profile.FailOperationOnEmptyValueUseDestinationFolder;
            set
            {
                if (
                    SetProperty(
                        _profile.FailOperationOnEmptyValueUseDestinationFolder,
                        value,
                        (v) => _profile.FailOperationOnEmptyValueUseDestinationFolder = v
                    )
                )
                {
                    NotifyPropertyChanged(
                        nameof(FailOperationOnEmptyValueDestinationFolderEnabled)
                    );
                }
            }
        }

        /// <summary>
        /// The file template to use during the operation if file naming is enabled.
        ///
        /// <see cref="UseFileNaming"/>
        /// </summary>
        public string FileTemplate
        {
            get => _profile.FileTemplate;
            set => SetProperty(_profile.FileTemplate, value, (v) => _profile.FileTemplate = v);
        }

        /// <summary>
        /// The folder template to use during the operation if folder organization is enabled.
        ///
        /// <see cref="UseFolderOrganization"/>
        /// </summary>
        public string FolderTemplate
        {
            get => _profile.FolderTemplate;
            set => SetProperty(_profile.FolderTemplate, value, (v) => _profile.FolderTemplate = v);
        }

        /// <summary>
        /// Configurable replacements to use for illegal characters when inserted into the template.
        /// </summary>
        public List<IllegalCharacterReplacement> IllegalCharacterReplacements =>
            _profile.IllegalCharacterReplacements;

        /// <summary>
        /// Replacements to use for month numbers when inserted into the template.
        /// </summary>
        public List<MonthReplacement> MonthReplacements => _profile.MonthReplacements;

        /// <summary>
        /// The name of this profile
        /// </summary>
        public string Name
        {
            get => _profile.Name;
            set => SetProperty(_profile.Name, value, (v) => _profile.Name = v);
        }

        /// <summary>
        /// Replaces multiple spaces in the generated file and folders with a single space when true.
        /// </summary>
        public bool NormalizeMultipleSpaces
        {
            get => _profile.NormalizeMultipleSpaces;
            set =>
                SetProperty(
                    _profile.NormalizeMultipleSpaces,
                    value,
                    (v) => _profile.NormalizeMultipleSpaces = v
                );
        }

        /// <summary>
        /// The operation mode to use when this profile is run.
        /// </summary>
        public OperationMode OperationMode
        {
            get => _profile.OperationMode;
            set => SetProperty(_profile.OperationMode, value, (v) => _profile.OperationMode = v);
        }

        /// <summary>
        /// Remove empty folders from the source locations when moving the last file from it.
        /// </summary>
        public bool RemoveEmptyFolders
        {
            get => _profile.RemoveEmptyFolders;
            set =>
                SetProperty(
                    _profile.RemoveEmptyFolders,
                    value,
                    (v) => _profile.RemoveEmptyFolders = v
                );
        }

        /// <summary>
        /// The list of folders to exclude when removing empty source folders
        /// </summary>
        public List<string> RemoveEmptyFoldersExclusions => _profile.RemoveEmptyFoldersExclusions;

        /// <summary>
        /// If file naming should be used during the operation
        /// </summary>
        public bool UseFileNaming
        {
            get => _profile.UseFileNaming;
            set => SetProperty(_profile.UseFileNaming, value, (v) => _profile.UseFileNaming = v);
        }

        /// <summary>
        /// If folder organization should be used during the operation
        /// </summary>
        public bool UseFolderOrganization
        {
            get => _profile.UseFolderOrganization;
            set =>
                SetProperty(
                    _profile.UseFolderOrganization,
                    value,
                    (v) => _profile.UseFolderOrganization = v
                );
        }
    }
}
