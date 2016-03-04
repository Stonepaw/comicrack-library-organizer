"""
Copyright 2016 Andrew Feltham (Stonepaw)

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

from System import UnauthorizedAccessException
from System.IO import DirectoryInfo, Path, DirectoryNotFoundException, IOException
from System.Security import SecurityException

clr.AddReference("NLog.dll")
from NLog import LogManager

# Module Logger
_LOG = LogManager.GetLogger("FileUtils")


def delete_empty_folders(directory, excluded_folders=None):
    """Recursively deletes directories until an non-empty directory is
    found or the directory is in the excluded list.

    Args:
        directory: A DirectoryInfo object which starts at the first folder
            to remove.
        excluded_folders: A list of folders to never delete
    """
    _LOG.Trace("Starting to delete empty folders.")
    directory.Refresh()
    if not directory.Exists:
        return
    # Only delete if no file or folder and not in folder never to delete
    if excluded_folders is not None and directory.FullName in excluded_folders:
        return
    parent = directory.Parent
    try:
        directory.Delete()
    except (UnauthorizedAccessException, DirectoryNotFoundException,
            IOException, SecurityException) as ex:
        _LOG.Warn("Unable to delete {0}: {1}", directory.FullName,
                  ex.Message)
    else:
        _LOG.Info("Deleted empty folder {0}", directory.FullName)
        delete_empty_folders(parent)


def get_files_with_different_ext(full_path):
    """Looks for files in the same directory of the file with the same name
    but different extensions.

    Args:
        full_path: The full path of the file.

    Returns:
        A list or System.Array containing any FileInfo objects that were
        found, possibly empty.
    """
    d = DirectoryInfo(Path.GetDirectoryName(full_path))
    if d.Exists:
        file_name = Path.GetFileNameWithoutExtension(full_path)
        files = d.GetFiles(file_name + ".*")
        return files
    return []


def create_folder(new_directory):
    """Creates the directory tree.

    If the profile mode is simulate it doesn't create any folders, instead
    it saves the folders that would need to be created to _created_paths

    Args:
        new_directory: A DirectoryInfo object of the folder that is needed.
    Raises:
        From Directory.Create()
        IOException,
        UnauthorizedAccessException,
        ArgumentException,
        NotSupportedException,
    """
    _LOG.Trace("Starting create folder: {0}", new_directory.FullName)
    if new_directory.Exists:
        _LOG.Info("Folder already exists")
        return
    new_directory.Create()
    new_directory.Refresh()
    _LOG.Info("Created folder {0}", new_directory.FullName)
