using System;
using System.Linq;
using System.Windows.Forms;
using AlbaLibrary.IO;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using AlbaLibrary.Interpolation;
using AlbaLibrary.XL;

namespace AlbaGraphEngineering
{
    public static class myIO
    {
        /// <summary>
        /// Assign directories 
        /// </summary>
        /// <param name="myFiles">list of files and directories sorted by directories</param>
        /// <param name="updateIndex"></param>
        /// <param name="myFolderComboBox">Combobox that contains the list of directories</param>
        /// <param name="myListBox">List box that contains the list of data</param>
        public static void AssignFolders(SortedDirectoryFiles myFiles, bool toUpdate, ComboBox myFolderComboBox, ListBox myListBox)
        {
            try
            {
                if (myFiles.Exists) // check if there are directory files
                {
                    
                    // clear current data from controls
                    myFolderComboBox.Items.Clear();
                    myListBox.Items.Clear();
                    // select the current folder index
                    var currentIndex = myFolderComboBox.SelectedIndex;

                    // add sub directories to the folder combobox
                    #region
                    // add any sub directories below the main directory
                    if (myFiles.SubDirectories != null)
                    { // Add an option for all - shows all data 
                        myFolderComboBox.Items.Add("All");
                        // add the top level folder to the selection drop down list
                        myFolderComboBox.Items.Add(myFiles.TopDirectory.Name);
                        myFolderComboBox.Items.AddRange(myFiles.SubDirectories.Select(i => i.Name).ToArray());
                    }
                    else // add the top level folder to the selection drop down list
                        myFolderComboBox.Items.Add(myFiles.TopDirectory.Name);
                    #endregion

                    // check if the top directory has files to display in list box
                    // else select the first directory with files
                    #region
                    if (toUpdate)
                    {
                        // if there are files in the first directory select it else
                        // look for the first list of files in the directory and select
                        if (myFiles.TopFiles.Length > 0)
                            myFolderComboBox.SelectedIndex = 0;
                        else
                        {
                            // obtain the number of files in each subdirectory
                            myFolderComboBox.SelectedIndex = myFiles.SubFiles
                                .Select(i => i.Length).ToList()
                                .FindIndex(i => i != 0) + 2; // index is offset by one due to the top level
                        }
                    }
                    else
                        myFolderComboBox.SelectedIndex = currentIndex;
                    #endregion
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Assign Folders Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Switched based on the files in the input directory string
        /// </summary>
        /// <param name="directoryString">the string of the directory</param>
        /// <param name="index">the selected subdirectory index</param>
        public static void SwitchFiles(string myDirectoryString, int index, ComboBox myFolderComboBox, ListBox myListBox, bool summaryRead, bool nameSort)
        {
            try
            {
                // load the file
                var fileList = new SortedDirectoryFiles(myDirectoryString, "*.xls");

                // switch based on the selected file list
                switch (index)
                {
                    case 0:
                        myListBox.Items.Clear();

                        var tempFiles = new List<FileInfo>();

                        tempFiles.AddRange(fileList.TopFiles);

                        if (fileList.SubFiles != null)
                        {
                            foreach (var tempSub in fileList.SubFiles)
                                tempFiles.AddRange(tempSub);
                        }

                        AssignFiles(tempFiles, false, myListBox, summaryRead, nameSort);
                        break;
                    case 1:
                        AssignFiles(fileList.TopFiles, true, myListBox, summaryRead, nameSort);
                        break;
                    default:
                        AssignFiles(fileList.SubFiles[index - 2], true, myListBox, summaryRead, nameSort);
                        break;
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Switch Files", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Assign the files to the list box
        /// </summary>
        /// <param name="fileArray">files to assign</param>
        public static void AssignFiles(IEnumerable<FileInfo> files, bool clear, ListBox myListBox, bool summaryRead, bool nameSort)
        {
            try
            {
                if (files != null)
                {
                    if (clear) // clear the current data in the list box 
                        myListBox.Items.Clear();

                    // Create the appropriate file list
                    var myFileList = ReturnAppropriateFiles(files, summaryRead, nameSort)
                        .Select(i => Path.GetFileNameWithoutExtension(i.Name))
                        .ToArray();

                    // assign the file names to the listbox based on summary or trans raw
                    myListBox.Items.AddRange(myFileList);            
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Assign Files Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Separate the transraw and summary files
        /// </summary>
        /// <param name="fileArray">files collection to be sorted</param>
        /// <param name="summaryRead">Return summary files (true for yes and false to return transraw)</param>
        /// <param name="nameSort">Sort by name (true for yes or false for by element number)</param>
        /// <returns></returns>
        public static IEnumerable<FileInfo> ReturnAppropriateFiles(IEnumerable<FileInfo> fileArray, bool summaryRead, bool nameSort)
        {
            try
            {
                if (summaryRead)
                {
                    return fileArray.ToList()
                       .Select(i => i)
                       .Where(j => j.Name.ToLower().Contains("summary"))
                       .ToArray();
                }
                else
                {
                    var files = fileArray.ToList()
                       .Select(i => i)
                       .Where(j => !j.Name.ToLower().Contains("summary") && j.Name.Contains('_'))
                       .ToArray();

                    if (!nameSort)
                        return SortByElementNumber(files);
                    

                    return files;
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Return Appropriate Files", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public static IEnumerable<FileInfo> FilterFiles(IEnumerable<FileInfo> files, int start, int end, int step)
        {
            try
            {
                // offset to zero
                start = start - 1;

                return files.ToList()
                    .GetRange(start, end - start)
                    .Where(j => Array.IndexOf(files.ToArray(), j) % step == 0)
                    .ToArray();
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Filter Files", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public static IEnumerable<FileInfo> CreateSelectedFiles(SortedDirectoryFiles files, ComboBox myFolderList, bool summaryRead, bool nameSort)
        {
            try
            {
                switch (myFolderList.SelectedIndex)
                {
                    case 0:
                        var tempFiles = new List<FileInfo>();

                        tempFiles.AddRange(files.TopFiles);

                        if (files.SubFiles != null)
                        {
                            foreach (var tempSub in files.SubFiles)
                                tempFiles.AddRange(tempSub);
                        }
                        return ReturnAppropriateFiles(tempFiles, summaryRead, nameSort);
                    case 1:
                        return ReturnAppropriateFiles(files.TopFiles, summaryRead, nameSort);
                    default:
                        return ReturnAppropriateFiles(files.SubFiles[myFolderList.SelectedIndex - 2], summaryRead, nameSort);
                }
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Create Selected Files Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public static IEnumerable<FileInfo> SortByElementNumber(IEnumerable<FileInfo> files)
        {
            try
            {
                var data = files.Select(i => i).ToArray();
                var elementList = data.Select(i => i).Where(j => j.Name.Contains('_')).Select(k => k.Name.Substring(k.Name.LastIndexOf('_'), k.Name.Length - k.Name.LastIndexOf('_'))).ToArray();
                Array.Sort(elementList, data, AlbaLibrary.MyStrComp.MyStrComparer<string>(AlbaLibrary.MyStrCompare.ByStringLogicalComparer));
                return data;
            }
            catch (Exception _e)
            {
                MessageBox.Show(_e.Message, "Sort By Element Number", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return files.Select(i => i).ToArray();
            }
        }

        // outputs

      

    }

}
