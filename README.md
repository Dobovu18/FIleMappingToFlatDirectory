# FIleMappingToFlatDirectory

This is a project for a summer 2019 internship. To sum things up, I had to work out a way to rename over 9000 files given
information on a spreadsheet. This is the code, and a smaller test environment for it. I chose to do this in python.

The task:
    --A folder was pulled from a server.
    --The folder is filled with sub-folders which are named in accordance to some ID. This ID will be referred to was ID1.
    --In each of the ID1 folders is exactly one sub-folder which corresponds to a second ID, called ID2.
    --Under the ID2 folders are files, which have the same name as their parent folder, or a duplicate name (e.g the name of
      of the folder followed by a "(2)" etc.)
    --The task is to rename all the files to a new title determined by information in a spreadsheet.
    --The spreadsheet contains information relating the ID1s to a title, and the file type.
    --But the files are named ID2, so I have to define a way to map the titles in a spreadsheet to the ID2-named file in the
      spreadsheet.

The steps I need to take:
    --Given a file name in our Directory, find the matching ID1 in the spreadsheet ✓
    --Given ID1, get the title ✓
    --Open the file labelled ID1 and the file labelled ID2 underneath it ✓
    --Rename all files in Folder_ID2 to Title + ID2 ✓
    --Copy the renamed file somewhere else? A new directory ✓

Some checks I need to make:
    --What if the file name doesn't match an ID? ✓
    --Check all file IDs are unique ✓
    --That the file is a PDF (Folder Kind = Adobe Acrobat) ✓
    --Flag duplicates ✓
In each case there is some error, flag the error (in a .txt file?)
    --Create a textfile that documents what the python script is doing
    --Error handling ✓

To Do List:
    --Write the ID2s into the spreadsheet with ID1s and filenames ✓
        >In order to do this I need to move from xlrd to openpyxl to enable reading AND writing ✓
    --Allow user input for ID_col, Title_col and kind_col ✓
    --Progress bar ✓
    --Better handling of duplicate files ✓
