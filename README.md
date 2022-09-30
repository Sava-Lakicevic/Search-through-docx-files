# Search-through-docx-files
The point of this Python script is to search for a word within all files with the extension .docx within the current directory and all subdirectories (current folder and all subfolders). The scope of the search depends on where it is executed. I recommend to run it through the command prompt, change in to the directory where you need to run it from (unless you want to search through the entire computer) and launch the Python file from there. You don't need to have the Python file in the same folder, just know its path. The location of the file is irrelevant to the execution of the script.

The output is the parent folder, the name of the file and paragraph number (starting from 1). It take the input through the command prompt with the message "What are you looking for: ", and checks for the string in every docx file.

Output example: File: ['folder_name', 'file_name.docx']; paragraph number 3

It will run until a keyboard iterrupt (usually ctrl+C) or if you enter "end of work" into the input.
