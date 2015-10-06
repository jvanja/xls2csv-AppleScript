set workingFolder to (choose folder)
set workingFolderAsText to (workingFolder as text)
set theDoc to workingFolderAsText & "test.xls"
set outPath to workingFolderAsText & "test.csv"

tell application "Finder" to set myFiles to every file of workingFolder
repeat with aFile in myFiles
	set filename to name of aFile
	if not (filename is ".DS_Store") then
		set fileExtension to text ((length of filename) - 3) thru (length of filename) of filename
		if (fileExtension is ".xls") then
			tell application "Microsoft Excel"
				set theDoc to workingFolderAsText & filename
				set outPath to text 1 thru ((length of theDoc) - 3) of theDoc & "csv"
				open file theDoc
				tell workbook 1
					tell sheet 1
						save in outPath as CSV file format
					end tell
					close without saving
				end tell
			end tell
		end if
	end if
end repeat

display dialog "All done!!!"
