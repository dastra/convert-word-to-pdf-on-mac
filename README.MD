# Automating conversion of files from word to PDF on macOS

This uses:
* AppleScript to open microsoft word and save as a PDF 
* An Automator Application to wrap the AppleScript
* Find from the Terminal to find word documents to call the Automator Applicator

And it works on Catalina and with Office 365.

## The AppleScript script

The purpose of the script is, for the filename passed in, to open the file in word, and then save the file with a .pdf extension instead of .doc or .docx.  

~~~~ AppleScript
on run {input, parameters}
	repeat with aFile in input
		repeat 1 times -- # fake loop

			-- Create the output filename
			tell application "System Events"
				set inputFile to disk item (aFile as text)
				
				tell (info for input) to set Nm to name
				if (text -5 thru -1 of Nm) is ".docx" then
					set outputFileName to ((text 1 thru -6 of Nm) & ".pdf")
				else if (text -4 thru -1 of Nm) is ".doc" then
					-- .doc
					set outputFileName to ((text 1 thru -5 of Nm) & ".pdf")
				else
					-- If this isn't a word document, continue to the next file.
					exit repeat
				end if
			end tell
			
			-- Open word, save as PDF in default path
			tell application id "com.microsoft.Word"
				activate
				open aFile
				tell active document
					save as it file name outputFileName file format format PDF
					close saving no
				end tell
				set defaultPath to get default file path file path type documents path
			end tell
			
			-- Move to output location
			tell application "System Events"
				set outputPath to (container of inputFile)
				set outputFile to disk item outputFileName of folder defaultPath
				move outputFile to outputPath
			end tell
		end repeat
	end repeat
	return input
end run
~~~~



## The Automator Application

1. Open automator
2. File / New then select Application
3. From the actions Library in the left hand nav, select "Run AppleScript" and drag it into the window on the righ
4. Copy and paste the code above into the Run AppleScript window.
5. Save the application as "WordToPdf.app" on the Desktop

## Calling it from the command line

1. Open Terminal
2. Change Directory to the one you wish to run the script in.  Note that it is recursive and will process all files in all subdirectories
3. Run the following code:
    ~~~~shell script
     find . -type f \( -iname \*.doc -o -iname \*.docx \) -print0 | while read -d $'\0' file
     do
         automator -i "$file" ~/Desktop/WordToPdf.app
     done
    ~~~~
4. Double check that you are happy, and then optionally remove all of the word documents:
    ~~~~ shell script
    find . -type f \( -iname \*.doc -o -iname \*.docx \) -exec rm {} \;
    ~~~~

