-- Font Audit Droplet
-- Drag a .pptx file onto this app, or double-click to pick one.

on findPython()
	set homePath to POSIX path of (path to home folder)
	set pythonCandidates to {homePath & "font-audit-env/bin/python3", "/opt/homebrew/bin/python3", "/usr/local/bin/python3", "/usr/bin/python3"}
	repeat with p in pythonCandidates
		try
			do shell script "test -x " & quoted form of p
			return p as text
		end try
	end repeat
	return missing value
end findPython

on getAuditScript()
	set homePath to POSIX path of (path to home folder)
	return homePath & "scripts/font_audit.py"
end getAuditScript

on runAudit(filePath)
	set py to findPython()
	if py is missing value then
		display dialog "Python 3 not found." & return & return & "Install it with:" & return & "  brew install python3" & return & return & "Or download from python.org" with title "Font Audit" buttons {"OK"} default button "OK" with icon stop
		return
	end if
	
	try
		do shell script quoted form of py & " -c 'import pptx'"
	on error
		display dialog "The python-pptx library is not installed." & return & return & "Open Terminal and run:" & return & "  pip3 install python-pptx" & return & return & "Or if using a venv:" & return & "  source ~/font-audit-env/bin/activate" & return & "  pip install python-pptx" with title "Font Audit" buttons {"OK"} default button "OK" with icon stop
		return
	end try
	
	if filePath does not end with ".pptx" then
		display dialog "Please drop a .pptx file onto the app." with title "Font Audit" buttons {"OK"} default button "OK" with icon stop
		return
	end if
	
	set baseName to do shell script "basename " & quoted form of filePath & " .pptx"
	set defaultName to baseName & "_font_audit.csv"
	
	try
		set savePath to POSIX path of (choose file name with prompt "Save font audit CSV as:" default name defaultName default location (path to desktop))
	on error
		return
	end try
	
	if savePath does not end with ".csv" then
		set savePath to savePath & ".csv"
	end if
	
	set auditScript to getAuditScript()
	
	try
		set auditOutput to do shell script quoted form of py & " " & quoted form of auditScript & " " & quoted form of filePath & " --csv " & quoted form of savePath & " 2>&1"
		
		set fontCount to do shell script "echo " & quoted form of auditOutput & " | grep 'Unique fonts:' | awk '{print $NF}'"
		set runCount to do shell script "echo " & quoted form of auditOutput & " | grep 'Total text runs:' | awk '{print $NF}'"
		
		set userChoice to button returned of (display dialog "Font audit complete!" & return & return & fontCount & " unique fonts found across " & runCount & " text runs." & return & return & "CSV saved to:" & return & savePath with title "Font Audit" buttons {"Open CSV", "Done"} default button "Done")
		
		if userChoice is "Open CSV" then
			do shell script "open " & quoted form of savePath
		end if
		
	on error errMsg
		display dialog "Error running font audit:" & return & return & errMsg with title "Font Audit" buttons {"OK"} default button "OK" with icon stop
	end try
end runAudit

on open theFiles
	repeat with theFile in theFiles
		set filePath to POSIX path of theFile
		runAudit(filePath)
	end repeat
end open

on run
	try
		set theFile to choose file with prompt "Select a PowerPoint file:" of type {"org.openxmlformats.presentationml.presentation"}
		set filePath to POSIX path of theFile
		runAudit(filePath)
	end try
end run
