# terminal code to package the file
terminal output = """

cd ~/Documents/Executables
python3 -m PyInstaller --clean --onefile \
    --name "Archival Automation" \
    --add-data "cnt:cnt" \
    run_cnt.py


"""


# applescript = CONST_PACKAGE_INDESIGN_FILE_APPLESCRIPT
CONST_PACKAGE_INDESIGN_FILE_APPLESCRIPT = f'''
use AppleScript version "2.7"
use scripting additions

set destRootPOSIX to "{dest_root_posix}"
set destRootAlias to POSIX file destRootPOSIX as alias

-------------------------------------------------------------------------------
--  Ensure Extensis Connect has all fonts active **before** packaging
-------------------------------------------------------------------------------
tell application "Extensis Connect" to activate
tell application "System Events" to tell process "Extensis Connect" ¬
    to keystroke "r" using command down
delay 10 -- wait while Connect re-syncs and enables fonts
-------------------------------------------------------------------------------

-- ⏱  DISABLE THE TIMEOUT FOR THE WHOLE INDESIGN SESSION
with timeout of 1200 seconds
    tell application id "com.adobe.InDesign"
        activate
        if (count documents) is 0 then error "No document is open in InDesign."

        -- suppress all UI
        set originalLevel to user interaction level of script preferences
        set user interaction level of script preferences to never interact

        try
            set myDoc to document 1
            set nm to name of myDoc
            if nm ends with ".indd" then set nm to text 1 thru -6 of nm

            -- create a *_Packaged folder for this document
            set pkgPathPOSIX to destRootPOSIX & "/" & nm & "_Packaged"
            do shell script "mkdir -p " & quoted form of pkgPathPOSIX
            set pkgFolderAlias to POSIX file pkgPathPOSIX as alias

            -- package with long-form option labels (per dictionary)
            tell myDoc to package ¬
                to pkgFolderAlias ¬
                copying fonts yes ¬
                copying linked graphics yes ¬
                copying profiles yes ¬
                updating graphics yes ¬
                including hidden layers yes ¬
                ignore preflight errors yes ¬
                include idml no ¬
                include pdf no ¬
                creating report yes

            set user interaction level of script preferences to originalLevel
            return pkgPathPOSIX
        on error errMsg number errNum
            set user interaction level of script preferences to originalLevel
            error errMsg number errNum
        end try
    end tell
end timeout
'''