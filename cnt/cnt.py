import tkinter as tk
from tkinter import filedialog, ttk
import os
import shutil
import subprocess
import time
import sys
import platform


class TKFolderSelector:
    def __init__(self):
        """
        Initialize the folder selector with no initial folder name.
        """
        self.folder_name = None
        self.folder_path = None

    def select_folder(self):
        """
        Prompts the user with a tkinter folder selection dialog,
        and stores the selected folder name in `self.folder_name`.

        Returns:
            str or None: The selected folder name, or None if no folder was selected
        """
        # Create the root window
        root = tk.Tk()

        # Make the window always on top and bring it to focus
        root.attributes('-topmost', True)

        # Set a reasonable initial size and center the window
        root.geometry('500x500')
        root.eval('tk::PlaceWindow . center')

        # Remove the main window's decorations
        root.overrideredirect(True)

        # Let user pick a directory
        folder_selected = filedialog.askdirectory(
            title="Select Your Project Directory",
            parent=root  # Use the root as the parent window
        )

        # Destroy the window
        root.destroy()

        if folder_selected:
            # Store full path
            self.folder_path = folder_selected

            # Extract only the directory name
            self.folder_name = os.path.basename(folder_selected)

            print(f"Selected folder name: {self.folder_name}")
            return self.folder_name
        else:
            print("No folder selected.")
            return None


    def copy_specific_subdirectories(self, destination_path, folder_id):
        """
        Copy specific subdirectories from self.folder_path to destination_path.

        Args:
            destination_path (str): Path where subdirectories will be copied
            folder_id (str): Folder ID to replace in subdirectory names

        Returns:
            dict: Tracking of copied directories
        """
        # Dynamic list of subdirectories using f-strings
        target_subdirs = [
            f'{folder_id}_Digital_Content',
            f'{folder_id}_Logs',
            f'{folder_id}_Manuscript',
            f'{folder_id}_Office'
        ]

        # Ensure destination path exists
        os.makedirs(destination_path, exist_ok=True)

        # Dictionary to track copy status
        copy_status = {}

        # Iterate through target subdirectories
        for subdir in target_subdirs:
            # Construct full source path
            source_subdir_path = os.path.join(self.folder_path, subdir)

            # Construct full destination path
            dest_subdir_path = os.path.join(destination_path, subdir)

            try:
                # Check if source subdirectory exists
                if os.path.exists(source_subdir_path):
                    # Copy the entire directory
                    shutil.copytree(source_subdir_path, dest_subdir_path)
                    copy_status[subdir] = "Copied successfully"
                    print(f"Copied {subdir} to {dest_subdir_path}")
                else:
                    copy_status[subdir] = "Source directory not found"
                    print(f"Warning: {subdir} not found in source directory")

            except Exception as e:
                copy_status[subdir] = f"Error during copy: {str(e)}"
                print(f"Error copying {subdir}: {e}")

        return copy_status

    def create_project_subdirectories(self, archived_project_path, folder_id):
        """
        Create specific subdirectories for the project.

        Args:
            folder_id (str): Folder ID to use in subdirectory names

        Returns:
            dict: Status of subdirectory creation
        """
        # List of subdirectories to create
        subdirs = [
            f'{folder_id}_Printer_PDFs',
            f'{folder_id}_Layout'
        ]

        # Dictionary to track creation status
        creation_status = {}

        # Ensure the parent directory exists
        os.makedirs(archived_project_path, exist_ok=True)

        # Create each subdirectory
        for subdir in subdirs:
            # Full path for the subdirectory
            subdir_path = os.path.join(archived_project_path, subdir)

            try:
                # Create the subdirectory
                os.makedirs(subdir_path, exist_ok=True)
                creation_status[subdir] = "Created successfully"
                print(f"Created directory: {subdir_path}")

            except Exception as e:
                creation_status[subdir] = f"Error during creation: {str(e)}"
                print(f"Error creating {subdir}: {e}")

        return creation_status

    def copy_print_files(self, project_layout_path, archive_printer_pdfs_path, folder_id_print):
        """
        Copy print files from project layout path to archive printer PDFs path.

        Args:
        project_layout_path (str): Path to the layout folder
        archive_printer_pdfs_path (str): Destination path for archived print PDFs
        folder_id_print (str): Prefix to identify print files

        Returns:
        dict: Summary of copy operation
        """
        # Ensure the destination directory exists
        os.makedirs(archive_printer_pdfs_path, exist_ok=True)

        # Initialize tracking variables
        copied_files = []
        skipped_files = []

        # Check if source directory exists
        if not os.path.exists(project_layout_path):
            return {
                'success': False,
                'message': f'Source directory does not exist: {project_layout_path}',
                'copied_files': [],
                'skipped_files': []
            }

        # Iterate through files in the layout path
        for filename in os.listdir(project_layout_path):
            # Check if filename starts with folder_id_print
            if filename.startswith(folder_id_print):
                source_file = os.path.join(project_layout_path, filename)
                destination_file = os.path.join(archive_printer_pdfs_path, filename)

                try:
                    # Copy the file (preserving metadata)
                    shutil.copy2(source_file, destination_file)
                    copied_files.append(filename)
                except Exception as e:
                    skipped_files.append((filename, str(e)))

        # Prepare return dictionary
        return {
            'success': len(skipped_files) == 0,
            'message': f'Copied {len(copied_files)} files. Skipped {len(skipped_files)} files.',
            'copied_files': copied_files,
            'skipped_files': skipped_files
        }

class MakeDirectory:
    def __init__(self, file_path=None):
        self.file_path = file_path

    def create_archived_projects_directory(self):
        """
        Create a new directory called 'Archived_Projects' in the user's Documents folder.

        Returns:
            str: Path to the Archived_Projects directory
            None: If directory cannot be created
        """
        # Get the path to the Documents directory
        documents_path = os.path.expanduser("~/Documents")

        # Create the full path for the new directory
        archived_projects_path = os.path.join(documents_path, "Archived_Projects")

        # Check if directory already exists
        if os.path.exists(archived_projects_path):
            return archived_projects_path

        try:
            # Create the directory
            os.makedirs(archived_projects_path)
            return archived_projects_path

        except PermissionError:
            print("Permission denied. Unable to create directory.")
            return None
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def create_project_directory(self, project_name):
        """
        Create a new directory inside the Archived_Projects directory.

        Args:
            project_name (str): Name of the project directory to create

        Returns:
            str or None: Path to the created project directory, or None if creation failed
        """
        # Get the Archived_Projects directory path
        archived_projects_path = os.path.expanduser("~/Documents/Archived_Projects")

        # If Archived_Projects directory creation failed, return None
        if not archived_projects_path:
            return None

        # Create full path for the new project directory
        project_path = os.path.join(archived_projects_path, project_name)

        try:
            # Create the project directory
            os.makedirs(project_path, exist_ok=True)
            print(f"Project directory created successfully at: {project_path}")
            return project_path

        except PermissionError:
            print(f"Permission denied. Unable to create project directory: {project_name}")
            return None
        except Exception as e:
            print(f"An error occurred while creating project directory: {e}")
            return None


class AppleScript:
    def __init__(self, name="Alpha"):
        self.name = name

    def open_extensis_connect(self):
        """
        Opens Extensis Connect application.

        Returns:
        - True if successful, False otherwise
        """
        try:
            # Open Extensis Connect
            open_command = 'tell application "Extensis Connect" to activate'
            result = subprocess.run(['osascript', '-e', open_command],
                                    capture_output=True,
                                    text=True)

            # Check if the open command was successful
            if result.returncode == 0:
                print("Extensis Connect has been opened successfully.")
                return True
            else:
                print(f"Error opening Extensis Connect: {result.stderr}")
                return False

        except Exception as e:
            print(f"An unexpected error occurred while opening Extensis Connect: {e}")
            return False

    def refresh_extensis_connect(self):
        """
        Refreshes Extensis Connect using Command+R.

        Returns:
        - True if successful, False otherwise
        """
        try:
            # Refresh command using Command+R
            refresh_command = 'tell application "System Events" to tell process "Extensis Connect" to keystroke "r" using command down'
            result = subprocess.run(['osascript', '-e', refresh_command],
                                    capture_output=True,
                                    text=True)

            # Check if the refresh command was successful
            if result.returncode == 0:
                time.sleep(4)
                print("Extensis Connect has been refreshed successfully.")
                return True
            else:
                print(f"Error refreshing Extensis Connect: {result.stderr}")
                return False

        except Exception as e:
            print(f"An unexpected error occurred while refreshing Extensis Connect: {e}")
            return False

    def open_and_refresh_extensis_connect(self, load_time=10):
        """
        Opens Extensis Connect and refreshes it after a specified loading time.

        Args:
        load_time (int): Number of seconds to wait before refreshing. Default is 10 seconds.

        Returns:
        - True if both operations were successful, False otherwise
        """
        # Open the application
        open_success = self.open_extensis_connect()

        if not open_success:
            return False

        # Wait for the application to load
        print(f"Waiting {load_time} seconds for Extensis Connect to load...")
        time.sleep(load_time)

        # Refresh the application
        return self.refresh_extensis_connect()

    def minimize_extensis_connect(self):
        """
        Minimizes the Extensis Connect application window.

        Returns:
        - True if successful, False otherwise
        """
        try:
            # AppleScript command to minimize the Extensis Connect window
            minimize_command = 'tell application "System Events" to tell process "Extensis Connect" to set visible to false'
            result = subprocess.run(['osascript', '-e', minimize_command],
                                    capture_output=True,
                                    text=True)

            # Check if the minimize command was successful
            if result.returncode == 0:
                print("Extensis Connect has been minimized.")
                return True
            else:
                print(f"Error minimizing Extensis Connect: {result.stderr}")
                return False

        except Exception as e:
            print(f"An unexpected error occurred while minimizing Extensis Connect: {e}")
            return False

    def open_and_refresh_extensis_connect(self, load_time=10, minimize_after=True):
        """
        Opens Extensis Connect, refreshes it after a specified loading time,
        and optionally minimizes it afterward.

        Args:
        load_time (int): Number of seconds to wait before refreshing. Default is 10 seconds.
        minimize_after (bool): Whether to minimize the app after refreshing. Default is True.

        Returns:
        - True if all operations were successful, False otherwise
        """
        # Open the application
        open_success = self.open_extensis_connect()

        if not open_success:
            return False

        # Wait for the application to load
        print(f"Waiting {load_time} seconds for Extensis Connect to load...")
        time.sleep(load_time)

        # Refresh the application
        refresh_success = self.refresh_extensis_connect()

        if not refresh_success:
            return False

        # Minimize if requested
        if minimize_after:
            return self.minimize_extensis_connect()

        return True

    def is_extensis_connect_running(self):
        """
        Checks if Extensis Connect is currently running.

        Returns:
        - True if Extensis Connect is running
        - False if it's not running or an error occurred
        """
        try:
            # Use the 'ps' command to check for the Extensis Connect process
            process_path = "/Applications/Extensis Connect.app/Contents/MacOS/Extensis Connect"

            # Create the command to run
            command = f"ps -ef | grep '{process_path}' | grep -v grep"

            # Run the command and get the output
            result = subprocess.run(command, shell=True, capture_output=True, text=True)

            # If output contains anything, the process is running
            is_running = bool(result.stdout.strip())

            if is_running:
                print("Extensis Connect is currently running.")
            else:
                print("Extensis Connect is not running.")

            return is_running

        except Exception as e:
            print(f"Error checking if Extensis Connect is running: {e}")
            return False

    # 4/3/2025

    def open_file_with_indesign(self, file_path):
        """
        Opens a specific file with Adobe InDesign 2025.

        Args:
        file_path (str): Path to the file to open

        Returns:
        - True if successful, False otherwise
        """
        try:
            # Ensure file exists
            if not os.path.isfile(file_path):
                print(f"File does not exist: {file_path}")
                return False

            print(f"Opening file: {file_path}")

            # Prepare AppleScript command to open file with InDesign
            applescript_cmd = f'''
            tell application "Adobe InDesign 2025"
                activate
                open "{file_path}"
            end tell
            '''

            # Execute AppleScript
            result = subprocess.run(['osascript', '-e', applescript_cmd],
                                    capture_output=True,
                                    text=True)

            # Check if command was successful
            if result.returncode == 0:
                print(f"Successfully opened {os.path.basename(file_path)} with InDesign.")

                # Check for missing links dialog
                has_missing_links = self.check_for_missing_links_dialog()
                if has_missing_links:
                    print("Missing links detected. Exiting process.")
                    return False

                return True
            else:
                print(f"Error opening file with InDesign: {result.stderr}")
                return False

        except Exception as e:
            print(f"An unexpected error occurred while opening {os.path.basename(file_path)}: {e}")
            return False

    def check_for_missing_links_dialog(self, wait_time=8):
        """
        Checks if Adobe InDesign 2025 has displayed a dialog about missing links.

        Args:
        wait_time (int): Time to wait for dialog to appear in seconds

        Returns:
        - True if missing links dialog was detected, False otherwise
        """
        try:
            # Wait briefly for any dialog to appear
            time.sleep(wait_time)

            # AppleScript to check for dialog with missing links message
            applescript_cmd = '''
            tell application "System Events"
                tell process "Adobe InDesign 2025"
                    set dialogExists to false

                    -- Check if any dialog exists
                    if exists (window 1 whose role is "AXDialog") then
                        set dialogWindow to window 1 whose role is "AXDialog"

                        -- Check static text elements for the target message
                        repeat with textElement in static texts of dialogWindow
                            set textContent to value of textElement

                            -- Look for text about missing links
                            if textContent contains "missing links" or textContent contains "missing" or textContent contains "Links panel" then
                                set dialogExists to true
                                exit repeat
                            end if
                        end repeat
                    end if

                    return dialogExists
                end tell
            end tell
            '''

            # Execute AppleScript
            result = subprocess.run(['osascript', '-e', applescript_cmd],
                                    capture_output=True,
                                    text=True)

            # Parse result - AppleScript returns "true" or "false" as strings
            is_dialog_present = result.stdout.strip().lower() == "true"

            if is_dialog_present:
                print("Missing links dialog detected.")

                # Optional: Click "OK" or close the dialog
                self.close_missing_links_dialog()

            return is_dialog_present

        except Exception as e:
            print(f"Error checking for missing links dialog: {e}")
            return False

    def close_missing_links_dialog(self):
        """
        Closes the missing links dialog if it's open.
        """
        try:
            # AppleScript to click OK button on the dialog
            applescript_cmd = '''
            tell application "System Events"
                tell process "Adobe InDesign 2025"
                    if exists (window 1 whose role is "AXDialog") then
                        -- Try to click OK or Cancel button
                        try
                            click button "OK" of window 1 whose role is "AXDialog"
                        on error
                            try
                                click button "Cancel" of window 1 whose role is "AXDialog"
                            end try
                        end try
                    end if
                end tell
            end tell
            '''

            subprocess.run(['osascript', '-e', applescript_cmd],
                           capture_output=True,
                           text=True)

        except Exception as e:
            print(f"Error closing missing links dialog: {e}")

    def select_and_process_indesign_files(self):
        """
        Prompts user to select a folder containing InDesign files and processes them.
        Detects and logs files with missing links.

        Returns:
        - dict: Summary of processing results
        """
        try:
            # Create the base AppleScript to select a folder and get all .indd files
            folder_selection_script = '''
            tell application "Finder"
                set selectedFolder to choose folder with prompt "Select a folder containing InDesign files"
                set inddFiles to files of selectedFolder whose name extension is "indd"
                set filePaths to {}

                repeat with eachFile in inddFiles
                    set end of filePaths to POSIX path of (eachFile as alias)
                end repeat

                return filePaths
            end tell
            '''

            # Execute AppleScript to get folder and files
            result = subprocess.run(['osascript', '-e', folder_selection_script],
                                    capture_output=True,
                                    text=True)

            if result.returncode != 0:
                print(f"Error selecting folder: {result.stderr}")
                return {"success": False, "message": "Error selecting folder", "processed": 0, "failed": 0}

            # Parse the result - AppleScript returns file paths as comma-separated values
            file_paths_output = result.stdout.strip()

            # If empty output, no files were selected
            if not file_paths_output:
                print("No InDesign files found in the selected folder.")
                return {"success": False, "message": "No InDesign files found", "processed": 0, "failed": 0}

            # Split the paths - AppleScript returns them comma-separated
            file_paths = [path.strip() for path in file_paths_output.split(',')]

            # Setup UI code (unchanged)...
            root = tk.Tk()
            root.withdraw()
            status_window = tk.Toplevel(root)
            status_window.title("Processing InDesign Files")
            status_window.attributes('-topmost', True)
            # Window size and position
            window_width = 500
            window_height = 250
            screen_width = status_window.winfo_screenwidth()
            screen_height = status_window.winfo_screenheight()
            center_x = int((screen_width - window_width) / 2)
            center_y = int((screen_height - window_height) / 2)
            status_window.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

            # Add status label and progress bar
            status_label = tk.Label(status_window, text="Starting processing...")
            status_label.pack(pady=20)
            file_label = tk.Label(status_window, text="")
            file_label.pack(pady=10)
            progress = ttk.Progressbar(status_window, length=400, mode='determinate')
            progress.pack(pady=20)
            progress['maximum'] = len(file_paths)

            # Initialize counters and results
            successful = 0
            failed = 0
            results = []
            missing_links_log = []

            # Create log directory and file
            log_dir = os.path.expanduser("~/Documents/InDesign_Processing_Logs")
            os.makedirs(log_dir, exist_ok=True)
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            log_file_path = os.path.join(log_dir, f"indesign_processing_log_{timestamp}.txt")

            with open(log_file_path, "w") as log_file:
                log_file.write(f"InDesign Processing Log - {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                log_file.write("=" * 50 + "\n\n")

                # Process each file individually with improved script
                for i, file_path in enumerate(file_paths):
                    file_name = os.path.basename(file_path)

                    # Update status
                    status_label.config(text=f"Processing file {i + 1} of {len(file_paths)}")
                    file_label.config(text=file_name)
                    progress['value'] = i
                    status_window.update()

                    # Log file being processed
                    log_file.write(f"Processing: {file_path}\n")

                    # IMPROVED: More robust AppleScript with better dialog detection and timeout
                    process_script = f'''
                    tell application "Adobe InDesign 2025"
                        set fileProcessed to false
                        set hasMissingLinks to false
                        set missingLinksCount to 0
                        set dialogMessage to ""

                        -- Make sure InDesign is frontmost
                        activate

                        -- Set a timeout for script operations
                        with timeout of 60 seconds
                            try
                                -- Open the file
                                open "{file_path}"
                                set fileProcessed to true

                                -- Function to check for dialog
                                on checkForDialog()
                                    tell application "System Events"
                                        if exists (process "Adobe InDesign 2025") then
                                            if exists (window 1 of process "Adobe InDesign 2025" whose subrole is "AXDialog") then
                                                return true
                                            else if exists (window 1 of process "Adobe InDesign 2025" whose role is "AXDialog") then
                                                return true
                                            end if
                                        end if
                                        return false
                                    end tell
                                end checkForDialog

                                -- Variable to track progress
                                set dialogFound to false
                                set dialogChecks to 0
                                set maxChecks to 10

                                -- Loop to check for dialog appearing (incremental checks)
                                repeat until dialogFound or dialogChecks ≥ maxChecks
                                    delay 0.5
                                    set dialogChecks to dialogChecks + 1
                                    set dialogFound to my checkForDialog()
                                end repeat

                                -- Process dialog if found
                                if dialogFound then
                                    tell application "System Events"
                                        tell process "Adobe InDesign 2025"
                                            set dialogWin to window 1 whose subrole is "AXDialog" or role is "AXDialog"

                                            -- Get all static text elements
                                            set allTexts to {{}}
                                            set dialogText to ""

                                            -- Try different approaches to get dialog text
                                            try
                                                set allTexts to get value of every static text of dialogWin
                                            on error
                                                try
                                                    set allTexts to get name of every static text of dialogWin
                                                end try
                                            end try

                                            -- Build complete dialog text
                                            repeat with t in allTexts
                                                set dialogText to dialogText & t & " "
                                            end repeat

                                            -- Check if it's a missing link dialog various ways
                                            if dialogText contains "missing link" or dialogText contains "Missing Link" or dialogText contains "links are missing" then
                                                set hasMissingLinks to true
                                                set dialogMessage to dialogText

                                                -- Extract the number through different patterns
                                                if dialogText contains "contains" then
                                                    set AppleScript's text item delimiters to " "
                                                    set textItems to every text item of dialogText
                                                    repeat with i from 1 to (count of textItems)
                                                        set thisItem to item i of textItems
                                                        if thisItem is "contains" and i < (count of textItems) then
                                                            try
                                                                set nextItem to item (i + 1) of textItems
                                                                set missingLinksCount to nextItem as integer
                                                            end try
                                                        end if
                                                    end repeat
                                                else if dialogText contains "links are missing" then
                                                    -- Different pattern for multiple links
                                                    set AppleScript's text item delimiters to " "
                                                    set textItems to every text item of dialogText
                                                    repeat with i from 1 to (count of textItems)
                                                        set thisItem to item i of textItems
                                                        if i < (count of textItems) and item (i + 1) of textItems is "links" then
                                                            try
                                                                set missingLinksCount to thisItem as integer
                                                            end try
                                                        end if
                                                    end repeat
                                                end if

                                                -- Click OK to dismiss the dialog - try different button labels
                                                try
                                                    if exists (button "OK" of dialogWin) then
                                                        click button "OK" of dialogWin
                                                    else if exists (button "Ok" of dialogWin) then
                                                        click button "Ok" of dialogWin
                                                    else
                                                        -- Last resort - try to find any button and click it
                                                        click button 1 of dialogWin
                                                    end if
                                                on error
                                                    -- If clicking fails, try to press return key
                                                    key code 36 -- Return key
                                                end try
                                            end if
                                        end tell
                                    end tell
                                end if

                                -- Close the document safely
                                tell application "Adobe InDesign 2025"
                                    if exists document 1 then
                                        set docRef to document 1
                                        close docRef saving no
                                    end if
                                end tell

                            on error errMsg
                                set fileProcessed to false

                                -- Try to close any open documents if there was an error
                                try
                                    tell application "Adobe InDesign 2025"
                                        if exists document 1 then
                                            close document 1 saving no
                                        end if
                                    end tell
                                end try

                                return "ERROR: " & errMsg
                            end try
                        end timeout

                        if hasMissingLinks then
                            return "MISSING_LINKS:" & missingLinksCount & ":" & dialogMessage
                        else if fileProcessed then
                            return "SUCCESS"
                        else
                            return "FAILED"
                        end if
                    end tell
                    '''

                    # Execute with timeout at Python level too
                    try:
                        result = subprocess.run(['osascript', '-e', process_script],
                                                capture_output=True,
                                                text=True,
                                                timeout=90)  # 90-second timeout

                        output = result.stdout.strip()

                        # Process the result - unchanged from your original code
                        if output.startswith("MISSING_LINKS:"):
                            parts = output.split(":", 2)
                            count = parts[1] if len(parts) > 1 else "unknown"
                            message = parts[2] if len(parts) > 2 else "Missing links detected"

                            log_message = f"WARNING: File '{file_name}' contains {count} missing links and needs review.\n"
                            log_message += f"Dialog message: {message}\n"
                            log_file.write(log_message + "\n")

                            missing_links_log.append({
                                "file": file_name,
                                "path": file_path,
                                "missing_links_count": count,
                                "message": message
                            })

                            failed += 1
                            results.append({"file": file_name, "success": False, "reason": "missing_links"})

                        elif output == "SUCCESS":
                            successful += 1
                            results.append({"file": file_name, "success": True})
                            log_file.write("Result: Successfully processed\n\n")

                        elif output.startswith("ERROR:"):
                            error_msg = output.replace("ERROR:", "").strip()
                            failed += 1
                            results.append(
                                {"file": file_name, "success": False, "reason": "error", "message": error_msg})
                            log_file.write(f"Result: FAILED - {error_msg}\n\n")

                        else:
                            failed += 1
                            results.append({"file": file_name, "success": False, "reason": "unknown"})
                            log_file.write("Result: FAILED - Missing links\n\n")

                    except subprocess.TimeoutExpired:
                        # Handle the case where the script times out
                        failed += 1
                        results.append({"file": file_name, "success": False, "reason": "timeout"})
                        log_file.write(f"Result: FAILED - Script timeout after 90 seconds\n\n")

                        # Force quit InDesign if it's stuck
                        try:
                            force_quit_script = '''
                            tell application "System Events"
                                if exists process "Adobe InDesign 2025" then
                                    do shell script "killall 'Adobe InDesign 2025'"
                                end if
                            end tell
                            '''
                            subprocess.run(['osascript', '-e', force_quit_script],
                                           capture_output=True,
                                           timeout=10)
                            # Give it a moment before continuing
                            time.sleep(2)
                        except:
                            pass

                # Write summary to log
                log_file.write("\n" + "=" * 50 + "\n")
                log_file.write(
                    f"SUMMARY: Processed {successful} files successfully. Failed to process {failed} files.\n")

                if missing_links_log:
                    log_file.write("\nFiles with missing links that need review:\n")
                    for entry in missing_links_log:
                        log_file.write(f"- {entry['file']}: {entry['missing_links_count']} missing links\n")

            # Create summary
            summary = {
                "success": failed == 0,
                "message": f"Processed {successful} files successfully. Failed to process {failed} files.",
                "processed": successful,
                "failed": failed,
                "results": results,
                "missing_links_files": missing_links_log,
                "log_file": log_file_path
            }

            # Show final status
            if missing_links_log:
                status_text = f"Processed {successful} files successfully.\n{len(missing_links_log)} files have missing links.\nSee log file for details."
            else:
                status_text = summary["message"]

            status_label.config(text=status_text)
            file_label.config(text=f"Log saved to: {log_file_path}")
            progress['value'] = len(file_paths)
            status_window.update()

            # Add a close button
            close_button = tk.Button(status_window, text="Close",
                                     command=lambda: [status_window.destroy(), root.destroy()])
            close_button.pack(pady=10)

            print(summary["message"])
            if missing_links_log:
                print(f"{len(missing_links_log)} files have missing links and need review.")
                print(f"See log file for details: {log_file_path}")

            # Wait for user to close the status window
            status_window.protocol("WM_DELETE_WINDOW", lambda: [status_window.destroy(), root.destroy()])
            root.wait_window(status_window)

            return summary

        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            if 'root' in locals():
                root.destroy()
            return {"success": False, "message": f"Error: {str(e)}", "processed": 0, "failed": 0}