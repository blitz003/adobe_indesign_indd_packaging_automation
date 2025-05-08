import glob, tkinter as tk
from tkinter import filedialog, ttk
import os
import shutil
import time
import sys
import platform
import subprocess
from pathlib import Path
from typing import Dict, List, Union, Any

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

    # AppleScript to close Finder

    def close_finder(self) -> bool:
        """
        Attempts to quit the Finder.
        """
        script = '''
            tell application "Finder"
                quit
            end tell
            '''

        result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True)

        if result.returncode != 0:
            print("✗ Failed to close Finder:",
                  result.stderr.strip() or "(no message)")
            return False
        return True

    def open_extensis_connect(self):
        """
        Opens Extensis Connect application.

        Returns:
        - True if successful, False otherwise
        """
        try:
            open_command = 'tell application "Extensis Connect" to activate'
            result = subprocess.run(['osascript', '-e', open_command],
                                    capture_output=True,
                                    text=True)

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
            refresh_command = 'tell application "System Events" to tell process "Extensis Connect" to keystroke "r" using command down'
            result = subprocess.run(['osascript', '-e', refresh_command],
                                    capture_output=True,
                                    text=True)

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

    def select_indesign_file(self):
        """
        Prompts the user to select an InDesign file.

        Returns:
        - str: Path to the selected InDesign file, or None if no file is selected
        """
        root = tk.Tk()
        root.withdraw()  # Hide the root window

        file_path = filedialog.askopenfilename(
            title="Select an InDesign file",
            filetypes=[("InDesign Files", "*.indd")])

        if file_path:
            print(f"Selected file: {file_path}")
            return file_path
        else:
            print("No file selected.")
            return None

    def open_indesign_file(self, file_path):
        """
        Opens an InDesign file with Adobe InDesign.

        Args:
        - file_path (str): Path to the InDesign file
        """
        try:
            applescript_cmd = f'''
            tell application "Adobe InDesign 2025"
                activate
                open "{file_path}"
            end tell
            '''

            result = subprocess.run(['osascript', '-e', applescript_cmd],
                                    capture_output=True,
                                    text=True)

            if result.returncode == 0:
                print(f"Successfully opened {os.path.basename(file_path)} with InDesign.")
                self.press_skip_on_missing_fonts_dialog()
            else:
                print(f"Error opening file with InDesign: {result.stderr}")
        except Exception as e:
            print(f"An unexpected error occurred while opening {os.path.basename(file_path)}: {e}")

    def press_skip_on_missing_fonts_dialog(self):
        """
        Presses the "Esc" key twice to close any dialog in Adobe InDesign.
        No dialog detection - simply presses Escape keys after a delay.

        Returns:
            - True if both button presses were executed without errors, False otherwise
        """
        try:
            # Initial delay to give InDesign time to fully launch or show any dialogs
            print("Executing Escape key sequence...")

            # Focus on InDesign application first
            applescript_focus_cmd = '''
            tell application "Adobe InDesign 2025"
                activate
            end tell
            '''
            subprocess.run(['osascript', '-e', applescript_focus_cmd], check=True)

            # Short pause after focusing
            time.sleep(3)

            # Define AppleScript to press Escape key
            applescript_escape_cmd = '''
            tell application "System Events"
                key code 53  # Direct key code for Escape key
            end tell
            '''

            # First Escape press
            subprocess.run(['osascript', '-e', applescript_escape_cmd], check=True)
            print("First Escape key sent")

            # Brief pause between key presses
            time.sleep(1)

            # Second Escape press
            subprocess.run(['osascript', '-e', applescript_escape_cmd], check=True)
            print("Second Escape key sent")

            return True

        except Exception as e:
            print(f"Error executing Escape key sequence: {e}")
            return False

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

    from pathlib import Path
    from typing import Optional, Dict

    def select_and_process_indesign_files(self):
        """
        Select an InDesign file, open it, and press the "Escape" key twice to close any dialogs.

        Returns:
        - dict: The results of the processing
        """
        result = {"results": []}

        # Step 1: Select an InDesign file
        selected_file = self.select_indesign_file()

        if selected_file:
            # Step 2: Open the selected InDesign file
            self.open_indesign_file(selected_file)

            # Step 3: Press the Escape key twice
            print("Handling potential dialogs...")
            if self.press_skip_on_missing_fonts_dialog():
                result["results"].append({"file": selected_file, "success": True})
                print("Dialogs handled successfully.")
            else:
                result["results"].append({"file": selected_file, "success": False})
                print("Error handling dialogs.")
        else:
            result["results"].append({"file": None, "success": False})
            print("No file selected.")

        return result

    def package_indesign_file(
            self,
            folder_id: str,
            project_name: Optional[str] = None
    ) -> Dict[str, str | bool]:
        """
        Package the *currently‑open* InDesign document into
        ~/Documents/Archived_Projects/<project_name>/<folder_id>_Layout.

        Args
        ----
        folder_id     unique numeric / text ID stamped on the Layout folder
        project_name  optional project folder; defaults to "Unnamed"

        Returns
        -------
        dict with 'success', 'message' and (on success) 'package_path'
        """
        try:
            # ------------------------------------------------------------------ #
            # 1. Build the destination folder on the Mac file‑system
            # ------------------------------------------------------------------ #
            root = Path.home() / "Documents" / "Archived_Projects"
            project_dir = root / (project_name or "Unnamed")
            layout_dir = project_dir / f"{folder_id}_Layout"
            layout_dir.mkdir(parents=True, exist_ok=True)

            # ------------------------------------------------------------------ #
            # 2. Build the AppleScript
            # ------------------------------------------------------------------ #
            # Escape any quotes in the POSIX path before embedding
            dest_root_posix = str(layout_dir).replace('"', r'\"')

            applescript = f'''
    use AppleScript version "2.7"
    use scripting additions

    -- destination folder provided by Python
    set destRootPOSIX to "{dest_root_posix}"
    set destRootAlias to POSIX file destRootPOSIX as alias

    tell application "Adobe InDesign 2025"
        activate
        if (count of documents) is 0 then
            error "No document is open in InDesign."
        end if

        -- suppress all UI
        set originalLevel to user interaction level of script preferences
        set user interaction level of script preferences to never interact

        try
            set myDoc to document 1
            set nm to name of myDoc -- e.g. "Brochure.indd"
            if nm ends with ".indd" then set nm to text 1 thru -6 of nm

            -- create a *_Packaged folder for this document
            set pkgPathPOSIX to destRootPOSIX & "/" & nm & "_Packaged"
            do shell script "mkdir -p " & quoted form of pkgPathPOSIX
            set pkgFolderAlias to POSIX file pkgPathPOSIX as alias

            -- package with long‑form option labels (per dictionary)
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
    '''

            # ------------------------------------------------------------------ #
            # 3. Run the AppleScript (feed via stdin instead of multiple -e flags)
            # ------------------------------------------------------------------ #
            result = subprocess.run(
                ["/usr/bin/osascript", "-"],  # "-" = read script from stdin
                input=applescript,
                text=True,
                capture_output=True,
                check=False
            )

            if result.returncode == 0:
                pkg_path = result.stdout.strip()
                return {
                    "success": True,
                    "message": f"Package created at {pkg_path}",
                    "package_path": pkg_path
                }
            else:
                return {
                    "success": False,
                    "error": result.stderr.strip()
                }

        except Exception as exc:
            return {
                "success": False,
                "error": str(exc)
            }

    def count_indesign_files(self):
        """
        Pops up a folder‑chooser, counts *.indd files inside, and
        returns (paths, integer_count).  If the user cancels, both
        values are empty/zero so callers can bail out gracefully.
        """
        root = tk.Tk();
        root.withdraw()
        folder = filedialog.askdirectory(
            title="Choose the folder that contains your InDesign files"
        )
        if not folder:  # user hit Cancel
            return [], 0

        paths = glob.glob(os.path.join(folder, "*.indd"))
        return paths, len(paths)

    def count_cover_indesign_files(self):
        """
        Pops up a file‑chooser for a single .indd file and
        returns (paths, integer_count). If the user cancels,
        both values are empty/zero so callers can bail out gracefully.
        """
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="Select your Full Cover InDesign file",
            filetypes=[("InDesign Files", "*.indd")],
        )
        if not file_path:  # user hit Cancel
            return [], 0

        return [file_path], 1

    # ──────────────────────────────────────────────────────────────
    # NEW: close docs and quit InDesign so the next run is clean
    # ──────────────────────────────────────────────────────────────
    def close_indesign(self):
        """
        Closes every open document (no save) and quits the app.
        Based on Adobe’s CloseDocument/CloseAll examples.
        """
        script = '''
        tell application "Adobe InDesign 2025"
            if (count documents) > 0 then
                tell documents to close saving no
            end if
            quit saving no
        end tell
        '''
        subprocess.run(["osascript", "-e", script],
                       capture_output=True, text=True)


class FileCheck:
    def __init__(self, name="Alpha"):
        self.name = name

    def verify_nonzero_file_sizes(self, root_dir: str) -> Dict[str, Any]:
        """
        Recursively check every file inside *root_dir* and all its subdirectories
        to ensure each file's size is larger than 0 bytes.

        Parameters
        ----------
        root_dir : str
            Path to the root directory to search

        Returns
        -------
        dict with:
            success        : True if every file passed
            empty_files    : list of file paths whose size == 0
            checked_count  : total number of regular files examined
        """
        empty_files: List[str] = []
        checked: int = 0

        # First, verify the root_dir exists and is actually a directory
        if not os.path.exists(root_dir):
            print(f"⚠️  Directory does not exist: {root_dir}")
            return {
                "success": False,
                "empty_files": [],
                "checked_count": 0,
                "error": f"Directory does not exist: {root_dir}"
            }

        if not os.path.isdir(root_dir):
            print(f"⚠️  Not a directory: {root_dir}")
            return {
                "success": False,
                "empty_files": [],
                "checked_count": 0,
                "error": f"Not a directory: {root_dir}"
            }

        # Proceed with recursive file checking (including all subdirectories)
        print(f"Starting recursive file size check on directory: {root_dir}")

        for dirpath, _dirs, filenames in os.walk(root_dir):
            # This will go through the root directory and all subdirectories
            current_dir = os.path.relpath(dirpath, root_dir)
            if current_dir == ".":
                current_dir = "root directory"
            else:
                print(f"Checking subdirectory: {current_dir}")

            for fname in filenames:
                fpath = os.path.join(dirpath, fname)
                try:
                    # Check if it's a file (not a symlink or other special file)
                    if os.path.isfile(fpath):
                        file_size = os.path.getsize(fpath)
                        if file_size == 0:
                            empty_files.append(fpath)
                            # Print the full absolute path of the empty file for easier identification
                            print(f"🚫 EMPTY FILE DETECTED: {os.path.abspath(fpath)} (0 KB)")
                        checked += 1
                except OSError as exc:
                    # Something went wrong reading the file size
                    print(f"⚠️  Could not stat {fpath}: {exc}")

        # Final summary and return results
        print(f"File check complete: examined {checked} files across all subdirectories")
        if empty_files:
            print(f"⚠️ WARNING: Found {len(empty_files)} empty (0 KB) files:")
            for idx, empty_file in enumerate(empty_files, 1):
                print(f"  {idx}. {os.path.abspath(empty_file)}")
        else:
            print("✅ No empty files found - all files have content!")

        return {
            "success": len(empty_files) == 0,
            "empty_files": empty_files,
            "checked_count": checked,
            "directories_checked": sum(1 for _ in os.walk(root_dir))  # Count of directories processed
        }


