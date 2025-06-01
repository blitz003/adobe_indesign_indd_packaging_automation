import sys
import os
import time
from cnt.cnt import TKFolderSelector, MakeDirectory, AppleScript, FileCheck


def main():
    # Create an instance of the folder selector
    folder_selector = TKFolderSelector()

    # Call the select_folder method
    selected_folder_name = folder_selector.select_folder()

    # Use the selected folder name if needed
    if selected_folder_name:
        # Perform operations with the folder name
        # Tokenize selected directory
        folder_tokens = selected_folder_name.split("_")
        folder_id = folder_tokens[0]  # 11492
        folder_semester = folder_tokens[1]  # S24
        folder_last_name = folder_tokens[2]  # Monroe
        folder_print_type = folder_tokens[3]  # Color

        output_directory_name = f"{folder_id}_{folder_last_name}"  # 11492_Monroe
    else:
        sys.exit()  # Close the program if no project directory is selected.

    # Initialize MakeDirectory instance
    directory_handler = MakeDirectory()

    # Step 1: Create Archived_Projects directory inside Documents directory
    archived_projects_path = directory_handler.create_archived_projects_directory()

    # Step 2: Create the project directory inside Archived_Projects
    directory_handler.create_project_directory(project_name=output_directory_name)

    # Step 3: Move specific subdirectories into the new Project Archive directory
    documents_path = os.path.expanduser("~/Documents")
    archived_project_path = os.path.join(documents_path, "Archived_Projects", output_directory_name)

    # Move subdirectories: Digital_Content, Logs, Manuscript, Office
    folder_selector.copy_specific_subdirectories(destination_path=archived_project_path, folder_id=folder_id)

    # Step 4: Create the following subdirectories in the new Project Archive directory
    folder_selector.create_project_subdirectories(archived_project_path=archived_project_path, folder_id=folder_id)

    # Step 5: Make AppleScript command to handle font software
    apple_script_agent = AppleScript(name=output_directory_name)

    # Step 5.5: Declare the full project directory path
    # Check file size > 0
    documents_root = os.path.expanduser("~/Documents")
    project_directory_path = os.path.join(documents_root, archived_project_path)
    print(archived_project_path)
    print(output_directory_name)
    # Step 6: Ensure Extensis Connect is running and refreshed
    if apple_script_agent.is_extensis_connect_running():
        apple_script_agent.refresh_extensis_connect()
    else:
        apple_script_agent.open_and_refresh_extensis_connect()
        apple_script_agent.minimize_extensis_connect()

    # Step 7: Move Print PDF files to /11492_Printer_PDFs from /11492_Layout
    printer_pdfs_endpoint = f"{folder_id}_Printer_PDFs"
    archive_printer_pdfs_path = os.path.join(archived_project_path, printer_pdfs_endpoint)
    folder_id_print = f"{folder_id}_Print"
    layout_endpoint = f"{folder_id}_Layout"
    project_layout_path = os.path.join(folder_selector.folder_path, layout_endpoint)

    # Copy print files from the layout folder to the Printer PDFs folder
    folder_selector.copy_print_files(
        project_layout_path=project_layout_path,
        archive_printer_pdfs_path=archive_printer_pdfs_path,
        folder_id_print=folder_id_print
    )

    apple_script_agent.close_finder()

    # STEP 1 – Get all .indd paths first
    paths, total = apple_script_agent.count_indesign_files()
    if total == 0:
        print("Nothing to process – exiting.")
        sys.exit(0)

    # STEP 2 – Iterate once per file
    for idx, path in enumerate(paths, start=1):
        print(f"[{idx}/{total}]  {os.path.basename(path)}")
        print(repr(path))
        apple_script_agent.open_indesign_file(path)


        pkg = apple_script_agent.package_indesign_file(
            folder_id=folder_id,
            project_name=archived_project_path
        )
        if pkg["success"]:
            print(f"{datetime.now():%Y-%m-%d %H:%M:%S} ✓ packaged → {pkg['message']}")
        else:
            print(f"{datetime.now():%Y-%m-%d %H:%M:%S}  ✗ packaging failed:", pkg.get("error"))

        # STEP 3 – Always start next iteration with a fresh app
        apple_script_agent.close_indesign()
        time.sleep(5)

    # STEP 4 – Get all .indd paths first
    paths, total = apple_script_agent.count_cover_indesign_files()
    if total == 0:
        print("Nothing to process – exiting.")
        sys.exit(0)

    # STEP 5 – Iterate once per file
    for idx, path in enumerate(paths, start=1):
        print(f"[{idx}/{total}]  {os.path.basename(path)}")
        print(repr(path))
        apple_script_agent.open_indesign_file(path)


        pkg = apple_script_agent.package_indesign_file(
            folder_id=folder_id,
            project_name=archived_project_path
        )
        if pkg["success"]:
            print(f"{datetime.now():%Y-%m-%d %H:%M:%S} ✓ packaged → {pkg['message']}")
        else:
            print(f"{datetime.now():%Y-%m-%d %H:%M:%S}  ✗ packaging failed:", pkg.get("error"))

        # STEP 6 – Always start next iteration with a fresh app
        apple_script_agent.close_indesign()
        time.sleep(5)


    file_checker_agent = FileCheck()
    result = file_checker_agent.verify_nonzero_file_sizes(project_directory_path)
    if result["success"]:
        print("✅ All files have non-zero sizes.")
    else:
        print(f"⚠️ Check complete with {result['checked_count']} files checked.")
        if result["empty_files"]:
            print(f"⚠️ Found {len(result['empty_files'])} empty files:")
            for idx, file in enumerate(result['empty_files'], 1):
                print(f"  {idx}. {file}")
        else:
            print("✅ No empty files found.")
    input("\nPress Enter to close the program ")

if __name__ == "__main__":
    main()
