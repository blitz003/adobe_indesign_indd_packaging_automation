import sys
import os

from cnt.cnt import TKFolderSelector, MakeDirectory, AppleScript

def main():
    # Create an instance of the folder selector
    folder_selector = TKFolderSelector()

    # Call the select_folder method
    selected_folder_name = folder_selector.select_folder()

    # Use the selected folder name if needed
    if selected_folder_name:
        # Perform operations with the folder name
        # Tokenize selected directory (
        folder_tokens = selected_folder_name.split("_")
        folder_id = folder_tokens[0] # 11492
        folder_semester = folder_tokens[1] # S24
        folder_last_name = folder_tokens[2] # Monroe
        folder_print_type = folder_tokens[3] # Color

        output_directory_name = folder_id +"_"+ folder_last_name # 11492_Monroe
    else:
        sys.exit() # Close the program if no project directory is selected.

    directory_handler = MakeDirectory()
    # Make Archived_Projects directory inside Documents directory
    directory_handler.create_archived_projects_directory()
    # Pass our output_directory_name to be created inside Archived_Projects
    directory_handler.create_project_directory(project_name=output_directory_name)

    # Need to move the following subdirectories into the new Project Archive directory
    # 11492_Digital_Content
    # 11492_Logs
    # 11492_Manuscript
    # 11492_Office

    documents_path = os.path.expanduser("~/Documents")
    archived_project_path = os.path.join(documents_path, "Archived_Projects", output_directory_name)
    folder_selector.copy_specific_subdirectories(destination_path=archived_project_path, folder_id=folder_id)

    # Need to create the following subdirectories into the new Project Archive directory
    # 11492_Layout
    # 11492_Printer_PDFs

    folder_selector.create_project_subdirectories(archived_project_path=archived_project_path, folder_id=folder_id)

    # Make AppleScript command to open the font software.
    apple_script_agent = AppleScript()
    if apple_script_agent.is_extensis_connect_running():
        apple_script_agent.refresh_extensis_connect()
    else:
        apple_script_agent.open_and_refresh_extensis_connect()
        apple_script_agent.minimize_extensis_connect()




    # Move Print PDF files to /11492_Printer_PDFs from /11492_Layout
    printer_pdfs_endpoint = folder_id + "_Printer_PDFs"
    archive_printer_pdfs_path = os.path.join(archived_project_path, printer_pdfs_endpoint)
    folder_id_print = folder_id + "_Print"
    layout_endpoint = folder_id + "_Layout"
    project_layout_path = os.path.join(folder_selector.folder_path, layout_endpoint)
    folder_selector.copy_print_files(
        project_layout_path=project_layout_path,
        archive_printer_pdfs_path=archive_printer_pdfs_path,
        folder_id_print=folder_id_print
    )

    # 4/3/2025
    result = apple_script_agent.select_and_process_indesign_files()

    # Optionally, show detailed results
    if result["results"]:
        print("\nDetailed results:")
        for item in result["results"]:
            status = "Success" if item["success"] else "Failed"
            print(f"- {item['file']}: {status}")

if __name__ == "__main__":
    main()