import tkinter as tk
import os
import win32com.client
import winreg as reg
import re

def run_command(command):
    # Opens PowerShell and runs the specified command
    os.system(f'start powershell -NoExit -Command {command}')

def fetch_updates():
    # Initialize the update session and service
    update_session = win32com.client.Dispatch('Microsoft.Update.Session')
    update_service = update_session.CreateUpdateSearcher()

    # Perform search to fetch all updates, installed or not
    search_result = update_service.Search('')

    # Clear the Text widget to prepare for new output
    output_text.delete(1.0, tk.END)

    if search_result.Updates.Count > 0:
        output_text.insert(tk.END, f"{search_result.Updates.Count} update(s) found.\n\n")
        for i in range(search_result.Updates.Count):
            update = search_result.Updates.Item(i)
            update_title = update.Title
            update_description = update.Description
            update_is_installed = update.IsInstalled
            update_guid = update.Identity.UpdateID  # GUID of the update
            update_size = get_update_size(update)  # Download size of the update
            update_is_uninstallable = update.IsUninstallable  # Whether the update is uninstallable
            update_kb_number = extract_kb_number(update_title)  # Extract KB number

            # Insert the update details into the text widget
            output_text.insert(tk.END, f"Update {i + 1}: {update_title}\n")
            output_text.insert(tk.END, f"Description: {update_description}\n")
            output_text.insert(tk.END, f"Installed: {get_installed_status_string(update_is_installed)}\n")
            output_text.insert(tk.END, f"GUID: {update_guid}\n")
            output_text.insert(tk.END, f"Download Size: {update_size}\n")
            output_text.insert(tk.END, f"Uninstallable: {get_uninstallable_status_string(update_is_uninstallable)}\n")
            output_text.insert(tk.END, f"KB Number: {update_kb_number}\n")
            output_text.insert(tk.END, "-" * 50 + "\n\n")
    else:
        output_text.insert(tk.END, "No updates found.\n")

def create_registry_key():
    """Create the 'ExternallyManaged' key in the registry"""
    reg_path = r"SOFTWARE\Microsoft\Enrollments"
    key_name = "ExternallyManaged"
    key_value = 1  # Value to set the key to (indicating commanaged device)
    
    try:
        # Open the registry key or create it if it doesn't exist
        reg_key = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, reg_path, 0, reg.KEY_WRITE)
        
        # Set the value for the registry key
        reg.SetValueEx(reg_key, key_name, 0, reg.REG_DWORD, key_value)
        
        # Close the registry key
        reg.CloseKey(reg_key)
        
        output_text.insert(tk.END, f"Registry key '{key_name}' created successfully in {reg_path}.\n")
    except PermissionError:
        output_text.insert(tk.END, "Permission error: Make sure to run the script as Administrator.\n")
    except Exception as e:
        output_text.insert(tk.END, f"An error occurred: {e}\n")

def remove_registry_key():
    """Remove the 'ExternallyManaged' key from the registry"""
    reg_path = r"SOFTWARE\Microsoft\Enrollments"
    key_name = "ExternallyManaged"
    
    try:
        # Open the registry key
        reg_key = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, reg_path, 0, reg.KEY_WRITE)
        
        # Delete the registry key
        reg.DeleteValue(reg_key, key_name)
        
        # Close the registry key
        reg.CloseKey(reg_key)
        
        output_text.insert(tk.END, f"Registry key '{key_name}' removed successfully from {reg_path}.\n")
    except FileNotFoundError:
        output_text.insert(tk.END, f"Registry key '{key_name}' not found in {reg_path}.\n")
    except PermissionError:
        output_text.insert(tk.END, "Permission error: Make sure to run the script as Administrator.\n")
    except Exception as e:
        output_text.insert(tk.END, f"An error occurred: {e}\n")

def get_update_size(update):
    """Safely retrieves the size of the update, or returns 'N/A' if the size is not available."""
    try:
        return format_size(update.Size)  # Attempt to get the size and format it
    except AttributeError:
        return "N/A"  # Return 'N/A' if Size is not available

def extract_kb_number(update_title):
    """Extract KB number from the update title using regex."""
    kb_number = None
    match = re.search(r"KB(\d+)", update_title)
    if match:
        kb_number = match.group(0)  # Extract the matched KB number
    return kb_number if kb_number else "N/A"

def get_installed_status_string(is_installed):
    """Converts the IsInstalled flag to a human-readable string."""
    return "Yes" if is_installed else "No"

def get_uninstallable_status_string(is_uninstallable):
    """Converts the IsUninstallable flag to a human-readable string."""
    return "Yes" if is_uninstallable else "No"

def format_size(size_in_bytes):
    """Formats the download size from bytes to a human-readable format."""
    if size_in_bytes < 1024:
        return f"{size_in_bytes} Bytes"
    elif size_in_bytes < 1048576:
        return f"{size_in_bytes / 1024:.2f} KB"
    elif size_in_bytes < 1073741824:
        return f"{size_in_bytes / 1048576:.2f} MB"
    else:
        return f"{size_in_bytes / 1073741824:.2f} GB"

# Create the main application window
root = tk.Tk()
root.title("AgentHelper")

# Set window size
root.geometry("600x600")

# Create buttons for useful winget commands
commands = {
    "List Installed Packages (Winget)": "winget list",
    "Agent Logs": "Get-Content -Path 'C:\Hexnode\Hexnode Agent\Logs\hexnodeagent.log' -wait",
    "Installer Logs": "Get-Content -Path 'C:\hexnodeinstaller\hexnodeinstaller.log' -wait",
    "Updater Logs": "Get-Content -Path 'C:\Hexnode\Hexnode Updater\Logs\hexnodeupdater.log' -wait",
    "Fetch Windows Updates": fetch_updates  # Add the new button to fetch updates
}

# Create buttons dynamically for each command
for label, command in commands.items():
    button = tk.Button(root, text=label, command=lambda c=command: run_command(c) if isinstance(c, str) else c(), font=("Arial", 12), width=30)
    button.pack(pady=5)

# Add the "Set as Commanaged" and "Remove Commanaged" buttons
set_comanaged_button = tk.Button(root, text="Co-manage device", command=create_registry_key, font=("Arial", 12), width=30)
set_comanaged_button.pack(pady=5)

remove_comanaged_button = tk.Button(root, text="Remove Co-management", command=remove_registry_key, font=("Arial", 12), width=30)
remove_comanaged_button.pack(pady=5)

# Create a Text widget to display the output
output_text = tk.Text(root, wrap=tk.WORD, width=70, height=20, font=("Arial", 10))
output_text.pack(pady=10)

# Run the application
root.mainloop()
