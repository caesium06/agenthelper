import tkinter as tk
import os

def run_command(command):
    # Opens PowerShell and runs the specified command
    os.system(f'start powershell -NoExit -Command {command}')

# Create the main application window
root = tk.Tk()
root.title("AgentLogs")

# Set window size
root.geometry("400x400")

# Create buttons for useful winget commands
commands = {
    "List Installed Packages": "winget list",
    "Show Sources": "winget source list",
    "Update All Packages": "winget upgrade --all",
    "Search for Packages": "winget search",
    "Display Winget Settings": "winget settings",
    "Help/Available Commands": "winget --help"
}

for label, command in commands.items():
    button = tk.Button(root, text=label, command=lambda c=command: run_command(c), font=("Arial", 12), width=30)
    button.pack(pady=5)

# Run the application
root.mainloop()
