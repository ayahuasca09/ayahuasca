import subprocess
import os


def run_unreal_editor_with_script(ue_path, project_path, script_path):
    # Construct the command
    command = [
        ue_path,  # Path to the Unreal Engine executable
        project_path,  # Path to your .uproject file
        "-run=pythonscript",
        f"-script={script_path}"  # Path to your Python script
    ]

    # Print the command to verify correctness
    print("Running command:", " ".join(command))

    # Execute the command
    try:
        subprocess.run(command, check=True)
        print("Script executed successfully.")
    except subprocess.CalledProcessError as e:
        print("An error occurred while executing the script:", e)


# Set your paths here
ue_executable_path = "C:/Path/To/UnrealEditor.exe"  # Update this path
project_file_path = "C:/Path/To/YourProject.uproject"  # Update this path
python_script_path = "C:/Path/To/import_csv_to_datatable.py"  # Update this path

# Run the function
run_unreal_editor_with_script(ue_executable_path, project_file_path, python_script_path)