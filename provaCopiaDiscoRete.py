import shutil
import os

def copy_file(source_path, destination_path):
    """Copies a file from the source path to the destination path."""

    # Ensure directories exist
    os.makedirs(os.path.dirname(destination_path), exist_ok=True)

    try:
        shutil.copy2(source_path, destination_path)
        print(f"File copied successfully from '{source_path}' to '{destination_path}'")
    except FileNotFoundError:
        print(f"Error: Source file '{source_path}' not found.")
    except PermissionError:
        print(f"Error: Permission denied to write to '{destination_path}'.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


# Example usage
source_file = os.path.abspath(__file__)
print(source_file)
destination_directory = "R:/Pozzi"
destination_file = os.path.join(destination_directory, os.path.basename(source_file)) 

copy_file(source_file, destination_file)
