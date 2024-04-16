import os
import shutil

# Specify the filename
filename = "Template_BAYERISCHE.dotx"

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the full paths for the original and copy
original_filepath = os.path.join(script_dir, filename)
copy_filepath = os.path.join(script_dir, filename + "_copy.dotx")

# Use shutil.copyfile to create the copy
shutil.copyfile(original_filepath, copy_filepath)

print(f"File '{filename}' copied to '{copy_filepath}'")
