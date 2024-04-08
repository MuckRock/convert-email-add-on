import os
import glob

def find_attachments(directory):
    attachments_pattern = os.path.join(directory, "*attachments*")
    attachments_dirs = glob.glob(attachments_pattern)
    return attachments_dirs

if __name__ == "__main__":
    search_directory = "."  # Replace with the path to your test directory
    attachments_dirs = find_attachments(search_directory)
    print("Attachments Directories Found:")
    for attachments_dir in attachments_dirs:
        print(attachments_dir)