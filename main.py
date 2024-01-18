"""
This is an DocumentCloud Add-On that,
when given a public Google Drive or Dropbox link containing EML/MSG files,
will convert them to PDFs and upload them to DocumentCloud
"""
import os
import sys
import stat
import glob
import shutil
import subprocess
import requests
from documentcloud.addon import AddOn
from clouddl import grab


class ConvertEmail(AddOn):
    """DocumentCloud Add-On that converts EML/MSG files to PDFs and uploads them to DocumentCloud"""
    extract_attachments = False
    def check_permissions(self):
        """The user must be a verified journalist to upload a document"""
        self.set_message("Checking permissions...")
        user = self.client.users.get("me")
        if not user.verified_journalist:
            self.set_message(
                "You need to be verified to use this add-on. Please verify your "
                "account here: https://airtable.com/shrZrgdmuOwW0ZLPM"
            )
            sys.exit()

    def fetch_files(self, url):
        """Fetch the files from either a cloud share link or any public URL"""
        self.set_message("Retrieving EML/MSG files...")
        os.makedirs(os.path.dirname("./out/"), exist_ok=True)
        os.makedirs(os.path.dirname("./attach/"), exist_ok=True)
        print("Current working directory in fetch_files")
        print(os.getcwd())
        downloaded = grab(url, "./out/")
        print("List of directories in fetch_files")
        print(os.listdir(os.getcwd()))
        
    def eml_to_pdf(self, file_path):
        """Uses a java program to convert EML/MSG files to PDFs
        extracts attachments if selected"""
        file_extension = os.path.splitext(os.path.basename(file_path.strip("'")))[1].lower().strip()
        if not (file_extension == ".eml" or file_extension == ".msg"):
            print(f"Skipping non-EML/MSG file with {file_extension} extension")
            return

        if self.extract_attachments:
            attachments_pattern = os.path.join(os.path.dirname(file_path), "EMLs", "*attachments*")
            attachments_dirs = glob.glob(attachments_pattern)
            if attachments_dirs:
                for attachments_dir in attachments_dirs:
                    shutil.move(attachments_dir, "./attach")
            else:
                print("No attachments directory found.")
        
        bash_cmd = f"java -jar email.jar {file_path}"
        subprocess.call(bash_cmd, shell=True)

    def main(self):
        """Fetches files from Google Drive/Dropbox,
        converts them to EMLs, extracts attachments,
        uploads PDFs to DocumentCloud, zips attachments for download.
        """
        url = self.data["url"]
        self.check_permissions()
        self.fetch_files(url)
        if "attachments" in self.data:
            self.extract_attachments = True
        access_level = self.data["access_level"]
        project_id = self.data.get("project_id")
        successes = 0
        errors = 0
        for current_path, folders, files in os.walk("./out/"):
            for file_name in files:
                file_name = os.path.join(current_path, file_name)
                self.set_message("Attempting to convert EML/MSG files to PDFs...")
                abs_path = os.path.abspath(file_name)
                abs_path = f"'{abs_path}'"
                try:
                    self.eml_to_pdf(abs_path)
                except RuntimeError as re:
                    self.send_mail("Runtime Error for Email Conversion", "Please forward this to info@documentcloud.org \n" + str(re))
                    errors += 1
                    continue
                else:
                    try:
                        if project_id is not None:
                            kwargs = {"project": project_id}
                        else:
                            kwargs = {}
                        self.set_message("Uploading converted file to DocumentCloud...")
                        file_name_no_ext = os.path.splitext(abs_path)[0].replace("'", "")
                        self.client.documents.upload(f"{file_name_no_ext}.pdf", access=access_level, **kwargs)
                        successes += 1
                    except OSError as e: 
                        errors +=1
                        continue
                        

        if self.extract_attachments:
            attachments_exist = any(os.listdir("./attach"))
            if attachments_exist:
                subprocess.call("zip -q -r attachments.zip attach", shell=True)
                self.upload_file(open("attachments.zip"))
            else:
                print("No attachments found")

        sfiles = "file" if successes == 1 else "files"
        efiles = "file" if errors == 1 else "files"
        self.set_message(f"Converted {successes} {sfiles}, skipped {errors} {efiles}")
        shutil.rmtree("./out", ignore_errors=False, onerror=None)
        shutil.rmtree("./attach", ignore_errors=False, onerror=None)


if __name__ == "__main__":
    ConvertEmail().main()
