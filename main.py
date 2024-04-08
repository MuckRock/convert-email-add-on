"""
This is an DocumentCloud Add-On that,
when given a public Google Drive or Dropbox link containing EML/MSG files,
will convert them to PDFs and upload them to DocumentCloud
"""
import os
import sys
import glob
import shutil
import subprocess
from urllib.error import HTTPError
from documentcloud.addon import AddOn
from documentcloud.exceptions import APIError
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
        try:
            grab(url, "./out/")
        except HTTPError as http_error:
            self.set_message(
                "There was an issue with downloading emails from the provided URL, please ensure it is public and available."
            )
            print(f"HTTP Error: {http_error}")
            sys.exit(0)
        print("Contents of ./out/ after downloading:")
        print(os.listdir("./out/"))
        os.chdir("out")
        self.strip_white_spaces(os.getcwd())
        os.chdir("..")

    def strip_white_spaces(self, file_path):
        """Strips white space from filename before running it"""
        current_directory = os.getcwd()
        files = os.listdir(current_directory)
        for file_name in files:
            if file_name.strip() != file_name:
                old_file_path = os.path.join(current_directory, file_name)
                new_file_path = os.path.join(current_directory, file_name.strip())
                os.rename(old_file_path, new_file_path)
                # print(f"Renamed: {file_name} -> {file_name.strip()}")

    def eml_to_pdf(self, file_path):
        """Uses a java program to convert EML/MSG files to PDFs
        extracts attachments if selected"""
        file_extension = (
            os.path.splitext(os.path.basename(file_path.strip("'")))[1].lower().strip()
        )
        if not (file_extension == ".eml" or file_extension == ".msg"):
            print(f"Skipping non-EML/MSG file with {file_extension} extension")
            return

        if self.extract_attachments:
            bash_cmd = f"java -jar email.jar -a {file_path}"
            subprocess.call(bash_cmd, shell=True)
            attachments_pattern = os.path.join(
                os.path.dirname(file_path), "*attachments*"
            )
            attachments_dirs = glob.glob(attachments_pattern)
            if attachments_dirs:
                for attachments_dir in attachments_dirs:
                    shutil.move(attachments_dir, "./attach")
            else:
                print("No attachments directory found.")
            
        else:
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
        for current_path, _folders, files in os.walk("./out/"):
            for file_name in files:
                file_name = os.path.join(current_path, file_name)
                self.set_message("Attempting to convert EML/MSG files to PDFs...")
                abs_path = os.path.abspath(file_name)
                abs_path = f"'{abs_path}'"
                try:
                    self.eml_to_pdf(abs_path)
                except RuntimeError as re:
                    self.send_mail(
                        "Runtime Error for Email Conversion",
                        "Please forward this to info@documentcloud.org \n" + str(re),
                    )
                    errors += 1
                    continue
                else:
                    try:
                        if project_id is not None:
                            kwargs = {"project": project_id}
                        else:
                            kwargs = {}
                        self.set_message("Uploading converted file to DocumentCloud...")
                        file_name_no_ext = os.path.splitext(abs_path)[0].replace(
                            "'", ""
                        )
                        self.client.documents.upload(
                            f"{file_name_no_ext}.pdf", access=access_level, **kwargs
                        )
                        successes += 1
                    except OSError as e:
                        errors += 1
                        print(f"OS Error: {e}")
                        continue
                    except APIError as api_error:
                        error_message = str(api_error)
                        if "Invalid pk" in error_message:
                            self.set_message(
                                "You have provided an incorrect project ID, please try again"
                            )
                            sys.exit(0)
                        else:
                            # Handle other API errors if needed
                            print("API Error:", error_message)

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
