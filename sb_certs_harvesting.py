"""Collecting cert PDFs for later study."""

import os

import win32com.client as win32

from untracked_config.accounts_and_folder_paths import production_inbox_folders


class PathConverter:
    def __init__(self):
        self.translation_table = str.maketrans({
            "/": "Ω",
            "\\": "Δ",
            "?": "θ",
            "*": "μ",
            ":": "Λ",
            "|": "Ψ",
            "<": "ξ",
            ">": "φ",
        })
        self.detranslation_table = self._create_detranslation_table()

    def _create_detranslation_table(self):
        # Create a dictionary to detranslate characters
        detranslation_table = {v: k for k, v in self.translation_table.items()}
        for c in "ΩΔθμΛΨξφ":
            detranslation_table.pop(c, None)
        return detranslation_table

    def to_path(self, subject):
        # Replace disallowed characters with Greek Unicode characters
        return subject.translate(self.translation_table)

    def from_path(self, folder_name):
        # Restore disallowed characters from Greek Unicode characters
        return folder_name.translate(self.translation_table).translate(self.detranslation_table)


def export_pdfs_from_inbox():
    # Set the file path to write any errors to
    file_path = "C:/Users/lmcglaughlin/Downloads/certs_pdfs_from_inbox/ErrorLog.txt"

    # Create a file system object and open the output file
    fso = win32.Dispatch("Scripting.FileSystemObject")
    output_file = fso.CreateTextFile(file_path, True)

    # Get the Outlook application and namespace objects
    ol_app = win32.Dispatch("Outlook.Application")
    ol_namespace = ol_app.GetNamespace("MAPI")

    # Instantiate the path converter
    path_converter = PathConverter()

    # Loop through each store in the namespace
    for ol_store in ol_namespace.Stores:
        # Loop through each folder in the store
        for ol_folder in ol_store.GetRootFolder().Folders:
            # Check if the folder has the desired path
            if ol_folder.FolderPath in production_inbox_folders:
                inbox_folder = ol_folder
                break
        else:
            continue  # Inner loop not broken
        break  # Outer loop broken, inbox folder found
    else:
        # Inbox folder not found
        output_file.WriteLine("ERROR: Inbox folder not found.")
        print("Inbox folder not found.")
        return

    # Loop through each email in the inbox folder
    for i in range(inbox_folder.items.Count, 0, -1):
        # Get the current email
        email = inbox_folder.items(i)

        # Check if the email is from the desired sender
        if email.SenderEmailAddress == "CofC@nitto.com":
            # Convert the subject to a folder name
            folder_name = path_converter.to_path(email.Subject)

            # Create the folder if it does not already exist
            folder_path = os.path.join("C:/Users/lmcglaughlin/Downloads/certs_pdfs_from_inbox/", folder_name)
            os.makedirs(folder_path, exist_ok=True)

            # Loop through each attachment in the email
            for attachment in email.Attachments:
                # Check if the attachment is a PDF file
                if attachment.FileName.endswith(".pdf"):
                    # Save the attachment to the folder
                    try:
                        attachment.SaveAsFile(os.path.join(folder_path, attachment.FileName))
                    except Exception as e:
                        output_file.WriteLine(
                            f"ERROR saving attachment from email with subject '{email.Subject}': {e}"
                        )

            # # Convert the folder name back to the subject
            # subject_line = path_converter.from_path(folder_name)

    # Close the output file
    output_file.Close()

    # Display a message box indicating the export is complete
    win32.MessageBox(None, "PDFs exported from inbox successfully.", "Export Complete", win32.MB_ICONINFORMATION)

