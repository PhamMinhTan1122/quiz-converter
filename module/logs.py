"""
This class represents a log file.

The class has the following methods:

* `__init__(self)`: Initializes the class and sets the `folder_path` and `file_name` attributes.
* `check_folder(self)`: Checks if the folder exists and creates it if it does not exist.
* `check_file(self, logs)`: Writes the logs to the file.
* `write_new_line(self, logs)`: Appends the logs to the file.
* `read(self)`: Reads the contents of the file.
"""

import os

class Logs:
    def __init__(self, folder_path="logs", file_name="logs.txt"):
        """
        Initializes the class and sets the `folder_path` and `file_name` attributes.

        Args:
            folder_path: The path to the folder where the log file is located.
            file_name: The name of the log file.
        """
        self.folder_path = folder_path
        self.file_name = folder_path
        self.file_path = os.path.join(self.folder_path, self.file_name)

    def check_folder(self):
        """
        Checks if the folder exists and creates it if it does not exist.
        """
        if not os.path.exists(self.folder_path):
            os.makedirs(self.folder_path)

    def check_file(self, logs):
        """
        Writes the logs to the file.

        Args:
            logs: A list of strings to be written to the file.
        """
        self.check_folder()
        with open(self.file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(logs) + "\n")

    def write_new_line(self, logs):
        """
        Appends the logs to the file.

        Args:
            logs: A list of strings to be appended to the file.
        """
        with open(self.file_path, "a", encoding="utf-8") as f:
            f.write("\n".join(logs) + "\n")

    def read(self):
        """
        Reads the contents of the file.

        Returns:
            A list of strings containing the contents of the file.
        """
        with open(self.file_path, "r", encoding="utf-8") as f:
            return f.read()