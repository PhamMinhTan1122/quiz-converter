import os

class Logs():
    def __init__(self):
        self.folder_path = "logs"
        self.file_name = "logs.txt"
        self.file_path = os.path.join(self.folder_path, self.file_name)

    def check_folder(self):
        if not os.path.exists(self.folder_path):
            os.makedirs(self.folder_path)

    def check_file(self, logs):
        Logs().check_folder()
        with open(self.file_path, "w") as f:
            f.write("\n".join(logs) + "\n")
    def read(self):
        with open(self.file_path, "r") as f:
            return f.read()