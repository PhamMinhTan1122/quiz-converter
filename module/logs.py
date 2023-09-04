import os

class Logs:
    def __init__(self, folder_name="logs", file_name="tmp.txt"):
        self.folder_name = folder_name
        self.file_name = file_name
        self.file_path = os.path.join(self.folder_name, self.file_name).replace("\\", "/")


    def check_folder(self):
        if not os.path.exists(self.folder_name):
            os.mkdir(self.folder_name)

    def check_file(self):
        self.remove()
        self.check_folder()
        with open(self.file_path, "w", encoding="utf-8") as f:
            f.close()

    def write_new_line(self, logs):
        with open(self.file_path, "a", encoding="utf-8") as f:
            f.write("\n".join(logs) + "\n")

    def read(self):
        with open(self.file_path, "r", encoding="utf-8") as f:
            return f.read()

    def remove(self):
        if os.path.exists(self.file_path):
            os.remove(self.file_path)
