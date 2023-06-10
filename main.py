import tkinter as tk
from module.app.gui import QuizConverterGUI

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = "com.minhtan.quizcoverter"
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass
def app():
    root = tk.Tk()
    gui = QuizConverterGUI(root)
    root.mainloop()
    return app
if __name__ == "__main__":
    app()
