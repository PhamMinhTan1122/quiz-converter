"""
This code creates a Tkinter GUI for a quiz converter.

The `app` function creates the GUI and starts the main loop.
The `QuizConverterGUI` class provides the functionality for the GUI.
"""

from tkinter import Tk
from module.app.gui import QuizConverterGUI

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = "com.minhtan.quizcoverter"
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

def app():
    """
    Creates the GUI and starts the main loop.

    Returns:
        The `Tk` root window.
    """

    root = Tk()
    gui = QuizConverterGUI(root)
    root.mainloop()
    return root

if __name__ == "__main__":
    """
    Starts the application.
    """

    app()