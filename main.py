import tkinter as tk
from module.app.gui import QuizConverterGUI
def app():
    root = tk.Tk()
    gui = QuizConverterGUI(root)
    root.mainloop()
    return app
if __name__ == "__main__":
    app()
