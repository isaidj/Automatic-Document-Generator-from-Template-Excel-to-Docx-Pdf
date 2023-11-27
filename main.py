import tkinter as tk
from document_generator import DocumentGenerator
from ui import DocumentGeneratorUI

if __name__ == "__main__":
    root = tk.Tk()
    generator = DocumentGenerator()
    app = DocumentGeneratorUI(root, generator)
    root.mainloop()
