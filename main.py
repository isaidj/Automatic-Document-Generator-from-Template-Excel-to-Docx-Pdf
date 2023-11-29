import tkinter as tk
from document_generator import DocumentGenerator

# from ui import DocumentGeneratorUI
from new_ui import App

if __name__ == "__main__":
    generator = DocumentGenerator()
    app = App(generator)
