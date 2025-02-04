import tkinter as tk
from tkinter import filedialog
from methods import data_retrieval as dr

if __name__ == "__main__":

    # get user input for cells
    cells = dr.get_search_values()
    print(cells)

    # get user input for directory
    directory = dr.select_directory()
