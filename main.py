# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

# Press the green button in the gutter to run the script.
from gui import Fund
from read import read_list
from write import sheet
from tkinter import *
import tkinter.filedialog


if __name__ == '__main__':
    fund = Fund()
    fund.mainloop()



