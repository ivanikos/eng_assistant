import random
from tkinter import *

base = Tk()

a = Label(base, text="After() Demo")
a.pack()

contrive = Frame(base, width=450, height=500)
contrive.pack()

words = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri','Sat','Sun']
#Display words randomly one after the other.
def display_weekday():
    if not words:
        return

    rand = random.choice(words)
    character_frame = Label(contrive, text=rand)
    character_frame.pack()
    contrive.after(500,display_weekday)
    words.remove(rand)

base.after(0, display_weekday)
base.mainloop()