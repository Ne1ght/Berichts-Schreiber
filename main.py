from tkinter import *
#from docx import *

class MainWindow():
    def __init__(self, root_window):
        self.mainroot = root_window
        self.mainroot.title("Berichts-Schreiber")
        self.mainroot.geometry("860x600")

        self.mainframe = Frame()
        self.mainframe.pack()

        self.mainlabel = Label(self.mainframe,
                               font=30,
                               text="Berichts-Schreiber",
                               relief=RIDGE)
        self.mainlabel.grid(row=0, column=1)

        self.mainlistbox = Listbox(self.mainframe,
                                   font=20)
        self.mainlistbox.grid(row=1, column=1)

        self.get_info_button = Button(self.mainframe,
                                      font=20,
                                      text="Input Info",
                                      relief=RIDGE,
                                      command=lambda: self.get_info(self.mainlistbox))
        self.get_info_button.grid(row=0, column=0)

    def get_info(self, mainlistbox):
        self.entry_window = Toplevel()
        self.entry_window.title("Input Window")
        self.entry_window.geometry("620x400")

        self.mainlistbox = mainlistbox

        self.ask_name = Label(self.entry_window,
                              font=20,
                              text="Enter Name",
                              relief=RIDGE
                              )
        self.ask_name.grid(row=0, column=0)

        self.get_name = Entry(self.entry_window,
                              width=50
                              )
        self.get_name.grid(row=0, column=2)

        self.ask_year = Label(self.entry_window,
                              font=20,
                              text="Enter Year",
                              relief=RIDGE
                              )
        self.ask_year.grid(row=1, column=0)

        self.get_year = Entry(self.entry_window)
        self.get_year.grid(row=1, column=2)

        self.ask_abteilung = Label(self.entry_window,
                                   font=20,
                                   text="Enter Abteilung",
                                   relief=RIDGE
                                   )
        self.ask_abteilung.grid(row=2, column=0)

        self.get_abteilung = Entry(self.entry_window)
        self.get_abteilung.grid(row=2, column=2)

        self.ask_week_start = Label(self.entry_window,
                                    font=20,
                                    text="Week Start",
                                    relief=RIDGE
                                    )
        self.ask_week_start.grid(row=3, column=0)

        self.get_week_start = Entry(self.entry_window)
        self.get_week_start.grid(row=3, column=2)

        self.ask_week_end = Label(self.entry_window,
                                    font=20,
                                    text="Week End",
                                    relief=RIDGE
                                    )
        self.ask_week_end.grid(row=4, column=0)

        self.get_week_end = Entry(self.entry_window)
        self.get_week_end.grid(row=4, column=2)

        self.ask_free_day = Label(self.entry_window,
                                  font=20,
                                  text="Free Day",
                                  relief=RIDGE
                                  )
        self.ask_free_day.grid(row=5, column=0)

        self.get_free_day = Entry(self.entry_window)
        self.get_free_day.grid(row=5, column=2)

        self.submit_info_button = Button(self.entry_window,
                                         font=20,
                                         text="Submit Info",
                                         relief=RIDGE,
                                         command=lambda: self.retrieve_info(self.mainlistbox)
                                         )
        self.submit_info_button.grid(row=6, column=1)

    def retrieve_info(self, mainlistbox):
        self.mainlistbox = mainlistbox
        self.name = self.get_name.get()
        self.year = self.get_year.get()
        self.abteilung = self.get_abteilung.get()
        self.week_start = self.get_week_start.get()
        self.week_end = self.get_week_end.get()
        self.free_day = self.get_free_day.get()

        self.Info = [self.name, self.year, self.abteilung, self.week_start, self.week_end, self.free_day]
        print(self.Info)

        info_string = ', '.join(map(str, self.Info))
        mainlistbox.insert(END, info_string)
        #need to change




if __name__ == "__main__":
    root_window = Tk()
    MainWindow(root_window)
    root_window.mainloop()