from tkinter import *
from tkinter import messagebox
import shutil
from docx import *

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

        self.create_reports = Button(self.mainframe,
                                     font=20,
                                     text="Create Reports",
                                     relief=RIDGE,
                                     command=self.created_reports
                                     )
        self.create_reports.grid(row=2, column=1)

    def get_info(self, mainlistbox):
        self.var1 = IntVar()
        self.var2 = IntVar()

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

        self.ask_week_type = Label(self.entry_window,
                                   font=20,
                                   text="Week Type",
                                   relief=RIDGE
                                   )
        self.ask_week_type.grid(row=6, column=0)

        self.get_week_type1 = Checkbutton(self.entry_window,
                                          font=20,
                                          text="School Week",
                                          variable=self.var1,
                                          onvalue=1,
                                          offvalue=0
                                         )
        self.get_week_type1.grid(row=6, column=1)

        self.get_week_type2 = Checkbutton(self.entry_window,
                                          font=20,
                                          text="Work Week",
                                          variable=self.var2,
                                          onvalue=1,
                                          offvalue=0)
        self.get_week_type2.grid(row=6, column=2)

        self.submit_info_button = Button(self.entry_window,
                                         font=20,
                                         text="Submit Info",
                                         relief=RIDGE,
                                         command=lambda: self.retrieve_info(self.mainlistbox)
                                         )
        self.submit_info_button.grid(row=7, column=1)

    def retrieve_info(self, mainlistbox):
        self.mainlistbox = mainlistbox
        self.name = self.get_name.get()
        self.year = self.get_year.get()
        self.abteilung = self.get_abteilung.get()
        self.week_start = self.get_week_start.get()
        self.week_end = self.get_week_end.get()
        self.free_day = self.get_free_day.get()
        self.week_type = None

        if self.var1.get() == 1 and self.var2.get() == 0:
            self.week_type = "school week"
        elif self.var1.get() == 0 and self.var2.get() == 1:
            self.week_type = "work week"
        elif self.var1.get() == 1 and self.var2.get() == 1:
            messagebox.showerror("Invalid Input", "Can't have a School and Work Week!")
        else:
            messagebox.showerror("Invalid Input", "Need to select at least one week type!")

        self.Info = [self.name, self.year, self.abteilung, self.week_start, self.week_end, self.free_day, self.week_type]
        print(self.Info)

        info_string = ', '.join(map(str, self.Info))
        mainlistbox.insert(END, info_string)

    def created_reports(self):
        # Retrieve the contents of the listbox
        selected_items = self.mainlistbox.get(0, END)

        # Iterate over the selected items
        for item in selected_items:
            # Split the item by commas to separate elements
            elements = item.split(', ')

            # Access elements of each item
            name = elements[0]
            year = elements[1]
            abteilung = elements[2]
            week_start = elements[3]
            week_end = elements[4]
            free_day = elements[5]
            week_type = elements[6]

            which_doc_is_used = None

            if week_type == "school week":
                which_doc_is_used = shutil.copy("C:\\Users\\Maxim\\PycharmProjects\\pythonProject1\\ausbildungsnachweise-School-week.docx", f"ausbildungsnachweis{week_start} - {week_end}.docx")
            elif week_type == "work week":
                which_doc_is_used = shutil.copy("C:\\Users\\Maxim\\PycharmProjects\\pythonProject1\\ausbildungsnachweise-Work-week.docx", f"ausbildungsnachweis{week_start} - {week_end}.docx")

            if which_doc_is_used is not None:
                print(which_doc_is_used)

            doc = Document(which_doc_is_used)

            self.fill_doc_header(doc, "Name der/des Auszubildenden", name)
            #self.fill_doc_header(doc, "Ausbildungsjahr:", year)
            #self.fill_doc_header(doc, "Abteilung:", abteilung)
            #self.fill_doc_header(doc, "Ausbildungswoche vom:", week_start)
            #self.fill_doc_header(doc, "Bis:", week_end)

            # Do something with the elements
            #print("Name:", name)
            #print("Year:", year)
            #print("Abteilung:", abteilung)
            #print("Week start:", week_start)
            #print("Week end:", week_end)
            #print("Free day:", free_day)
            #print("Week Type:", week_type)

            doc.save(which_doc_is_used)

    def fill_doc_header(self, document, text_to_find, text_to_insert):
        for table in document.tables:
            for row in table.rows:
                for index, cell in enumerate(row.cells):
                    # Check if the cell contains the text_to_find
                    print(index)

                    if text_to_find in cell.text:
                        #print(cell.text)
                        current_index = index
                        #print(current_index)

                        next_index = current_index + 6

                        adjacent_index = row.cells[next_index]
                        adjacent_index.text = text_to_insert
                        print("Multipal")
                        break

                    """
                    elif text_to_find in cell.text:
                        current_index = index

                        next_index = current_index + 1
                        adjacent_index = row.cells[next_index]
                        adjacent_index.text = text_to_insert
                        break

                    elif text_to_find in cell.text:
                        current_index = index

                        next_index = current_index + 1
                        adjacent_index = row.cells[next_index]
                        adjacent_index.text = text_to_insert
                        break

                    elif text_to_find in cell.text:
                        current_index = index

                        next_index = current_index + 1
                        adjacent_index = row.cells[next_index]
                        adjacent_index.text = text_to_insert
                        break

                    elif text_to_find in cell.text:
                        current_index = index

                        next_index = current_index + 1
                        adjacent_index = row.cells[next_index]
                        adjacent_index.text = text_to_insert
                        break
                    """
        





if __name__ == "__main__":
    root_window = Tk()
    MainWindow(root_window)
    root_window.mainloop()