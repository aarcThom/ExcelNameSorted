from textual.app import App, ComposeResult
from textual.widgets import Button, Label, DataTable, Header
from textual.containers import HorizontalGroup, VerticalScroll
from  tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd

# ========= HELPER FUNCTIONS ======================================================================================================

def open_file_dialog():
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(
        filetypes=[("Excel files", "*.xlsx")],
        title="Select an Excel file"
    )
    root.destroy()
    return file_path


def sort_names(cell, sort_by_last = True):
    names = [name.strip() for name in cell.split(",")]
    last_names = [name.split(" ")[-1] for name in names]
    first_names = [name.split(" ")[0] for name in names]

    sort_list = last_names if sort_by_last else first_names
    
    _, sorted_names = zip(*sorted(zip(sort_list, names)))

    output = ""
    for i, name in enumerate(sorted_names):
        output += name
        if i != len(names) - 1:
            output += ", "
    
    return output


def get_df(file_path):
    return pd.read_excel(file_path)


def get_rows(df):
    headers = [df.columns.to_list()]
    headers.extend(df.values.tolist())
    return headers



# ======== GUI CLASSES ===========================================================================================================

class DFrame(VerticalScroll):
    def __init__(self, rows, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.rows = rows
    
    def compose(self) -> ComposeResult:
        dframe = DataTable()
        dframe.cursor_type = "column"
        dframe.add_columns(*self.rows[0])
        dframe.add_rows(self.rows[1:])
        yield dframe


class Importbar(HorizontalGroup):

    def compose(self) -> ComposeResult:
        yield Label("Click to Choose the Excel table you want sort.\nCtrl+q to quit.", id="import_label")
        yield Button("Click Here!", id="xl_file", variant="primary")


class Sortbar(HorizontalGroup):
    def compose(self) -> ComposeResult:
        yield Label("Click the column you want to sort then choose first or last name\nCtrl+q to quit.", id="sort_label")
        yield Button ("Sort by Last Name", id="sort_last", variant="success")
        yield Button ("Sort by First Name", id="sort_first", variant="error")


# ======== APP ===========================================================================================================

class ExcelSortApp(App):

    CSS_PATH = "style.tcss"
    selected_column = None
    file_path = None
    df = None
    rows = None

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "xl_file":

            # import the rows
            self.file_path = open_file_dialog()
            self.df = get_df(self.file_path)
            self.rows = get_rows(self.df)

            # Find and remove the Importbar
            importbar = self.query_one(Importbar)
            importbar.remove()

            # add the column edit bar
            self.mount(Sortbar())

            # add the dataframe
            self.mount(DFrame(rows=self.rows))

        if event.button.id == "sort_last" and self.selected_column != None:

            # sort and update the values
            self.df = self.df[self.selected_column].apply(lambda row: sort_names(row, True))
            self.rows = get_rows(self.df)

            # Find and remove the dataframe
            removeD = self.query_one(DFrame)
            removeD.remove()

            # add the sorted table
            self.mount(DFrame(rows=self.rows))





    def on_column_selected(self, event: DataTable.ColumnSelected) -> None:
        self.selected_column = event.column_key


    def compose(self) -> ComposeResult:
        yield Header()
        yield Importbar()


if __name__ == "__main__":
    app = ExcelSortApp()
    app.run()