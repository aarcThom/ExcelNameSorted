from textual.app import App, ComposeResult
from textual.widgets import Button, Label, DataTable, Header
from textual.containers import HorizontalGroup, VerticalScroll
from  tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import os

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

def write_excel(df: pd.DataFrame, in_path:str):
    file_name_wit_ext = os.path.basename(in_path)
    file_name_final = os.path.splitext(file_name_wit_ext)[0] + "_sorted.xlsx"
    contain_folder = os.path.dirname(in_path)
    revised_path = contain_folder + "/" + file_name_final

    df.to_excel(revised_path)

    return revised_path


#### NEED TO DO THIS: https://textual.textualize.io/widgets/data_table/#textual.widgets.DataTable.add_column



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
        yield Button ("Save File", id="save_file", variant="warning")


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

        if (event.button.id == "sort_last" or event.button.id == "sort_first") and self.selected_column is not None:
            # sort and update the values
            self.df[self.selected_column] = self.df[self.selected_column].apply(lambda row: sort_names(row, event.button.id == "sort_last"))
            self.rows = get_rows(self.df)
            # Find and remove the dataframe
            removeD = self.query_one(DFrame)
            removeD.remove()
            # add the sorted table
            self.mount(DFrame(rows=self.rows))

        if event.button.id == "save_file" and self.df is not None:
            test = write_excel(self.df, self.file_path)
            self.notify(f"Saved to {test}")


            

    async def on_data_table_column_selected(self, event: DataTable.ColumnSelected) -> None:
        col_key = event.column_key
        # Get the header value (first cell) from the column
        self.selected_column = str(self.query_one(DataTable).columns[event.column_key].label)
        self.notify(f"Selected column: {self.selected_column}")


    def compose(self) -> ComposeResult:
        yield Header()
        yield Importbar()


if __name__ == "__main__":
    app = ExcelSortApp()
    app.run()