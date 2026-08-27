import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
import os
import sys


# ============================================================
# GLOBAL VARIABLES
# ============================================================

workbook = None

# Original headers from selected worksheet
current_headers = []

# Column information
# {
#     index: {
#         "header": "...",
#         "include_empty": False,
#         "deleted": False
#     }
# }
column_settings = {}

# Prevent excessive preview updates
preview_updating = False


# ============================================================
# MAIN WINDOW
# ============================================================

root = tk.Tk()

# Set application/window icon
if getattr(sys, "frozen", False):
    # Running as EXE
    icon_path = os.path.join(sys._MEIPASS, "excelxml.ico")
else:
    # Running from Python
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excelxml.ico")

root.iconbitmap(icon_path)

root.title("Excel → XML")
root.geometry("1250x850")
root.minsize(1000, 700)


# ============================================================
# TKINTER VARIABLES
# ============================================================

excel_path = tk.StringVar()

root_element = tk.StringVar(value="Data")

record_element = tk.StringVar(value="Record")

status_text = tk.StringVar(
    value="Please select an Excel file."
)


# ============================================================
# MAKE VALID XML TAG
# ============================================================

def make_valid_xml_tag(name, fallback="Column"):

    if name is None:
        return fallback

    name = str(name).strip()

    if not name:
        return fallback

    # Replace invalid characters
    name = re.sub(
        r"[^a-zA-Z0-9_.-]",
        "_",
        name
    )

    # XML tag cannot start with number
    if name[0].isdigit():
        name = "_" + name

    # Avoid XML reserved beginning
    if name.lower().startswith("xml"):
        name = "_" + name

    return name


# ============================================================
# CLEAR TREEVIEW
# ============================================================

def clear_tree():

    for item in tree.get_children():
        tree.delete(item)

    tree["columns"] = ()


# ============================================================
# BROWSE EXCEL
# ============================================================

def browse_excel():

    global workbook

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel Files", "*.xlsx"),
            ("Excel Macro Files", "*.xlsm"),
            ("All Files", "*.*")
        ]
    )

    if not file_path:
        return

    try:

        if workbook is not None:

            try:
                workbook.close()
            except Exception:
                pass

        workbook = load_workbook(
            file_path,
            read_only=True,
            data_only=True
        )

        excel_path.set(file_path)

        sheet_names = workbook.sheetnames

        worksheet_combo["values"] = sheet_names

        if sheet_names:

            worksheet_combo.current(0)

            load_worksheet()

        status_text.set(
            f"Excel loaded successfully. "
            f"{len(sheet_names)} worksheet(s) found."
        )

    except Exception as e:

        workbook = None

        messagebox.showerror(
            "Excel Error",
            f"Unable to open Excel file.\n\n{e}"
        )


# ============================================================
# LOAD WORKSHEET
# ============================================================

def load_worksheet(event=None):

    global current_headers
    global column_settings

    if workbook is None:
        return

    sheet_name = worksheet_combo.get()

    if not sheet_name:
        return

    try:

        worksheet = workbook[sheet_name]

        # ----------------------------------------------------
        # Read first row as headers
        # ----------------------------------------------------

        rows = worksheet.iter_rows(
            values_only=True
        )

        headers = next(rows, None)

        if headers is None:

            clear_tree()

            current_headers = []

            column_settings = {}

            update_column_options()

            status_text.set(
                f"Worksheet '{sheet_name}' is empty."
            )

            return

        headers = list(headers)

        # ----------------------------------------------------
        # Find actual last used column
        # ----------------------------------------------------

        last_column = 0

        for index, value in enumerate(headers):

            if value is not None:

                if (
                    not isinstance(value, str)
                    or value.strip()
                ):
                    last_column = index + 1

        # If header row itself is strange, use max columns
        if last_column == 0:
            last_column = len(headers)

        headers = headers[:last_column]

        current_headers = headers

        # ----------------------------------------------------
        # Reset column settings for new worksheet
        # ----------------------------------------------------

        column_settings = {}

        for index, header in enumerate(headers):

            if header is None:

                display_name = f"Column {index + 1}"

            else:

                display_name = str(header).strip()

                if not display_name:

                    display_name = (
                        f"Column {index + 1}"
                    )

            column_settings[index] = {
                "header": display_name,
                "include_empty": False,
                "deleted": False
            }

        # ----------------------------------------------------
        # Update column options
        # ----------------------------------------------------

        update_column_options()

        # ----------------------------------------------------
        # Preview
        # ----------------------------------------------------

        refresh_preview()

        # ----------------------------------------------------
        # Default XML root = worksheet name
        # ----------------------------------------------------

        root_element.set(
            make_valid_xml_tag(
                sheet_name,
                "Data"
            )
        )

        status_text.set(
            f"Worksheet '{sheet_name}' loaded."
        )

    except Exception as e:

        messagebox.showerror(
            "Worksheet Error",
            f"Unable to read worksheet.\n\n{e}"
        )


# ============================================================
# UPDATE COLUMN OPTIONS
# ============================================================

def update_column_options():

    # Remove existing option widgets
    for widget in column_options_inner.winfo_children():
        widget.destroy()

    if not column_settings:

        ttk.Label(
            column_options_inner,
            text="No columns available."
        ).pack(
            anchor="w",
            padx=10,
            pady=10
        )

        return

    for index, settings in column_settings.items():

        create_column_option(
            index,
            settings
        )


# ============================================================
# CREATE ONE COLUMN OPTION
# ============================================================

def create_column_option(index, settings):

    frame = tk.Frame(
        column_options_inner,
        bd=1,
        relief="solid",
        padx=8,
        pady=6
    )

    frame.pack(
        fill="x",
        padx=5,
        pady=3
    )

    # --------------------------------------------------------
    # Column Name
    # --------------------------------------------------------

    name_label = tk.Label(
        frame,
        text=settings["header"],
        font=("Arial", 10, "bold"),
        anchor="w"
    )

    name_label.pack(
        side="left",
        fill="x",
        expand=True
    )

    # --------------------------------------------------------
    # Include Empty Checkbox
    # --------------------------------------------------------

    include_var = tk.BooleanVar(
        value=settings["include_empty"]
    )

    checkbox = tk.Checkbutton(
        frame,
        text="Include Empty",
        variable=include_var,
        command=lambda idx=index, var=include_var:
        toggle_include_empty(idx, var),
        font=("Arial", 9)
    )

    checkbox.pack(
        side="left",
        padx=10
    )

    # --------------------------------------------------------
    # Delete Button
    # --------------------------------------------------------

    delete_button = tk.Button(
        frame,
        text="Delete",
        width=8,
        command=lambda idx=index:
        delete_column(idx)
    )

    delete_button.pack(
        side="right",
        padx=5
    )

    # Store references
    settings["frame"] = frame
    settings["checkbox"] = checkbox
    settings["include_var"] = include_var
    settings["name_label"] = name_label

    update_column_color(index)


# ============================================================
# UPDATE COLUMN COLOR
# ============================================================

def update_column_color(index):

    if index not in column_settings:
        return

    settings = column_settings[index]

    frame = settings.get("frame")

    if frame is None:
        return

    # --------------------------------------------------------
    # GREEN = Include Empty selected
    # --------------------------------------------------------

    if settings["include_empty"]:

        frame.configure(
            bg="#90EE90"
        )

        settings["name_label"].configure(
            bg="#90EE90"
        )

        settings["checkbox"].configure(
            bg="#90EE90",
            activebackground="#90EE90"
        )

    else:

        frame.configure(
            bg="#F2F2F2"
        )

        settings["name_label"].configure(
            bg="#F2F2F2"
        )

        settings["checkbox"].configure(
            bg="#F2F2F2",
            activebackground="#F2F2F2"
        )


# ============================================================
# TOGGLE INCLUDE EMPTY
# ============================================================

def toggle_include_empty(index, variable):

    if index not in column_settings:
        return

    column_settings[index]["include_empty"] = (
        variable.get()
    )

    update_column_color(index)

    if variable.get():

        status_text.set(
            f"'{column_settings[index]['header']}' "
            f"will include empty XML tags."
        )

    else:

        status_text.set(
            f"'{column_settings[index]['header']}' "
            f"will omit empty XML tags."
        )


# ============================================================
# DELETE COLUMN
# ============================================================

def delete_column(index):

    if index not in column_settings:
        return

    column_name = column_settings[index]["header"]

    result = messagebox.askyesno(
        "Delete Column",
        f"Do you want to remove this column from XML?\n\n"
        f"Column: {column_name}\n\n"
        f"This column will not be exported."
    )

    if not result:
        return

    # Mark as deleted
    column_settings[index]["deleted"] = True

    # Remove it from UI
    update_column_options()

    # Refresh Excel preview
    refresh_preview()

    status_text.set(
        f"Column '{column_name}' removed from XML export."
    )


# ============================================================
# REFRESH PREVIEW
# ============================================================

def refresh_preview():

    if workbook is None:
        return

    sheet_name = worksheet_combo.get()

    if not sheet_name:
        return

    try:

        worksheet = workbook[sheet_name]

        clear_tree()

        # ----------------------------------------------------
        # Get active columns
        # ----------------------------------------------------

        active_columns = []

        for index, settings in column_settings.items():

            if not settings["deleted"]:

                active_columns.append(
                    index
                )

        if not active_columns:

            status_text.set(
                "All columns have been removed."
            )

            return

        # ----------------------------------------------------
        # Configure TreeView columns
        # ----------------------------------------------------

        tree_columns = []

        for index in active_columns:

            column_name = (
                column_settings[index]["header"]
            )

            # Make unique TreeView ID
            column_id = f"col_{index}"

            tree_columns.append(
                column_id
            )

        tree["columns"] = tree_columns

        tree["show"] = "headings"

        for index in active_columns:

            column_id = f"col_{index}"

            column_name = (
                column_settings[index]["header"]
            )

            tree.heading(
                column_id,
                text=column_name
            )

            tree.column(
                column_id,
                width=150,
                minwidth=80,
                anchor="w"
            )

        # ----------------------------------------------------
        # Read rows
        # ----------------------------------------------------

        rows = worksheet.iter_rows(
            values_only=True
        )

        # Skip header
        next(rows, None)

        row_count = 0

        for row in rows:

            # Skip completely empty rows
            if all(
                value is None
                or (
                    isinstance(value, str)
                    and not value.strip()
                )
                for value in row
            ):
                continue

            values = []

            for index in active_columns:

                if index < len(row):

                    value = row[index]

                else:

                    value = ""

                if value is None:

                    value = ""

                values.append(
                    str(value)
                )

            tree.insert(
                "",
                "end",
                values=values
            )

            row_count += 1

            if row_count >= 100:
                break

        status_text.set(
            f"Worksheet '{sheet_name}' loaded. "
            f"{len(active_columns)} active column(s), "
            f"showing {row_count} row(s)."
        )

    except Exception as e:

        messagebox.showerror(
            "Preview Error",
            f"Unable to display worksheet.\n\n{e}"
        )


# ============================================================
# EXPORT XML
# ============================================================

def export_xml():

    if workbook is None:

        messagebox.showwarning(
            "Warning",
            "Please select an Excel file first."
        )

        return

    sheet_name = worksheet_combo.get()

    if not sheet_name:

        messagebox.showwarning(
            "Warning",
            "Please select a worksheet."
        )

        return

    # --------------------------------------------------------
    # Get active columns
    # --------------------------------------------------------

    active_columns = []

    for index, settings in column_settings.items():

        if not settings["deleted"]:

            active_columns.append(
                index
            )

    if not active_columns:

        messagebox.showwarning(
            "Warning",
            "There are no columns available for export."
        )

        return

    # --------------------------------------------------------
    # XML names
    # --------------------------------------------------------

    root_name = make_valid_xml_tag(
        root_element.get(),
        "Data"
    )

    record_name = make_valid_xml_tag(
        record_element.get(),
        "Record"
    )

    # --------------------------------------------------------
    # Ask save location
    # --------------------------------------------------------

    output_path = filedialog.asksaveasfilename(
        title="Export XML",
        initialfile=f"{sheet_name}.xml",
        defaultextension=".xml",
        filetypes=[
            ("XML Files", "*.xml"),
            ("All Files", "*.*")
        ]
    )

    if not output_path:
        return

    try:

        worksheet = workbook[sheet_name]

        # ----------------------------------------------------
        # Read rows
        # ----------------------------------------------------

        rows = worksheet.iter_rows(
            values_only=True
        )

        # Skip header
        next(rows, None)

        # ----------------------------------------------------
        # Create XML root
        # ----------------------------------------------------

        root_xml = ET.Element(
            root_name
        )

        record_count = 0

        # ====================================================
        # PROCESS EVERY EXCEL ROW
        # ====================================================

        for row in rows:

            # ------------------------------------------------
            # Skip completely empty rows
            # ------------------------------------------------

            if all(
                value is None
                or (
                    isinstance(value, str)
                    and not value.strip()
                )
                for value in row
            ):
                continue

            # ------------------------------------------------
            # Create Record
            # ------------------------------------------------

            record = ET.SubElement(
                root_xml,
                record_name
            )

            used_tags = set()

            # ------------------------------------------------
            # Process active columns only
            # ------------------------------------------------

            for index in active_columns:

                settings = column_settings[index]

                header = settings["header"]

                include_empty = (
                    settings["include_empty"]
                )

                # --------------------------------------------
                # Get cell value
                # --------------------------------------------

                if index < len(row):

                    value = row[index]

                else:

                    value = None

                # --------------------------------------------
                # Check empty
                # --------------------------------------------

                is_empty = (
                    value is None
                    or (
                        isinstance(value, str)
                        and not value.strip()
                    )
                )

                # =================================================
                # DEFAULT:
                #
                # Empty value -> DON'T CREATE XML TAG
                # =================================================

                if is_empty and not include_empty:

                    continue

                # =================================================
                # INCLUDE EMPTY SELECTED:
                #
                # Empty value -> CREATE EMPTY XML TAG
                # =================================================

                tag_name = make_valid_xml_tag(
                    header,
                    f"Column{index + 1}"
                )

                # --------------------------------------------
                # Handle duplicate XML names
                # --------------------------------------------

                original_tag = tag_name

                counter = 2

                while tag_name in used_tags:

                    tag_name = (
                        f"{original_tag}_{counter}"
                    )

                    counter += 1

                used_tags.add(
                    tag_name
                )

                # --------------------------------------------
                # Create XML element
                # --------------------------------------------

                element = ET.SubElement(
                    record,
                    tag_name
                )

                # --------------------------------------------
                # Empty
                # --------------------------------------------

                if is_empty:

                    element.text = None

                # --------------------------------------------
                # Non-empty
                # --------------------------------------------

                else:

                    if isinstance(value, str):

                        value = value.strip()

                    element.text = str(
                        value
                    )

            record_count += 1

        # ====================================================
        # PRETTY XML
        # ====================================================

        xml_bytes = ET.tostring(
            root_xml,
            encoding="utf-8"
        )

        pretty_xml = minidom.parseString(
            xml_bytes
        ).toprettyxml(
            indent="  ",
            encoding="utf-8"
        )

        # ----------------------------------------------------
        # Remove blank lines generated by minidom
        # ----------------------------------------------------

        with open(
            output_path,
            "wb"
        ) as file:

            file.write(
                pretty_xml
            )

        # ----------------------------------------------------
        # Success
        # ----------------------------------------------------

        status_text.set(
            f"XML exported successfully. "
            f"{record_count} record(s)."
        )

        messagebox.showinfo(
            "Export Successful",
            f"XML file created successfully!\n\n"
            f"Worksheet: {sheet_name}\n"
            f"Records: {record_count}\n"
            f"Active Columns: {len(active_columns)}\n\n"
            f"File:\n{output_path}"
        )

    except Exception as e:

        messagebox.showerror(
            "Export Error",
            f"Unable to export XML.\n\n{e}"
        )


# ============================================================
# TITLE
# ============================================================

title = ttk.Label(
    root,
    text="Excel → XML",
    font=(
        "Arial",
        26,
        "bold"
    )
)

title.pack(
    pady=20
)


# ============================================================
# EXCEL FILE
# ============================================================

ttk.Label(
    root,
    text="Excel File:",
    font=("Arial", 12 ,"bold")
).pack(
    anchor="w",
    padx=35
)


file_frame = ttk.Frame(
    root
)

file_frame.pack(
    fill="x",
    padx=35,
    pady=6
)


excel_entry = ttk.Entry(
    file_frame,
    textvariable=excel_path,
    font=("Arial", 11)
)

excel_entry.pack(
    side="left",
    fill="x",
    expand=True
)


browse_button = ttk.Button(
    file_frame,
    text="Browse",
    command=browse_excel
)

browse_button.pack(
    side="right",
    padx=(10, 0)
)


# ============================================================
# WORKSHEET
# ============================================================

ttk.Label(
    root,
    text="Worksheet:",
    font=("Arial", 12 ,"bold")
).pack(
    anchor="w",
    padx=35,
    pady=(10, 4)
)


worksheet_combo = ttk.Combobox(
    root,
    state="readonly",
    font=("Arial", 11)
)

worksheet_combo.pack(
    fill="x",
    padx=35
)


worksheet_combo.bind(
    "<<ComboboxSelected>>",
    load_worksheet
)


# ============================================================
# EXCEL PREVIEW LABEL
# ============================================================

ttk.Label(
    root,
    text="Excel Data Preview:",
    font=(
        "Arial",
        12,
        "bold"
    )
).pack(
    anchor="w",
    padx=35,
    pady=(12, 4)
)


# ============================================================
# PREVIEW FRAME
# ============================================================

preview_frame = ttk.Frame(
    root
)

preview_frame.pack(
    fill="both",
    expand=True,
    padx=35,
    pady=4
)


# ============================================================
# TREEVIEW
# ============================================================

tree = ttk.Treeview(
    preview_frame
)

tree.pack(
    side="left",
    fill="both",
    expand=True
)


# ============================================================
# VERTICAL SCROLLBAR
# ============================================================

vertical_scrollbar = ttk.Scrollbar(
    preview_frame,
    orient="vertical",
    command=tree.yview
)

vertical_scrollbar.pack(
    side="right",
    fill="y"
)

tree.configure(
    yscrollcommand=vertical_scrollbar.set
)


# ============================================================
# HORIZONTAL SCROLLBAR
# ============================================================

horizontal_scrollbar = ttk.Scrollbar(
    root,
    orient="horizontal",
    command=tree.xview
)

horizontal_scrollbar.pack(
    fill="x",
    padx=35
)

tree.configure(
    xscrollcommand=horizontal_scrollbar.set
)


# ============================================================
# COLUMN OPTIONS TITLE
# ============================================================

ttk.Label(
    root,
    text="Column Options:",
    font=(
        "Arial",
        12,
        "bold"
    )
).pack(
    anchor="w",
    padx=35,
    pady=(10, 3)
)


ttk.Label(
    root,
    text=(
        "By default, empty Excel cells are omitted. "
        "Select 'Include Empty' to create an empty XML tag. "
        "Use 'Delete' to completely exclude a column."
    ),
    font=("Arial", 9)
).pack(
    anchor="w",
    padx=35,
    pady=(0, 5)
)


# ============================================================
# COLUMN OPTIONS OUTER FRAME
# ============================================================

column_options_outer = tk.Frame(
    root,
    bd=1,
    relief="sunken"
)

column_options_outer.pack(
    fill="x",
    padx=35,
    pady=4
)

column_options_outer.configure(
    height=145
)

column_options_outer.pack_propagate(
    False
)


# ============================================================
# COLUMN OPTIONS CANVAS
# ============================================================

column_canvas = tk.Canvas(
    column_options_outer,
    highlightthickness=0
)

column_canvas.pack(
    side="left",
    fill="both",
    expand=True
)


column_options_scrollbar = ttk.Scrollbar(
    column_options_outer,
    orient="vertical",
    command=column_canvas.yview
)

column_options_scrollbar.pack(
    side="right",
    fill="y"
)


column_canvas.configure(
    yscrollcommand=column_options_scrollbar.set
)


# ============================================================
# INNER FRAME
# ============================================================

column_options_inner = tk.Frame(
    column_canvas
)

column_window = column_canvas.create_window(
    (0, 0),
    window=column_options_inner,
    anchor="nw"
)


# ============================================================
# UPDATE CANVAS SCROLL REGION
# ============================================================

def update_column_canvas(event=None):

    column_canvas.configure(
        scrollregion=column_canvas.bbox("all")
    )

    column_canvas.itemconfigure(
        column_window,
        width=column_canvas.winfo_width()
    )


column_options_inner.bind(
    "<Configure>",
    update_column_canvas
)

column_canvas.bind(
    "<Configure>",
    update_column_canvas
)


# ============================================================
# XML SETTINGS
# ============================================================

xml_settings_frame = ttk.Frame(
    root
)

xml_settings_frame.pack(
    fill="x",
    padx=35,
    pady=8
)


# XML ROOT

ttk.Label(
    xml_settings_frame,
    text="XML Root:",
    font=("Arial", 10)
).grid(
    row=0,
    column=0,
    padx=(0, 6)
)


root_entry = ttk.Entry(
    xml_settings_frame,
    textvariable=root_element,
    width=25
)

root_entry.grid(
    row=0,
    column=1,
    padx=(0, 30)
)


# XML RECORD

ttk.Label(
    xml_settings_frame,
    text="XML Record:",
    font=("Arial", 10)
).grid(
    row=0,
    column=2,
    padx=(0, 6)
)


record_entry = ttk.Entry(
    xml_settings_frame,
    textvariable=record_element,
    width=25
)

record_entry.grid(
    row=0,
    column=3
)


# ============================================================
# STATUS
# ============================================================

status_label = ttk.Label(
    root,
    textvariable=status_text,
    font=("Arial", 9)
)

status_label.pack(
    anchor="w",
    padx=35,
    pady=3
)


# ============================================================
# EXPORT BUTTON
# ============================================================

export_button = ttk.Button(
    root,
    text="EXPORT TO XML",
    command=export_xml
)

export_button.pack(
    pady=5,
    ipadx=35,
    ipady=5
)


# ============================================================
# START APPLICATION
# ============================================================

root.mainloop()