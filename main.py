import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
from tkinter import StringVar
import os
import pandas as pd

# Globale Variablen für GUI-Einstellungen
SCROLLBOX_HEIGHT = 1000
FONT = ("Helvetica", 12)
FONT_BOLD = ("Helvetica", 14, 'bold')
FONT_SMALL = ("Helvetica", 10)
BACKGROUND_COLOR = "#2e2e2e"
FRAME_COLOR = "#3e3e3e"
BUTTON_COLOR = "#4e4e4e"
TEXT_COLOR = "#e0e0e0"
ALIGNMENT_CENTER = 'ew'
ALIGNMENT_LEFT = 'w'
LEFT_MARGIN = 20
FRAME_BORDER = 2
FRAME_RELIEF = 'groove'
COLUMN_WIDTH = 150 * 1.5
TEXTBOX_HEIGHT = 30
CHOICEBOX_WIDTH = 25

# Variable für GUI-Breite (Multiplikator der Breite jeder Spalte)
GUI_WIDTH_FACTOR = 2.0  # Setze eine sinnvolle Breite für die GUI-Anzeige

def load_xml(file_path):
    """Lädt und parse die XML-Datei."""
    tree = ET.parse(file_path)
    return tree.getroot()


def load_config(file_path):
    """Lädt die Konfiguration aus einer Excel-Datei."""
    config_data = {}
    df = pd.read_excel(file_path, engine='openpyxl')
    for _, row in df.iterrows():
        tool_name = str(row['Tool']).strip()
        category = str(row['Kategorie']).strip()

        if category not in config_data:
            config_data[category] = []

        config_data[category].append({'name': tool_name, 'tags': [category]})

    return config_data


def export_config(selected_tools, file_name, destination_path):
    """Exportiert die ausgewählten Werkzeuge in eine Datei im angegebenen Verzeichnis."""
    with open(os.path.join(destination_path, file_name + '.conf'), 'w') as file:
        for tool in selected_tools:
            file.write(tool + '\n')


def update_selection(tag_selected):
    """Aktualisiert die Auswahl der Werkzeuge basierend auf dem ausgewählten Tag."""
    for category, tools in config_data.items():
        for tool in tools:
            if tag_selected in tool['tags']:
                tool_vars[tool['name']].set(True)
            else:
                tool_vars[tool['name']].set(False)


def reset_selection(tag_var):
    """Setzt die Tag-Auswahl und alle Checkboxen zurück."""
    tag_var.set("Groups")
    for var in tool_vars.values():
        var.set(False)


def toggle_selection_in_category(category):
    """Schaltet die Auswahl aller Checkboxen in der angegebenen Kategorie um."""
    all_selected = all(tool_vars[tool['name']].get() for tool in config_data[category])
    new_state = not all_selected

    for tool in config_data[category]:
        tool_vars[tool['name']].set(new_state)


def on_load_config():
    """Lädt eine Konfigurationsdatei und prüft die entsprechenden Checkboxen."""
    file_path = filedialog.askopenfilename(
        title="Konfiguration laden",
        filetypes=[("Config files", "*.conf"), ("All files", "*.*")]
    )
    
    if not file_path:
        return
    
    # Zuerst alle Checkboxen deaktivieren
    for var in tool_vars.values():
        var.set(False)
    
    # Config-Datei laden und Checkboxen entsprechend setzen
    try:
        with open(file_path, 'r') as file:
            for line in file:
                tool_name = line.strip()
                if tool_name and tool_name in tool_vars:
                    tool_vars[tool_name].set(True)
    except Exception as e:
        print(f"Fehler beim Laden der Konfiguration: {e}")


def on_export():
    """Exportiert die Konfiguration aus den ausgewählten Werkzeugen."""
    selected_tools = [tool for tool, var in tool_vars.items() if var.get()]
    destination_path = filedialog.askdirectory(title="Speicherort wählen")
    export_config(selected_tools, file_name_var.get(), destination_path)


def create_gui(config_data, xml_root):
    """Erstellt und zeigt die grafische Benutzeroberfläche an."""
    window = tk.Tk()
    window.title("DFIR Installer Selector")
    window.configure(bg=BACKGROUND_COLOR)

    window.state('zoomed')

    window.grid_columnconfigure(0, weight=1)
    window.grid_rowconfigure(0, weight=1)

    global file_name_var
    file_name_var = StringVar()

    # TextBox mit Rahmen und reduzierter Höhe
    textbox_frame = tk.Frame(window, bd=FRAME_BORDER, relief=FRAME_RELIEF, height=TEXTBOX_HEIGHT, bg=FRAME_COLOR)
    textbox_frame.grid_propagate(False)
    textbox_frame.grid(row=0, column=0, sticky='nsew', padx=LEFT_MARGIN, pady=2)

    textbox_elem = xml_root.find('TextBox')
    tk.Label(textbox_frame, text=textbox_elem.attrib['label'], font=FONT_BOLD, bg=FRAME_COLOR, fg=TEXT_COLOR).pack(
        side=tk.TOP, fill=tk.X, pady=1)
    tk.Entry(textbox_frame, textvariable=file_name_var, font=FONT, bg=BACKGROUND_COLOR, fg=TEXT_COLOR,
             insertbackground=TEXT_COLOR).pack(side=tk.TOP, fill=tk.X, padx=5, pady=1)

    # Dropdown, Reset-Button und Export-Button mit Rahmen nebeneinander
    dropdown_frame = tk.Frame(window, bd=FRAME_BORDER, relief=FRAME_RELIEF, bg=FRAME_COLOR)
    dropdown_frame.grid(row=1, column=0, sticky='nsew', padx=LEFT_MARGIN, pady=2)

    all_tags = sorted({tag for tools in config_data.values() for tool in tools for tag in tool['tags']})
    tag_var = StringVar(value="Groups")

    tag_dropdown = tk.OptionMenu(dropdown_frame, tag_var, *all_tags, command=update_selection)
    tag_dropdown.config(font=FONT, width=CHOICEBOX_WIDTH, bg=BUTTON_COLOR, fg=TEXT_COLOR,
                        highlightbackground=BUTTON_COLOR)
    tag_dropdown.pack(side=tk.LEFT, expand=True, padx=5)

    reset_button = tk.Button(dropdown_frame, text="Reset", command=lambda: reset_selection(tag_var), font=FONT,
                             bg=BUTTON_COLOR, fg=TEXT_COLOR, highlightbackground=BUTTON_COLOR)
    reset_button.pack(side=tk.LEFT, padx=5)

    load_config_button = tk.Button(dropdown_frame, text="Load Config", command=on_load_config, font=FONT,
                                   bg=BUTTON_COLOR, fg=TEXT_COLOR, highlightbackground=BUTTON_COLOR)
    load_config_button.pack(side=tk.LEFT, padx=5)

    button_elem = xml_root.find('Button')
    export_button = tk.Button(dropdown_frame, text=button_elem.attrib['label'], command=on_export, font=FONT,
                              bg=BUTTON_COLOR, fg=TEXT_COLOR, highlightbackground=BUTTON_COLOR)
    export_button.pack(side=tk.LEFT, padx=5)

    global tool_vars
    tool_vars = {}

    num_categories = len(config_data)
    canvas_width = int(num_categories * COLUMN_WIDTH * GUI_WIDTH_FACTOR)

    checkbox_frame_outer = tk.Frame(window, width=canvas_width + 100, height=SCROLLBOX_HEIGHT, bg=FRAME_COLOR)
    checkbox_frame_outer.grid(row=2, column=0, sticky='nsew', padx=LEFT_MARGIN, pady=5)
    checkbox_frame_outer.pack_propagate(False)

    canvas = tk.Canvas(checkbox_frame_outer, bg=FRAME_COLOR)
    horizontal_scrollbar = tk.Scrollbar(checkbox_frame_outer, orient="horizontal", command=canvas.xview)
    vertical_scrollbar = tk.Scrollbar(checkbox_frame_outer, orient="vertical", command=canvas.yview)

    canvas.configure(xscrollcommand=horizontal_scrollbar.set, yscrollcommand=vertical_scrollbar.set)

    scrollable_frame = tk.Frame(canvas, bg=FRAME_COLOR)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.pack(side="left", fill="both", expand=True)
    horizontal_scrollbar.pack(side="bottom", fill="x")
    vertical_scrollbar.pack(side="right", fill="y")

    def _on_mouse_wheel(event):
        if event.state & 0x0004:
            canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

    column_count = 0
    for category, tools in config_data.items():
        category_label = tk.Label(scrollable_frame, text=category, font=FONT_BOLD, bg=FRAME_COLOR, fg=TEXT_COLOR)
        category_label.grid(row=0, column=column_count, sticky=ALIGNMENT_CENTER, padx=LEFT_MARGIN, pady=5)

        category_label.bind("<Button-1>", lambda e, cat=category: toggle_selection_in_category(cat))

        for i, tool in enumerate(tools):
            tool_vars[tool['name']] = tk.BooleanVar()
            tk.Checkbutton(
                scrollable_frame, text=tool['name'], variable=tool_vars[tool['name']], font=FONT_SMALL,
                bg=FRAME_COLOR, fg=TEXT_COLOR, selectcolor=BUTTON_COLOR
            ).grid(row=i + 1, column=column_count, sticky=ALIGNMENT_LEFT, padx=LEFT_MARGIN)

        column_count += 1

    def scroll_left():
        canvas.xview_scroll(-1, "units")

    def scroll_right():
        canvas.xview_scroll(1, "units")

    button_left = tk.Button(window, text="<", command=scroll_left, font=FONT, bg=BUTTON_COLOR, fg=TEXT_COLOR,
                            highlightbackground=BUTTON_COLOR)
    button_left.grid(row=3, column=0, sticky='w', padx=5)

    button_right = tk.Button(window, text=">", command=scroll_right, font=FONT, bg=BUTTON_COLOR, fg=TEXT_COLOR,
                             highlightbackground=BUTTON_COLOR)
    button_right.grid(row=3, column=0, sticky='e', padx=5)

    window.state('zoomed')

    window.mainloop()


xml_file_path = os.path.join(os.path.dirname(__file__), 'gui_template.xml')
xml_root = load_xml(xml_file_path)

config_file_path = os.path.join(os.path.dirname(__file__), 'dfir-installer-selector.xlsx')
config_data = load_config(config_file_path)

create_gui(config_data, xml_root)
