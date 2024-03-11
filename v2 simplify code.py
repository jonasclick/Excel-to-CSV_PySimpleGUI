#I made this following a tutorial https://www.youtube.com/watch?v=LzCfNanQ_9c
from pathlib import Path

import pandas as pd
import PySimpleGUI as sg


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct")
    return False

def display_excel_file(excel_file_path, sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name) #convert excel into pandas data frame (df)
    filename = Path(excel_file_path).stem   #.stem keeps only the filename from the whole path
    sg.popup_scrolled(df.dtypes, "=" * 25, df, title=filename)

def convert_to_csv(excel_file_path, output_folder, sheet_name, separator, decimal):
    df = pd.read_excel(excel_file_path, sheet_name) #convert excel into pandas data frame (df)
    filename = Path(excel_file_path).stem   #.stem keeps only the filename from the whole path
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, sep=separator, decimal=decimal, index=False) #df.to_csv converts dataframe into a csv file 
    sg.popup_no_titlebar("Your file has been converted.")

def settings_window(settings):
    # – – – – – – GUI Definition – – – – – – #
    layout = [[sg.T("Adjust your preferred behaviour for the CSV file.")],
              [sg.T("Separator"), sg.I(settings["CSV"]["separator"], s=1, key="-SEPARATOR-"),
               sg.T("Decimal"), sg.Combo(settings["CSV"]["decimal"].split("|"), default_value=settings["CSV"]["decimal_default"], s=1, key="-DECIMAL-"),
                sg.T("Sheet Name"), sg.I(settings["EXCEL"]["sheet_name"], s=20, key="-SHEET_NAME-")],
                [sg.B("Save Settings", s=16)]]
    window = sg.Window("Settings Window", layout, modal=True, use_custom_titlebar=True) #modal: blocks all other windows of the app until you close this one.
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "Save Settings":
            # Write to ini file
            settings["CSV"]["separator"] = values["-SEPARATOR-"] #overwrites current settings. Just reassign the value basically. Easy.
            settings["CSV"]["decimal_default"] = values["-DECIMAL-"]
            settings["EXCEL"]["sheet_name"] = values["-SHEET_NAME-"]

            # Display success message & close window
            sg.popup_no_titlebar("Settings saved!")
            break
    window.close()


def main_window():
    # - - - - - - Menu Definition - - - - - - #
    menu_def = [["Toolbar", ["Command 1", "Command 2", "---", "Command 3", "Command 4"]], #Devider is detected when "---"
                ["Help", ["Settings", "About", "Exit"]]]
    
    # - - - - - - GUI Definition - - - - - - #
    layout = [[sg.MenubarCustom(menu_def, tearoff=False)], #tearoff=true lets you drag the extended menu list around like a window
        [sg.Text("Select Input File", s=15, justification="r"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))], #key is the key assigned in the dictionary created from the input
        [sg.Text("Select Output Folder", s=15, justification="r"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
        [sg.Button("Display Excel File", s=16), sg.Button("Convert To CSV", s=16, button_color="green"), sg.B("Settings", s=16, button_color="orange"), sg.Exit(s=16, button_color="teal")],
    ]

    # Run the GUI Window
    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        print(event, values) #can be removed
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "About":
            window.disappear()
            sg.popup(window_title, "Version 1.0", "Convert Excel Files to CSV", grab_anywhere=True, modal=True) #when I try to exit out of the popup it seems to freeze until I come back to the popup.
            window.reappear()
        if event in ("Command 1", "Command 2", "Command 3", "Command 4"):
            sg.popup_error("This feature is not yet implemented.", modal=True)
        if event == "Settings":
            settings_window(settings)
        if event == "Display Excel File":
            if is_valid_path(values["-IN-"]):
                display_excel_file(values["-IN-"], settings["EXCEL"]["sheet_name"])
        if event == "Convert To CSV":
            if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):
                convert_to_csv(
                    excel_file_path=values["-IN-"],
                    output_folder=values["-OUT-"],
                    sheet_name=settings["EXCEL"]["sheet_name"],
                    separator=settings["CSV"]["separator"],
                    decimal=settings["CSV"]["decimal"],
                )
    window.close()


if __name__ == "__main__": #this can be run from the commandline? __name__ is a built in variable in Python? if run, it executes the code after this line
    # SETTINGS_PATH = Path.cwd() #cwd = current working directory, returned a directory too high for me!
    SETTINGS_PATH = "/Users/jonas/JONAS VETSCH/2_PRIVAT/5_Projekte/Coding Projects 2024/GPT Coding Interface/XLS to CSV with GUI"
    #create the settings object and use ini format
    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename="config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"] #1st access a section, 2nd get the value for theme
    font_family = settings["GUI"]["font_family"]
    font_size = int(settings["GUI"]["font_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()