import sys
import tkinter
from tkinter import messagebox, filedialog as fd
import customtkinter
import threading
from Modules import SplitPBIX, FieldReference, ListPageNames, PowerShell, Explode, ImportSettings

SETTINGS = ImportSettings.import_script_settings('config.yaml')


def redirect_output(input_string):
    output.insert(tkinter.END, input_string)
    output.see('end')


def process_thread(func, arguments: list):
    func = threading.Thread(target=func, args=arguments)
    sys.stdout.write = redirect_output
    func.start()


def completion_dialogue():
    messagebox.showinfo('Process Completed', 'Process completed successfully.', )


def get_workspace_id():
    entry_screen = customtkinter.CTkInputDialog(text='Enter Workspace ID', title='Workspace ID')
    root.eval(f'tk::PlaceWindow {str(entry_screen)} center')
    workspace_id = entry_screen.get_input()
    return workspace_id


def workspace_id_menu(workspace_list):
    input_window = customtkinter.CTkToplevel(root)
    input_window.title('Select Workspace')
    input_window.geometry('500x230')
    input_window.grid_columnconfigure(0, weight=1)
    input_window.grid_rowconfigure(0, weight=1)
    root.eval(f'tk::PlaceWindow {str(input_window)} center')

    # Tabs
    tabs = customtkinter.CTkTabview(master=input_window, width=380)
    tabs.grid(column=0, row=0, padx=3, pady=3, sticky='nsew')

    # First Tab
    tabs.add('Workspace Selection')
    tabs.tab('Workspace Selection').grid_columnconfigure(0, weight=1)

    option_list = customtkinter.CTkOptionMenu(
        master=tabs.tab('Workspace Selection'),
        width=450,
        dynamic_resizing=True,
        values=workspace_list

    )
    option_list.grid(row=0, column=0, pady=(25, 0))

    selection_accept_button = customtkinter.CTkButton(
        master=tabs.tab('Workspace Selection'),
        text='Accept and Continue',
        font=('Segoe UI', 14),
        text_color=('#000000', '#FFFFFF'),
        corner_radius=5,
        command=input_window.destroy
    )
    selection_accept_button.grid(row=1, column=0, pady=50)

    # Second Tab
    tabs.add('Manual Entry')
    tabs.tab('Manual Entry').grid_columnconfigure(0, weight=1)

    man_entry = tkinter.StringVar()
    manual_entry = customtkinter.CTkEntry(
        master=tabs.tab('Manual Entry'),
        placeholder_text='Workspace ID',
        width=450,
        textvariable=man_entry
    )
    manual_entry.grid(row=0, column=0, pady=(25, 0))

    manual_entry_warning_label = customtkinter.CTkLabel(
        master=tabs.tab('Manual Entry'),
        text='Any value entered here will override a selection in the previous screen',
        font=('Segoe UI', 12)
    )
    manual_entry_warning_label.grid(row=1, column=0, pady=3)

    manual_accept_button = customtkinter.CTkButton(
        master=tabs.tab('Manual Entry'),
        text='Accept and Continue',
        font=('Segoe UI', 14),
        text_color=('#000000', '#FFFFFF'),
        corner_radius=5,
        command=input_window.destroy
    )
    manual_accept_button.grid(row=2, column=0, pady=15)

    input_window.wait_window()
    return man_entry.get(), option_list.get()


def divide_pbix_file(settings):
    filetypes = (('Power BI', '*.pbix'),)
    pbix_name = fd.askopenfilename(title='Select the Power BI source file (*.pbix)', filetypes=filetypes)
    if pbix_name == '':
        return
    filetypes = (('Excel', '*.xlsx'),)
    mapping_name = fd.askopenfilename(title='Select report mapping .xlsx file', filetypes=filetypes)
    if mapping_name == '':
        return
    if settings['save_options']['split_pbix']['ask_save_location']:
        save_directory = fd.askdirectory(title='Select Output Location', )
        if save_directory == '':
            return
    else:
        save_directory = None
    subfolder = settings['save_options']['split_pbix']['create_subfolder']
    SplitPBIX.split_pbix(pbix_name, mapping_name, save_directory, subfolder)
    completion_dialogue()


def remap_pbix_table(settings):
    filetypes = (('Power BI', '*.pbix'),)
    pbix_name = fd.askopenfilename(title='Select the Power BI source file (*.pbix)', filetypes=filetypes)
    if pbix_name == '':
        return
    filetypes = (('Excel', '*.xlsx'),)
    table_mapping_name = fd.askopenfilename(title='Select report table mapping .xlsx file', filetypes=filetypes)
    if table_mapping_name == '':
        return
    if settings['save_options']['field_remapping']['ask_save_location']:
        save_directory = fd.askdirectory(title='Select Output Location', )
        if save_directory == '':
            return
    else:
        save_directory = None
    subfolder = settings['save_options']['field_remapping']['create_subfolder']
    FieldReference.rename_tables(pbix_name, table_mapping_name, save_directory, subfolder)
    completion_dialogue()


def remap_pbix_file(settings):
    filetypes = (('Power BI', '*.pbix'),)
    pbix_name = fd.askopenfilename(title='Select the Power BI source file (*.pbix)', filetypes=filetypes)
    if pbix_name == '':
        return
    filetypes = (('Excel', '*.xlsx'),)
    field_mapping_name = fd.askopenfilename(title='Select report field mapping .xlsx file', filetypes=filetypes)
    if field_mapping_name == '':
        return
    if settings['save_options']['field_remapping']['ask_save_location']:
        save_directory = fd.askdirectory(title='Select Output Location', )
        if save_directory == '':
            return
    else:
        save_directory = None
    subfolder = settings['save_options']['field_remapping']['create_subfolder']
    FieldReference.rename_fields(pbix_name, field_mapping_name, save_directory, subfolder)
    completion_dialogue()


def create_report_page_file(settings):
    filetypes = (('Power BI', '*.pbix'),)
    filename = fd.askopenfilename(title='Select the Power BI source file (*.pbix)', filetypes=filetypes)
    if filename == '':
        return
    if settings['save_options']['page_list']['ask_save_location']:
        save_directory = fd.askdirectory(title='Select Output Location', )
        if save_directory == '':
            return
    else:
        save_directory = None
    subfolder = settings['save_options']['page_list']['create_subfolder']
    ListPageNames.report_info(filename, save_directory, subfolder)
    completion_dialogue()


def create_report_documentation(settings):
    filetypes = (('Power BI', '*.pbix'),)
    filename = fd.askopenfilename(title='Select the Power BI source file (*.pbix)', filetypes=filetypes)
    if filename == '':
        return
    if settings['save_options']['report_documentation']['ask_save_location']:
        save_directory = fd.askdirectory(title='Select Output Location', )
        if save_directory == '':
            return
    else:
        save_directory = None
    wireframe_settings = settings['page_wireframe_options']
    Explode.explode_pbix(filename, save_directory, wireframe_settings)
    completion_dialogue()


def publish_multiple(*args):
    directory = fd.askdirectory(title='Select Folder Containing .pbix Files', )
    if directory == '':
        return
    workspace_list = list(SETTINGS['workspaces'].keys())
    workspace_id = workspace_id_menu(workspace_list)
    if workspace_id == "":
        return
    elif workspace_id[0] != "":
        process_thread(PowerShell.publish_single_report, [filename, workspace_id[0]])
    else:
        wid = SETTINGS['workspaces'][workspace_id[1]]
        process_thread(PowerShell.publish_single_report, [filename, wid])


def publish_single(*args):
    filetypes = (('Power BI', '*.pbix'),)
    filename = fd.askopenfilename(title='Select the Power BI source file (*.pbix)', filetypes=filetypes)
    if filename == "":
        return
    workspace_list = list(SETTINGS['workspaces'].keys())
    workspace_id = workspace_id_menu(workspace_list)
    if workspace_id == "":
        return
    elif workspace_id[0] != "":
        process_thread(PowerShell.publish_single_report, [filename, workspace_id[0]])
    else:
        wid = SETTINGS['workspaces'][workspace_id[1]]
        process_thread(PowerShell.publish_single_report, [filename, wid])


# Setup GUI
customtkinter.set_appearance_mode('System')
customtkinter.set_default_color_theme('blue')

# create the root window
root = customtkinter.CTk()
root.title('BI Blacklight')
root.resizable(True, True)
root.geometry('720x480')
root.minsize(500, 230)
root.grid_columnconfigure(0, weight=0)
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# Create actions frame
action_frame = customtkinter.CTkFrame(root, width=240, corner_radius=0)
action_frame.grid(column=0, row=0, sticky='nsew')
action_frame.grid_rowconfigure('all', weight=1)

action_label = customtkinter.CTkLabel(
    action_frame,
    text='Utilities',
    text_color=('#000000', '#FFFFFF'),
    font=('Segoe UI SemiBold', 18),
    anchor='w'
)
action_label.grid(column=0, row=0, padx=15, sticky='w')

# Create output frame
output_frame = customtkinter.CTkFrame(root, corner_radius=0)
output_frame.grid(column=1, row=0, padx=(2, 0), sticky='nsew')
output_frame.grid_columnconfigure(0, weight=1)
output_frame.grid_rowconfigure(0, weight=1)

output = customtkinter.CTkTextbox(output_frame, wrap='none', font=('Consolas', 12))
output.configure(spacing2=1)
output.grid(column=0, padx=3, pady=3, sticky='nsew')

# report info button
report_info_button = customtkinter.CTkButton(
    master=action_frame,
    text='Create Report Page List',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=lambda: process_thread(create_report_page_file, [SETTINGS])
)
report_info_button.grid(column=0, row=1, padx=10, pady=(10, 2), sticky='w')

# split report button
split_report_button = customtkinter.CTkButton(
    master=action_frame,
    text='Split Reports',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=lambda: process_thread(divide_pbix_file, [SETTINGS])
)
split_report_button.grid(column=0, row=2, padx=10, pady=(2, 10), sticky='w')

# remap report tables button
remap_report_table_button = customtkinter.CTkButton(
    master=action_frame,
    text='Remap Report Tables (Alpha)',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=lambda: process_thread(remap_pbix_table, [SETTINGS])
)
remap_report_table_button.grid(column=0, row=3, padx=10, pady=2, sticky='w')

# remap report button
remap_report_button = customtkinter.CTkButton(
    master=action_frame,
    text='Remap Report Fields (Alpha)',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=lambda: process_thread(remap_pbix_file, [SETTINGS])
)
remap_report_button.grid(column=0, row=4, padx=10, pady=(2, 10), sticky='w')

# Explode report button
explode_report_button = customtkinter.CTkButton(
    master=action_frame,
    text='Create Report Documentation',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=lambda: process_thread(create_report_documentation, [SETTINGS])
)
explode_report_button.grid(column=0, row=5, padx=10, pady=(2, 10), sticky='w')

# publish single report button
publish_single_report_button = customtkinter.CTkButton(
    master=action_frame,
    text='Publish Report',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=publish_single
)
publish_single_report_button.grid(column=0, row=6, padx=10, pady=2, sticky='w')

# publish multiple reports button
publish_multiple_reports_button = customtkinter.CTkButton(
    master=action_frame,
    text='Publish Multiple Reports',
    width=230,
    font=('Segoe UI', 16),
    text_color=('#000000', '#FFFFFF'),
    anchor='w',
    corner_radius=5,
    fg_color='transparent',
    command=publish_multiple
)
publish_multiple_reports_button.grid(column=0, row=7, padx=10, pady=2, sticky='w')

# settings button
# settings_button = customtkinter.CTkButton(
#     master=action_frame,
#     text='Settings',
#     width=70,
#     font=('Segoe UI', 16),
#     text_color=('#242424', '#f2f2f2'),
#     anchor='w',
#     corner_radius=5,
#     fg_color='transparent',
#     # command=publish_multiple
# )
# settings_button.grid(column=0, row=7, padx=10, pady=10, sticky='w')

# run the application
root.eval('tk::PlaceWindow . center')

root.mainloop()
