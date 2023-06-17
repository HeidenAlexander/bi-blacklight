from Modules import PowerBIReport
import os
from datetime import datetime


def report_info(file_path, save_location, create_subfolder):
    report_file_name = os.path.basename(file_path)
    report_name = report_file_name[:report_file_name.index(".")]
    report_directory = os.path.dirname(file_path)
    time = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')

    if save_location is not None and create_subfolder:
        save_directory = f"{save_location}/{report_name} Page List {time}"
        os.makedirs(save_directory)
    elif save_location is None and create_subfolder:
        save_directory = f"{report_directory}/{report_name} Page List {time}"
        os.makedirs(save_directory)
    elif save_location is not None:
        save_directory = save_location
    else:
        save_directory = report_directory

    pbix = PowerBIReport.PowerBIReport(file_path)
    pbix.export_page_info(report_name, save_directory)
