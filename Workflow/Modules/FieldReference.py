from Modules import ImportSettings, PowerBIReport, PowerBIReportVisual
import os
from datetime import datetime


def rename_tables(pbix_location, table_mapping_file, save_location, create_subfolder):
    table_list = ImportSettings.import_table_mapping(table_mapping_file)
    report_file_name = os.path.basename(pbix_location)
    report_name = report_file_name[:report_file_name.index(".")]
    report_directory = os.path.dirname(pbix_location)
    time = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')

    # Set save behaviour based on config
    if save_location is not None and create_subfolder:
        save_directory = f"{save_location}/{report_name} remapped {time}"
        os.makedirs(save_directory)
    elif save_location is None and create_subfolder:
        save_directory = f"{report_directory}/{report_name} remapped {time}"
        os.makedirs(save_directory)
    elif save_location is not None:
        save_directory = save_location
    else:
        save_directory = report_directory

    pbix = Report.PowerBIReport(pbix_location)
    pbix.rename_table(table_list)
    save_path = os.path.join(save_directory, f"{report_name}.pbix")
    pbix.create_report(save_path)


def rename_fields(pbix_location, field_mapping_file, save_location, create_subfolder):
    field_list = ImportSettings.import_field_mapping(field_mapping_file)
    report_file_name = os.path.basename(pbix_location)
    report_name = report_file_name[:report_file_name.index(".")]
    report_directory = os.path.dirname(pbix_location)
    time = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')

    # Set save behaviour based on config
    if save_location is not None and create_subfolder:
        save_directory = f"{save_location}/{report_name} remapped {time}"
        os.makedirs(save_directory)
    elif save_location is None and create_subfolder:
        save_directory = f"{report_directory}/{report_name} remapped {time}"
        os.makedirs(save_directory)
    elif save_location is not None:
        save_directory = save_location
    else:
        save_directory = report_directory

    # Iterate through fields to be renamed
    pbix = Report.PowerBIReport(pbix_location)
    pbix.repoint_field(field_list)
    save_path = os.path.join(save_directory, f"{report_name}.pbix")
    pbix.create_report(save_path)
    print("All fields remapped.")
