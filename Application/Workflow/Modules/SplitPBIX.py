from Modules import PowerBIReport, ImportSettings
import os
from datetime import datetime


def split_pbix(pbix_location, mapping_file, save_location, create_subfolder):
    reports = ImportSettings.import_report_settings(mapping_file)
    report_file_name = os.path.basename(pbix_location)
    report_name = report_file_name[:report_file_name.index(".")]
    report_count = len(reports)
    report_directory = os.path.dirname(pbix_location)
    time = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')

    if save_location is not None and create_subfolder:
        save_directory = f"{save_location}/{report_name} Subreports {time}"
        os.makedirs(save_directory)
    elif save_location is None and create_subfolder:
        save_directory = f"{report_directory}/{report_name} Subreports {time}"
        os.makedirs(save_directory)
    elif save_location is not None:
        save_directory = save_location
    else:
        save_directory = report_directory

    current_report = 1
    for sub_report in reports:
        pbix = PowerBIReport.PowerBIReport(pbix_location)
        pbix.remove_other_pages(sub_report['Pages'], sub_report['Page_Name'])
        pbix.remove_bookmarks(sub_report['Pages'])
        save_path = os.path.join(save_directory, f"{sub_report['Name']}.pbix")
        pbix.create_report(save_path)
        print(f"{current_report} of {report_count} reports processed.")
        current_report += 1
    print("All reports created.")