import pandas as pd
import yaml


def import_report_settings(settings_file):
    wb = pd.ExcelFile(settings_file)
    all_sheets = wb.sheet_names

    # Remove template and system Excel sheets
    removal_list = ['_SYSTEM', '_TEMPLATE']
    for sheet in removal_list:
        sys_sheet_index = all_sheets.index(sheet)
        del all_sheets[sys_sheet_index]

    reports = []
    for sheet in all_sheets:
        df = pd.read_excel(settings_file, sheet, header=None)
        sub_report = {'Name': df.iloc[0, 0]}
        df = df.iloc[3:]
        sub_report['Pages'] = df[2].values.tolist()
        sub_report['Page_Name'] = {}
        for page_id in df.values:
            if pd.isnull(page_id[1]):
                continue
            else:
                sub_report['Page_Name'][page_id[2]] = page_id[1]
        reports.append(sub_report)
    print(reports)
    return reports


def import_table_mapping(table_mapping_file):
    wb = pd.ExcelFile(table_mapping_file)
    sheet_name = wb.sheet_names
    df = pd.read_excel(table_mapping_file, sheet_name[0], header=None)
    df = df.iloc[1:]
    mappings = []
    for row in df.values:
        table_reference = []
        table_reference.append(row[0])
        table_reference.append(row[1])
        mappings.append(table_reference)
    return mappings


def import_field_mapping(field_mapping_file):
    wb = pd.ExcelFile(field_mapping_file)
    sheet_name = wb.sheet_names
    df = pd.read_excel(field_mapping_file, sheet_name[0], header=None)
    df = df.iloc[1:]
    mappings = []
    for row in df.values:
        field_reference = []
        field_reference.append(row[0])
        field_reference.append(row[1])
        field_reference.append(row[2])
        field_reference.append(row[3])
        mappings.append(field_reference)
    return mappings


def import_script_settings(config_yaml):
    stream = open(config_yaml, 'r')
    settings = yaml.safe_load(stream)
    return settings
