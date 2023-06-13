import json
from Modules import PowerBIReport, Vectorise
from datetime import datetime
import os
import zipfile
import shutil
import cairosvg
from pypdf import PdfWriter


INVALID_CHARACTERS = {'<', '>', ':', '/', '\\', '|', '?', '*', '(', ')'}


def convert_to_pdf(svg_directory, output_name):
    svg_files = os.listdir(svg_directory)
    file_count = len(svg_files)
    current_count = 1
    merged_pdf = PdfWriter()
    for file in svg_files:
        name = file[:-4]
        file_path = f'{svg_directory}/{file}'
        temp_pdf = f'{svg_directory}/_temp_{name}.pdf'
        cairosvg.svg2pdf(url=file_path, write_to=temp_pdf)
        merged_pdf.append(temp_pdf)
        os.remove(temp_pdf)
        print(f'{current_count} of {file_count} PDF files merged.')
        current_count += 1

    merged_pdf.write(f'{svg_directory}/{output_name}.pdf')
    merged_pdf.close()
    print('Merged PDF saved.')


def explode_pbix(pbix_location, save_location, wireframe_settings):
    report_file_name = os.path.basename(pbix_location)
    report_name = report_file_name[:report_file_name.index('.')]
    report_directory = os.path.dirname(pbix_location)
    time = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')
    merge_images = wireframe_settings['merge_canvas_background']
    create_pdf = wireframe_settings['create_pdf']

    if save_location is not None:
        save_directory = save_location
    else:
        save_directory = report_directory

    # Create output folders
    output_folder = f"{save_directory}/{report_name} Documentation {time}"
    subfolder_list = ['Formatted Layout', 'Raw Layout', 'Exploded', 'Page Wireframe']
    for folder in subfolder_list:
        os.makedirs(f"{output_folder}/{folder}")

    # Create report backup
    shutil.copy2(pbix_location, output_folder)
    print("Created report backup")

    # Extract pbix
    with zipfile.ZipFile(pbix_location, 'r') as archive:
        save_location_exploded = os.path.abspath(f"{output_folder}/Exploded")
        archive.extractall(save_location_exploded)
    print("Extracted report files")

    # Load pbix layout file
    pbix = PowerBIReport.PowerBIReport(pbix_location)
    page_count = len(pbix.layout['sections'])
    current_page = 1

    # Extract report sections
    for page in pbix.layout['sections']:
        # Clean page display name, so it can be used as a valid file name
        sanitised_name = page['displayName'].translate({ord(x): ' ' for x in INVALID_CHARACTERS}).strip()
        file_name = f'{current_page}-{sanitised_name}.json'
        save_location_formatted = os.path.abspath(f"{output_folder}/Formatted Layout/{file_name}")
        save_location_raw = os.path.abspath(f"{output_folder}/Raw Layout/{file_name}")

        # Save raw section export
        with open(save_location_raw, 'w') as f:
            json.dump(page, f)

        # Expand config section json
        parsed_json = json.loads(page['config'])
        page['config'] = parsed_json
        # Expand filters section json
        parsed_json = json.loads(page['filters'])
        page['filters'] = parsed_json

        # Save pretty printed section
        with open(save_location_formatted, 'w') as f:
            json.dump(page, f, indent=4)
        print(f"Parsed page {current_page} of {page_count}\n")
        current_page += 1

        # Generate a report page preview from pretty printed section json
        Vectorise.generate_report_preview(output_folder, file_name, merge_images)

    print("All pages parsed.")

    if create_pdf:
        svg_location = f'{output_folder}/Page Wireframe'
        convert_to_pdf(svg_location, report_name)
