import json
from Modules import PowerBIReport, Vectorise
from datetime import datetime
import os
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
        print(f'{current_count} of {file_count} SVG files merged into PDF.')
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

    # Create output folder
    output_folder = f"{save_directory}/{report_name} Wireframes {time}"
    os.makedirs(output_folder)

    # Load pbix layout file
    pbix = PowerBIReport.PowerBIReport(pbix_location)
    current_page = 1

    # Extract report sections
    for page in pbix.layout['sections']:
        # Remove invalid filename characters
        revised_name = page['displayName'].translate({ord(x): ' ' for x in INVALID_CHARACTERS}).strip()
        name = f'{current_page}-{revised_name}'
        current_page += 1

        # Expand config section json
        parsed_json = json.loads(page['config'])
        page['config'] = parsed_json
        # Expand filters section json
        parsed_json = json.loads(page['filters'])
        page['filters'] = parsed_json

        # Assign background filename and location if exists
        try:
            page_background = (page['config']['objects']['background'][0]['properties']
                               ['image']['image']['url']['expr']['ResourcePackageItem']['ItemName'])
            has_background = True
        except KeyError:
            page_background = None
            has_background = False

        if has_background:
            background_image = pbix.binary.extract(f'Report/StaticResources/RegisteredResources/{page_background}')
            background_type = os.path.splitext(page_background)[1]
        else:
            background_image = None
            background_type = None

        # Generate a report page wireframe
        Vectorise.generate_wireframe(output_folder, page, name, has_background, background_image,
                                     background_type, merge_images)

    print("All pages parsed.")

    if create_pdf:
        svg_location = output_folder
        convert_to_pdf(svg_location, report_name)
