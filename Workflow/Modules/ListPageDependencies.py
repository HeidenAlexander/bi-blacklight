import os
import datetime
from Modules import Report


def list_page_dependencies(file_path, save_location, create_subfolder):
    report_file_name = os.path.basename(file_path)
    report_name = report_file_name[:report_file_name.index(".")]
    report_directory = os.path.dirname(file_path)
    time = datetime.now().strftime('%Y-%m-%dT%H-%M-%S')

    if save_location is not None and create_subfolder:
        save_directory = f"{save_location}/{report_name} Page Dependencies {time}"
        os.makedirs(save_directory)
    elif save_location is None and create_subfolder:
        save_directory = f"{report_directory}/{report_name} Page Dependencies {time}"
        os.makedirs(save_directory)
    elif save_location is not None:
        save_directory = save_location
    else:
        save_directory = report_directory

    pbix = Report.PowerBIReport(file_path)
    pbix.export_page_info(report_name, save_directory)


def list_parents(save_location, layout_json, svg_width, svg_height, has_canvas_image, background_extension):
    # Create a list of visuals and groups
    visual_hierarchy = []
    visual_count = 0
    group_count = 0
    for visual in layout_json['visualContainers']:
        vis_parent = {
            'category': None,
            'x': visual['x'],
            'y': visual['y'],
            'width': visual['width'],
            'height': visual['height'],
            'parent': None,
            'name': None
        }
        if 'singleVisual' in visual['config']:
            visual_count += 1
            vis_parent['category'] = 'singleVisual'
            vis_parent['name'] = visual['config']['singleVisual']['visualType']
            if 'parentGroupName' in visual['config']:
                vis_parent['parent'] = visual['config']['parentGroupName']
        elif 'singleVisualGroup' in visual['config']:
            group_count += 1
            vis_parent['category'] = 'visualGroup'
            vis_parent['name'] = visual['config']['singleVisualGroup']['displayName']
            if 'parentGroupName' in visual['config']:
                vis_parent['parent'] = visual['config']['parentGroupName']

        parent = vis_parent['parent']
        while parent is not None:
            for vis_name in layout_json['visualContainers']:
                if parent == vis_name['config']['name']:
                    vis_parent['x'] += vis_name['x']
                    vis_parent['y'] += vis_name['y']
                    if 'parentGroupName' in vis_name['config']:
                        parent = vis_name['config']['parentGroupName']
                        break
                    else:
                        parent = None
                        break
                else:
                    continue

        visual_hierarchy.append(vis_parent)
