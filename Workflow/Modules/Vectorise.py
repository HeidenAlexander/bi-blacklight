import json
import os
import cairo
from svgutils.compose import Figure, SVG
import svgutils.transform as sg
from Modules import ColourDictionary


def create_svg_layout(save_location, layout_json, svg_width, svg_height, has_canvas_image, background_extension):
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

    # Define canvas size and draw border
    with cairo.SVGSurface(save_location, svg_width, svg_height) as surface:
        context = cairo.Context(surface)
        if has_canvas_image and background_extension == '.svg':
            context.set_source_rgba(1, 1, 1, 0)
        else:
            context.set_source_rgba(1, 1, 1, 1)

        context.set_line_width(1)
        context.rectangle(0, 0, svg_width, svg_height)
        context.fill_preserve()
        context.set_source_rgb(0, 0, 0)
        context.stroke()

        # Begin drawing all shapes corresponding to report page visuals
        for visual in visual_hierarchy:
            viz_name = visual['name']
            viz_category = visual['category']
            viz_x = visual['x']
            viz_y = visual['y']
            visual_width = visual['width']
            visual_height = visual['height']

            # Set colour and alpha
            if viz_name in ColourDictionary.visual_colours and viz_category == 'singleVisual':
                r = ColourDictionary.visual_colours[viz_name][0]
                g = ColourDictionary.visual_colours[viz_name][1]
                b = ColourDictionary.visual_colours[viz_name][2]
                a = .2
                text_y_offset = 0
                context.set_source_rgba(r, g, b, a)
                context.rectangle(viz_x, viz_y, visual_width, visual_height)
                context.fill()

            elif viz_category == 'singleVisual':
                r = 1
                g = 0
                b = 0
                a = .2
                text_y_offset = 0
                context.set_source_rgba(r, g, b, a)
                context.rectangle(viz_x, viz_y, visual_width, visual_height)
                context.fill()

            elif viz_category == 'visualGroup':
                r = 0
                g = 0
                b = 0
                a = .5
                text_y_offset = 12
                context.set_line_width(1)
                context.set_dash([10, 5])
                context.set_source_rgba(r, g, b, a)
                context.rectangle(viz_x, viz_y, visual_width, visual_height)
                context.stroke()

            # Annotate shape
            font_size = 10
            context.set_source_rgba(0, 0, 0, .5)
            xbearing, ybearing, width, height, dx, dy = context.text_extents(viz_name)
            context.select_font_face("Bahnschrift", cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_NORMAL)
            context.set_font_size(font_size)
            x_loc = (viz_x + visual_width / 2) - width / 2
            y_loc = (viz_y + visual_height/2) + font_size/2 + text_y_offset

            if viz_category == 'visualGroup':
                x_loc = viz_x + 5
                y_loc = viz_y + font_size + 5
                context.select_font_face("Bahnschrift", cairo.FONT_SLANT_ITALIC, cairo.FONT_WEIGHT_NORMAL)

            if x_loc < 0:
                context.move_to(viz_x + 5, y_loc)
            else:
                context.move_to(x_loc, y_loc)

            context.show_text(viz_name)

        print("Layout SVG file saved.")
    return visual_count, group_count


def merge_svg(svg_1, svg_2, save_location):
    # Obtain svg dimensions and scale difference
    canvas_background = sg.fromfile(svg_1)
    visual_diagram = sg.fromfile(svg_2)

    vis_width = float(visual_diagram.width)
    vis_height = float(visual_diagram.height)
    canvas_width = float(canvas_background.width)
    canvas_height = float(canvas_background.height)

    x_scale_ratio = vis_width / canvas_width
    y_scale_ratio = vis_height / canvas_height

    # Merge and save
    Figure(
        vis_width,
        vis_height,
        SVG(svg_1).scale(x_scale_ratio, y_scale_ratio),
        SVG(svg_2)
    ).save(save_location)

    print("Merged layout SVG saved.")


def generate_wireframe(folder_path, layout_file, page_name, has_background, background, background_type, merge):
    save_path = folder_path
    json_file = layout_file
    filename = page_name
    page_width = json_file['width']
    page_height = json_file['height']

    # Configure initial output depending on if the report page has a canvas background
    if has_background and merge and background_type == '.svg':
        layout_svg = f"_temp_{filename}.svg"
    else:
        layout_svg = f"{filename}.svg"

    svg_file_path = os.path.abspath(f'{save_path}/{layout_svg}')

    # Create report page layout and return visual and group counts
    visual_count, group_count = create_svg_layout(svg_file_path, json_file,
                                                  page_width, page_height, has_background, background_type)

    # If report page has a svg canvas background and config is set to true, merge with the recently created
    # svg page layout & delete temp svg
    if has_background and merge and background_type == '.svg':
        merged_save_path = f'{save_path}/{filename}.svg'
        merge_svg(background, svg_file_path, merged_save_path)
        os.remove(svg_file_path)

    if has_background:
        print(f"The {filename} report page is {page_width} x {page_height}px. Contains {visual_count} visuals,"
              f" {group_count} groups and uses a canvas background")
    else:
        print(f"The {filename} report page is {page_width} x {page_height}px. Contains {visual_count} visuals,"
              f" {group_count} groups and has no canvas background")

# To do: Add non-svg canvas background handling - currently will ignore all non svg's for merge operation.
