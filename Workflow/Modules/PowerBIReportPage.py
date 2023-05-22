import copy
import json
import re
from Modules import PowerBIReport, PowerBIReportVisual


class PowerBIReportPage:
    def __init__(self, parent: PowerBIReport, page_json):
        self.parent = parent
        self.page_json = page_json
        self.visuals = []
        for visual in page_json['visualContainers']:
            self.visuals.append(PowerBIReportVisual.PowerBIReportVisual(self, visual))

    def rename_page(self, new_name: str):
        self.page_json['displayName'] = new_name

    def get_page_json(self):
        page_json = []
        for visual in self.visuals:
            page_json.append(visual.get_visual_json())
        self.page_json['visualContainers'] = page_json
        return copy.deepcopy(self.page_json)

    def repoint_field(self, field_mapping_file):
        name = self.page_json['displayName']
        self.visuals = []  # Clear visual list
        page_counter = 1
        print(f"Remapping {name} report page")
        # Iterate through visual containers in report page.
        for visual in self.page_json['visualContainers']:
            visual_string = json.dumps(visual)
            # Iterate through each field in mapping file, if field is found in visual container, replace with new value.
            for row in field_mapping_file:
                old_table = row[0]
                old_field = '{}"'.format(row[1])  # Add " to old_field to ensure it doesn't match longer strings
                new_table = row[2]
                new_field = '{}"'.format(row[3])  # Add " to new_field to ensure it doesn't match longer strings
                if visual_string.find(old_field) != -1:
                    visual_string = visual_string.replace(old_field, new_field)
                    visual_string = re.sub(
                        # Regex search term: Match any table which lies between an " and .new_field
                        fr'"[^"]*\.{re.escape(new_field)}',
                        '"' + new_table + '.' + new_field,
                        visual_string
                    )
                    # visual_string = re.sub(fr'"([^"]*)\.{re.escape(new_field)}', new_table, visual_string)
                    visual_string = visual_string.replace(
                        '"Entity": "{}"'.format(old_table),
                        '"Entity": "{}"'.format(new_table)
                    )
                else:
                    continue

                # Trying to insert additional table references if needed
                # visual = json.loads(visual_string)
                # try:
                #     match_count = 0
                #     for query in visual['config']['singleVisual']['prototypeQuery']['From']:
                #         if new_table == query['Entity']:
                #             match_count += 1
                #     if match_count == 0:
                #         table_dict = {}
                #         table_dict['Name'] = new_table[:1]
                #         table_dict['Entity'] = new_table
                #         table_dict['Type'] = 0
                #         visual['config']['singleVisual']['prototypeQuery']['From'].append(table_dict)
                #         visual_string = json.dumps(visual)
                #     else:
                #         visual_string = json.dumps(visual)
                # except KeyError:
                #     visual_string = json.dumps(visual)

            # Append repacked page to visual list
            visual = json.loads(visual_string)
            self.visuals.append(PowerBIReportVisual.PowerBIReportVisual(self, visual))
        print(f"{name} report page remapped.")

