import copy
import json
import csv
import zipfile
from Modules import PowerBIReportPage


class PowerBIReport:
    def __init__(self, pbix_location: str):
        self.binary = zipfile.ZipFile(pbix_location)
        self.layout = {}
        self.config_json = {}
        self.pages = []
        self.bookmarks = []
        self.resources = []
        self.page_sequence = 0
        self.layout_json = json.loads(self.binary.read('Report/Layout').decode('utf-16-le'))
        self.expand_report(self.layout_json)
        self.expand_config(self.layout_json)
        self.resource_items = self.layout['resourcePackages'][0]["resourcePackage"]["items"]
        try:
            self.connection = (self.binary.read('Connections'))
        except:
            self.connection = None

        print("Power BI file loaded")

    def expand_report(self, layout_json):
        self.layout = layout_json
        self.pages = []
        for page in self.layout['sections']:
            self.pages.append(PowerBIReportPage.PowerBIReportPage(self, page))
        self.page_sequence = len(self.pages) - 1

    def expand_config(self, layout_json):
        self.config_json = json.loads(layout_json['config'])

    def create_report(self, save_location: str):
        page_json = []
        self.layout['config'] = json.dumps(self.config_json)
        for page in self.pages:
            page_json.append(page.get_page_json())
        self.layout['sections'] = page_json

        with zipfile.ZipFile(save_location, 'w') as zout:
            for item in self.binary.infolist():
                if item.filename == 'Report/Layout':
                    zout.writestr(item, json.dumps(self.layout).encode('utf-16-le'))
                elif item.filename == '[Content_Types].xml':
                    xml = self.binary.read(item.filename).decode('utf-8')
                    xml = xml.replace("<Override PartName=\"/SecurityBindings\" ContentType=\"\" />", "")
                    zout.writestr(item, xml)
                elif item.filename == 'SecurityBindings':
                    continue
                elif item.filename == 'Connections':
                    zout.writestr(item, self.connection)
                # Remove unused resources
                elif item.filename.startswith('Report/StaticResources/RegisteredResources/'):
                    resource_name = item.filename.replace('Report/StaticResources/RegisteredResources/', '')
                    if resource_name in self.resources:
                        zout.writestr(item, self.binary.read(item.filename))
                    # elif resource_name.endswith('.json'):
                    #     zout.writestr(item, self.binary.read(item.filename))
                    else:
                        continue
                else:
                    zout.writestr(item, self.binary.read(item.filename))

    def retrieve_report_pages(self):
        return self.pages

    def export_page_info(self, file_name, save_directory):
        """Creates a CSV file containing page names, ID's and config info from the target pbix file."""
        with open(f"{save_directory}/{file_name} Report Pages.csv", "w", newline="") as output:
            writer = csv.writer(output)
            writer.writerow(["Page Name", "Page ID"])
            for page in self.layout["sections"]:
                writer.writerow([page['displayName'], page['name']])
        print("Finished extracting report information.")

    def required_resources(self, page):
        resources = []
        for visual in page.visuals:
            try:
                resources.append(visual.visual_json['config']['singleVisual']['objects']['general'][0]['properties']
                                  ['imageUrl']['expr']['ResourcePackageItem']['ItemName'])
            except KeyError:
                pass
            try:
                resources.append(visual.visual_json['config']['singleVisual']['objects']['shape'][0]['properties']
                                  ['map']['geoJson']['content']['expr']['ResourcePackageItem']['ItemName'])
            except KeyError:
                continue
        try:
            resources.append(page.config['objects']['background'][0]['properties']['image']['image']['url']['expr']
                             ['ResourcePackageItem']['ItemName'])
        except KeyError:
            pass

        try:
            if self.layout_json['theme'].endswith('.json'):
                resources.append(self.layout_json['theme'])
        except KeyError:
            pass
        finally:
            print(resources)
            return resources

    def add_retained_page(self, page_json: json):
        page_json = copy.deepcopy(page_json)
        if 'id' in page_json and page_json['name'] != 'ReportSection':
            del page_json['id']
        elif self.page_sequence == 0:
            page_json['id'] = 0
        page_json["ordinal"] = self.page_sequence
        new_page = PowerBIReportPage.PowerBIReportPage(self, page_json)
        self.pages.append(new_page)
        self.resources.extend(self.required_resources(new_page))
        return new_page

    def remove_resource_references(self, resources: list):
        retained_items = []
        for resource in resources:
            if resource['name'] in self.resources:
                retained_items.append(resource)
            elif resource['name'].endswith('.json'):
                retained_items.append(resource)
            else:
                continue
        return retained_items


    def remove_other_pages(self, pages_to_retain: list, pages_to_rename: dict):
        self.page_sequence = 0
        old_pages = self.layout.copy()
        self.pages = []
        page_count = len(old_pages['sections'])
        added_count = 0

        # Iterate through report pages
        for page in old_pages['sections']:
            # Check if the current page is in the list of pages to keep
            if page['name'] in pages_to_retain:
                # Check if the current page needs to change its display name
                for key in pages_to_rename.keys():
                    if page['name'] == key:
                        page['displayName'] = pages_to_rename[key]
                # Add the section to the new pbix file
                self.add_retained_page(page)
                self.page_sequence += 1
                added_count += 1
                print(f"Page: {page['displayName']} retained")
            else:
                print(f"Page: {page['displayName']} removed")
        print(f"Page removal complete.\nRemoved {page_count-added_count} pages.\nKept {added_count} pages.\n")
        # Remove unused resource references
        self.layout['resourcePackages'][0]['resourcePackage']['items'] = (
            self.remove_resource_references(self.resource_items))

    def remove_bookmarks(self, pages_to_retain: list):
        config_bookmarks = copy.deepcopy(self.config_json['bookmarks'])
        self.config_json['bookmarks'] = []
        self.bookmarks = []
        # Iterate through report bookmarks
        for bookmark in config_bookmarks:
            # Check if the current bookmark is used in the current report page
            if 'children' in bookmark.keys():
                for childBookmark in bookmark['children']:
                    if childBookmark['explorationState']['activeSection'] in pages_to_retain:
                        self.bookmarks.append(bookmark)
                        break
            elif bookmark['explorationState']['activeSection'] in pages_to_retain:
                self.bookmarks.append(bookmark)
        self.config_json['bookmarks'] = self.bookmarks

    def rename_table(self, table_list):
        table_count = len(table_list)
        current_table = 1
        json_str = json.dumps(self.layout)
        for row in table_list:
            old_table_name = row[0]
            new_table_name = row[1]
            json_str = json_str.replace(old_table_name + '.', new_table_name + '.')
            json_str = json_str.replace(
                '"Entity": "{}"'.format(old_table_name),
                '"Entity": "{}"'.format(new_table_name)
            )
            print(f'Updated table {current_table} of {table_count}.')
            current_table += 1
        self.expand_report(json.loads(json_str))
        print('Updated all tables')

    def repoint_field(self, field_mapping_file):
        page_count = len(self.pages)
        current_page = 1
        for page in self.pages:
            page.repoint_field(field_mapping_file)
            print(f"{current_page} of {page_count} report sections remapped.\n")
            current_page += 1
    # def repoint_field(self, old_field_reference, new_field_reference, old_table, new_table):
    #     json_str = json.dumps(self.layout)
    #     json_str = json_str.replace(old_field_reference, new_field_reference)
    #     json_str = json_str.replace('"Entity": "{}"'.format(old_table), '"Entity": "{}"'.format(new_table))
    #     self.init_report(json.loads(json_str))
