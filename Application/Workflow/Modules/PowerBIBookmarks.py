import copy
import json
import re
from Modules import PowerBIReport


class PowerBIBookmark:
    def __init__(self, parent: PowerBIReport, page_json):
        self.parent = parent
        self.page_json = page_json
        self.visuals = []
        for visual in page_json['visualContainers']:
            self.visuals.append(PowerBIReportVisual.PowerBIReportVisual(self, visual))

    def get_page_json(self):
        page_json = []
        for visual in self.visuals:
            page_json.append(visual.get_visual_json())
        self.page_json['visualContainers'] = page_json
        return copy.deepcopy(self.page_json)
