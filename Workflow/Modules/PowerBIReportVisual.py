import copy
import json
from Modules import PowerBIReportPage


class PowerBIReportVisual:

    def __init__(self, parent: PowerBIReportPage, visual_json: json):
        self.parent = parent
        self.visual_json = visual_json
        for key in ['config', 'filters', 'query', 'dataTransforms']:
            if key in self.visual_json.keys() and isinstance(self.visual_json[key], str):
                self.visual_json[key] = json.loads(self.visual_json[key])

    def get_visual_json(self):
        return_json = copy.deepcopy(self.visual_json)
        for key in ['config', 'filters', 'query', 'dataTransforms']:
            if key in return_json.keys():
                return_json[key] = json.dumps(return_json[key])
        return return_json

