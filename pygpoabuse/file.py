import re
import uuid
from datetime import datetime, timedelta
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
from pathlib import PureWindowsPath


class File:
    _actions = {
        'create': 'C',
        'replace': 'R',
        'update': 'U',
        'delete': 'D',
    }

    def __init__(self, source_path, destination_path, action, mod_date="", old_value=""):
        if mod_date:
            self._mod_date = mod_date
        else:
            mod_date = datetime.now() - timedelta(days=30)
            self._mod_date = mod_date.strftime("%Y-%m-%d %H:%M:%S")
        self._guid = str(uuid.uuid4()).upper()
        self._action = self._actions[action]

        self._old_value = old_value

        destination_filename = PureWindowsPath(destination_path).name
        self._file_str_begin = """<?xml version="1.0" encoding="utf-8"?><Files clsid="{215B2E53-57CE-475c-80FE-9EEC14635851}">"""
        self._file_str = f"""<File clsid="{{50BE44C8-567A-4ed1-B1D0-9234FE1F38AF}}" name="{destination_filename}" status="{destination_filename}" image="0" changed="{self._mod_date}" uid="{{{self._guid}}}" bypassErrors="1"><Properties action="{self._action}" fromPath="{source_path}" targetPath="{destination_path}" readOnly="0" archive="1" hidden="0"/></File>"""
        self._file_str_end = """</Files>"""

    def generate_file_xml(self):
        if self._old_value == "":
            return self._file_str_begin + self._file_str + self._file_str_end

        return re.sub(r"< */ *Files>", self._file_str.replace("\\", "\\\\") + self._file_str_end, self._old_value)

    def parse_files(self, xml_files):
        elem = ET.fromstring(xml_files)
        files = []
        for child in elem.findall("*"):
            file_properties = child.find("Properties")
            source_path = file_properties.get('fromPath', '<unknown>')
            destionation_path = file_properties.get('targetPath', '<unknown>')
            files.append([
                source_path,
                destionation_path
            ])
        return files
