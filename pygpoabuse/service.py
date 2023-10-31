import re
import uuid
from datetime import datetime, timedelta
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET


class Service:
    _actions = {
        'start': 'START',
        'restart': 'RESTART',
        'stop': 'STOP',
    }

    def __init__(self, service_name, action, mod_date="", old_value=""):
        if mod_date:
            self._mod_date = mod_date
        else:
            mod_date = datetime.now() - timedelta(days=30)
            self._mod_date = mod_date.strftime("%Y-%m-%d %H:%M:%S")
        self._guid = str(uuid.uuid4()).upper()
        self._action = self._actions[action]

        self._old_value = old_value

        self._file_str_begin = """<?xml version="1.0" encoding="utf-8"?><NTServices clsid="{2CFB484A-4E96-4b5d-A0B6-093D2F91E6AE}">"""
        self._file_str = f"""<NTService clsid="{{AB6F0B67-341F-4e51-92F9-005FBFBA1A43}}" name="{service_name}" image="0" changed="{self._mod_date}" uid="{{{self._guid}}}" userContext="0" removePolicy="0"><Properties startupType="NOCHANGE" serviceName="{service_name}" serviceAction="{self._action}" timeout="30"/></NTService>"""
        # self._file_str = f"""<NTService clsid="{{AB6F0B67-341F-4e51-92F9-005FBFBA1A43}}" name="{service_name}" image="0" changed="{self._mod_date}" uid="{{{self._guid}}}" userContext="0" removePolicy="0"><Properties startupType="NOCHANGE" serviceName="{service_name}" serviceAction="{self._action}" timeout="30"/><Filters><FilterRunOnce hidden="1" not="0" bool="AND" id="{{47511D93-4E18-4680-AFBA-3F463BC48C4C}}"/></Filters></NTService>"""
        self._file_str_end = """</NTServices>"""

    def generate_service_xml(self):
        if self._old_value == "":
            return self._file_str_begin + self._file_str + self._file_str_end

        return re.sub(r"< */ *NTServices>", self._file_str.replace("\\", "\\\\") + self._file_str_end, self._old_value)

    def parse_services(self, xml_files):
        elem = ET.fromstring(xml_files)
        services = []
        for child in elem.findall("*"):
            service_properties = child.find("Properties")
            service_name = service_properties.get('serviceName', '<unknown>')
            services.append(service_name)

        return services
