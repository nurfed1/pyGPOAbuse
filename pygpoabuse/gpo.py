import asyncio
import logging
import re
from pygpoabuse.ldap import Ldap
from pygpoabuse.scheduledtask import ScheduledTask
from pygpoabuse.file import File
from pygpoabuse.service import Service


class GPO:
    extension_guids = {
        'scheduled_task': '{CAB54552-DEEA-4691-817E-ED4A4D1AFC72}',
        'file': '{3BAE7E51-E3F4-41D0-853D-9BB9FD47605F}',
        'service': '{CC5746A9-9B74-4be5-AE2E-64379C86E0E4}',
    }
    cse_guids = {
        'scheduled_task': '{AADCED64-746C-4633-A97C-D61349046527}',
        'file': '{7150F9BF-48AD-4DA4-A49C-29EF4A8369BA}',
        'service': '{91FBB303-0CD5-4055-BF42-E512A681B325}',
    }

    def __init__(self, smb_session, ldap_url):
        self._smb_session = smb_session
        self._ldap_url = ldap_url

    def update_extension_names(self, extension_type, extension_names):
        nullguid = "{00000000-0000-0000-0000-000000000000}"

        extension_guid = self.extension_guids[extension_type]
        cse_guid = self.cse_guids[extension_type]

        assignment = f'{cse_guid}{extension_guid}'

        if extension_names is None:
            extension_names = f"[{nullguid}{extension_guid}][{assignment}]"
            return extension_names

        if assignment in extension_names:
            # extension exists
            return extension_names

        extension_list = extension_names[1:-1].split('][')
        extension_list.append(assignment)

        found = False
        for i in range(len(extension_list)):
            assignment = extension_list[i]
            if assignment.startswith(nullguid):
                found = True
                break
                print(i, assignment)

        if found is False:
            extension_list.append(f"{nullguid}{extension_guid}")
        else:
            extension_guids = [f"{{{extension}}}" for extension in assignment[1:-1].split('}{')]
            extension_guids.append(extension_guid)
            extension_guids = sorted(extension_guids)
            extension_list[i] = ''.join(extension_guids)

        extension_list = sorted(extension_list)
        extension_names = '[' + ']['.join(extension_list) + ']'

        return extension_names

    async def update_ldap(self, url, domain, gpo_id, gpo_type, extension_type):
        ldap = Ldap(url, gpo_id, domain)
        r = await ldap.connect()
        if not r:
            logging.debug("Could not connect to LDAP")
            return False

        version = await ldap.get_attribute("versionNumber")

        if gpo_type == "Machine":
            attribute_name = "gPCMachineExtensionNames"
            updated_version = version + 1
        else:
            attribute_name = "gPCUserExtensionNames"
            updated_version = version + 65536

        extension_names = await ldap.get_attribute(attribute_name)

        if extension_names is False:
            logging.debug("Could not get {} attribute".format(attribute_name))
            return False

        if isinstance(extension_names, list):
            extension_names = extension_names[0]

        if extension_names == ' ':
            extension_names = None

        logging.debug("Old extensionName: {}".format(extension_names))
        updated_extension_names = self.update_extension_names(extension_type, extension_names)

        logging.debug("New extensionName: {}".format(updated_extension_names))

        await ldap.update_attribute(attribute_name, updated_extension_names, extension_names)
        await ldap.update_attribute("versionNumber", updated_version, version)

        await ldap.ldap_client.disconnect()

        return updated_version

    def update_versions(self, domain, gpo_id, gpo_type, extension_type):
        updated_version = asyncio.run(self.update_ldap(self._ldap_url, domain, gpo_id, gpo_type, extension_type))

        if not updated_version:
            return False

        logging.debug("Updated version number : {}".format(updated_version))

        try:
            tid = self._smb_session.connectTree("SYSVOL")
            fid = self._smb_session.openFile(tid, domain + "/Policies/{" + gpo_id + "}/gpt.ini")
            content = self._smb_session.readFile(tid, fid)
            # Added by @Deft_ to comply with french active directories (mostly accents)
            try:
                new_content = re.sub('=[0-9]+', '={}'.format(updated_version), content.decode("utf-8"))
            except UnicodeDecodeError:
                new_content = re.sub('=[0-9]+', '={}'.format(updated_version), content.decode("latin-1"))
            self._smb_session.writeFile(tid, fid, new_content)
            self._smb_session.closeFile(tid, fid)
        except Exception:
            logging.error("Unable to update gpt.ini file", exc_info=True)
            return False

        logging.debug("gpt.ini file successfully updated")
        return True

    def _check_or_create(self, base_path, path):
        for dir in path.split("/"):
            base_path += dir + "/"
            try:
                self._smb_session.listPath("SYSVOL", base_path)
                logging.debug("{} exists".format(base_path))
            except Exception:
                try:
                    self._smb_session.createDirectory("SYSVOL", base_path)
                    logging.debug("{} created".format(base_path))
                except Exception:
                    logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
                    return False
        return True

    def update_scheduled_task(self, domain, gpo_id, gpo_type, name="", mod_date="", description="", powershell=False, command="", force=False):

        try:
            tid = self._smb_session.connectTree("SYSVOL")
            logging.debug("Connected to SYSVOL")
        except Exception:
            logging.error("Unable to connect to SYSVOL share", exc_info=True)
            return False

        path = domain + "/Policies/{" + gpo_id + "}/"

        try:
            self._smb_session.listPath("SYSVOL", path)
            logging.debug("GPO id {} exists".format(gpo_id))
        except:
            logging.error("GPO id {} does not exist".format(gpo_id), exc_info=True)
            return False

        root_path = gpo_type

        if not self._check_or_create(path, "{}/Preferences/ScheduledTasks".format(root_path)):
            return False

        path += "{}/Preferences/ScheduledTasks/ScheduledTasks.xml".format(root_path)

        try:
            fid = self._smb_session.openFile(tid, path)
            st_content = self._smb_session.readFile(tid, fid, singleCall=False).decode("utf-8")
            st = ScheduledTask(gpo_type, name=name, mod_date=mod_date, description=description,
                               powershell=powershell, command=command, old_value=st_content)
            tasks = st.parse_tasks(st_content)

            if not force:
                logging.error("The GPO already includes a ScheduledTasks.xml.")
                logging.error("Use -f to append to ScheduledTasks.xml")
                logging.error("Use -v to display existing tasks")
                logging.warning("C: Create, U: Update, D: Delete, R: Replace")
                for task in tasks:
                    logging.warning("[{}] {} (Type: {})".format(task[0], task[1], task[2]))
                return False

            new_content = st.generate_scheduled_task_xml()
        except Exception as e:
            # File does not exist
            logging.debug("ScheduledTasks.xml does not exist. Creating it...")
            try:
                fid = self._smb_session.createFile(tid, path)
                logging.debug("ScheduledTasks.xml created")
            except:
                logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
                return False
            st = ScheduledTask(gpo_type, name=name, mod_date=mod_date, description=description, powershell=powershell, command=command)
            new_content = st.generate_scheduled_task_xml()

        try:
            self._smb_session.writeFile(tid, fid, new_content)
            logging.debug("ScheduledTasks.xml has been saved")
        except:
            logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
            self._smb_session.closeFile(tid, fid)
            return False
        self._smb_session.closeFile(tid, fid)

        if self.update_versions(domain, gpo_id, gpo_type, "scheduled_task"):
            logging.info("Version updated")
        else:
            logging.error("Error while updating versions")

        return True

    def update_file(self, domain, gpo_id, gpo_type, source_path, destination_path, action, mod_date="", force=False):

        try:
            tid = self._smb_session.connectTree("SYSVOL")
            logging.debug("Connected to SYSVOL")
        except:
            logging.error("Unable to connect to SYSVOL share", exc_info=True)
            return False

        path = domain + "/Policies/{" + gpo_id + "}/"

        try:
            self._smb_session.listPath("SYSVOL", path)
            logging.debug("GPO id {} exists".format(gpo_id))
        except:
            logging.error("GPO id {} does not exist".format(gpo_id), exc_info=True)
            return False

        root_path = gpo_type

        if not self._check_or_create(path, "{}/Preferences/Files".format(root_path)):
            return False

        path += "{}/Preferences/Files/Files.xml".format(root_path)

        try:
            fid = self._smb_session.openFile(tid, path)
            st_content = self._smb_session.readFile(tid, fid, singleCall=False).decode("utf-8")
            st = File(source_path, destination_path, action, mod_date=mod_date, old_value=st_content)
            files = st.parse_files(st_content)

            if not force:
                logging.error("The GPO already includes a Files.xml.")
                logging.error("Use -f to append to Files.xml")
                logging.error("Use -v to display existing files")
                for file in files:
                    logging.warning(f"src: {file[0]} dst: {file[1]} ")
                return False

            new_content = st.generate_file_xml()
        except Exception as e:
            # File does not exist
            logging.debug("Files.xml does not exist. Creating it...")
            try:
                fid = self._smb_session.createFile(tid, path)
                logging.debug("Files.xml created")
            except:
                logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
                return False
            st = File(source_path, destination_path, action, mod_date=mod_date)
            new_content = st.generate_file_xml()

        try:
            self._smb_session.writeFile(tid, fid, new_content)
            logging.debug("Files.xml has been saved")
        except:
            logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
            self._smb_session.closeFile(tid, fid)
            return False
        self._smb_session.closeFile(tid, fid)

        if self.update_versions(domain, gpo_id, gpo_type, "file"):
            logging.info("Version updated")
        else:
            logging.error("Error while updating versions")

        return True

    def update_service(self, domain, gpo_id, gpo_type, service_name, action, mod_date="", force=False):

        try:
            tid = self._smb_session.connectTree("SYSVOL")
            logging.debug("Connected to SYSVOL")
        except:
            logging.error("Unable to connect to SYSVOL share", exc_info=True)
            return False

        path = domain + "/Policies/{" + gpo_id + "}/"

        try:
            self._smb_session.listPath("SYSVOL", path)
            logging.debug("GPO id {} exists".format(gpo_id))
        except:
            logging.error("GPO id {} does not exist".format(gpo_id), exc_info=True)
            return False

        root_path = gpo_type

        if not self._check_or_create(path, "{}/Preferences/Services".format(root_path)):
            return False

        path += "{}/Preferences/Services/Services.xml".format(root_path)

        try:
            fid = self._smb_session.openFile(tid, path)
            st_content = self._smb_session.readFile(tid, fid, singleCall=False).decode("utf-8")
            st = Service(service_name, action, mod_date=mod_date, old_value=st_content)
            services = st.parse_services(st_content)

            if not force:
                logging.error("The GPO already includes a Services.xml.")
                logging.error("Use -f to append to Services.xml")
                logging.error("Use -v to display existing services")
                for service in services:
                    logging.warning(f"service name: {service}")
                return False

            new_content = st.generate_service_xml()
        except Exception as e:
            # File does not exist
            logging.debug("Services.xml does not exist. Creating it...")
            try:
                fid = self._smb_session.createFile(tid, path)
                logging.debug("Services.xml created")
            except:
                logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
                return False
            st = Service(service_name, action, mod_date=mod_date)
            new_content = st.generate_service_xml()

        try:
            self._smb_session.writeFile(tid, fid, new_content)
            logging.debug("Services.xml has been saved")
        except:
            logging.error("This user doesn't seem to have the necessary rights", exc_info=True)
            self._smb_session.closeFile(tid, fid)
            return False
        self._smb_session.closeFile(tid, fid)

        if self.update_versions(domain, gpo_id, gpo_type, "service"):
            logging.info("Version updated")
        else:
            logging.error("Error while updating versions")

        return True
