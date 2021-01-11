import os, argparse, sys, pathlib,logging, shutil, zipfile
from io import BytesIO
from datetime import datetime, timedelta
from win32com.client.gencache import EnsureDispatch as Dispatch

class Oli():
    def __init__(self, outlook_object, folder):
        self._obj = outlook_object
        self.folder = folder

    def items(self):
        for item_index in range(1,self._obj.Count+1):
            folder = self._obj[item_index]
            logging.info(f"Folder Name: {folder.Name}")
            if folder.Name == self.folder:
                return folder.Items
            else:
                for subfolder in [folder.Folders[i] for i in range(1, folder.Folders.Count + 1) ]:
                    logging.info(f"Sub Folder Name: {subfolder.Name}")
                    if subfolder.Name == self.folder:
                        return subfolder.Items
                    else:
                        try:
                            for subsubfolder in [subfolder.Folders[i] for i in range(1, subfolder.Folders.Count + 1) ]:
                                logging.info(f"Sub-Sub Folder Name: {subsubfolder.Name}")
                                if subsubfolder.Name == self.folder:
                                    return subsubfolder.Items
                        except:
                            continue


class Message(Oli):
    def __init__(self, outlook_object, folder, agent_out, volume_out):
        Oli.__init__(self,outlook_object, folder)
        self.agentSubject = "AgentDaily"
        self.volumeSubject = "Traffic Data Daily"
        self.agentDirectory = pathlib.Path(agent_out)
        self.volumeDirectory = pathlib.Path(volume_out)
        self.messages = Oli.items(self)

    def lastMessage(self):
        return self.messages.GetFirst() #GetLast()

    def getLookupDate(self, dt=datetime.today().date()):
        if dt.weekday() == 0: # Check on Mondays for Friday Emails
            findDate = dt - timedelta(days=3)
        else: # Check for emails from yesterday
            findDate = dt - timedelta(days=1)
        logging.info(f"Date to find: {findDate.strftime('%Y-%m-%d')}")
        return findDate.strftime("%Y-%m-%d")

    def messageByDate(self):
        findDate = self.getLookupDate()
        for message in self.messages:
            if message.SentOn.strftime("%Y-%m-%d") == findDate:
                yield message


    def saveAttachmentToFolder(self, attachment, location):
        filePath = os.path.join(location, attachment.FileName)
        try:
            attachment.SaveAsFile(filePath)
            logging.info(f"{attachment.FileName} saved to {location}")
        except Exception as e:
            logging.warning(f"{attachment.FileName} - Unable to save file - \n{e}")
        else:
            with zipfile.ZipFile(filePath) as zf:
                for file in zf.namelist():
                    if file.endswith(".xlsx"):
                        zf.extract(file, location)
            try:
                os.remove(filePath)
                # shutil.rmtree(os.path.join(location,attachment.FileName))
            except Exception as e:
                logging.error("Exception occurred", exc_info=True)


    def moveMessage(self):
        try:
            for msg in self.messageByDate():
                for attachment in msg.Attachments:
                    if '.zip' in attachment.FileName:
                        if self.agentSubject.lower() in msg.Subject.lower():
                            self.saveAttachmentToFolder(attachment, self.agentDirectory)
                        elif self.volumeSubject.lower() in msg.Subject.lower():
                            self.saveAttachmentToFolder(attachment, self.volumeDirectory)
        except Exception as e:
            logging.error("Exception occurred", exc_info=True)
            # print(f"{e}")


def run(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--outlook_folder',
        dest='outlook_folder',
        required=True,
        help='Custom folder on Outlook for Agent Reports')
    parser.add_argument(
        '--agent_input_directory',
        dest='agent_input_directory',
        required=True,
        help='Input directory for Agent Report Attachment')
    parser.add_argument(
        '--volume_input_directory',
        dest='volume_input_directory',
        required=True,
        help='Input directory for Volume Report Attachment')
    known_args, _ = parser.parse_known_args(argv)
    
    try:
        logging.basicConfig(filename='email_reader.log', filemode='w', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                            datefmt='%d-%b-%y %H:%M:%S')
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Move Reports
        if os.path.isdir(known_args.agent_input_directory) == False:
            logging.error("Invalid directory for argument --agent_input_directory")
            sys.exit()
        elif os.path.isdir(known_args.volume_input_directory) == False:
            logging.error("Invalid directory for argument --volume_input_directory")
            sys.exit()
        else:
            print("Running ...")
            latest_message = Message(outlook.Folders, known_args.outlook_folder,known_args.agent_input_directory,known_args.volume_input_directory)
            latest_message.moveMessage()
    except Exception as e:
        logging.error("Exception occurred", exc_info=True)
        sys.exit()


if __name__ == '__main__':
  if datetime.now().weekday() < 5:
    run()