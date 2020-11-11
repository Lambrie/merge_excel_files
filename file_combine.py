import openpyxl, argparse, os, pathlib, shutil
from datetime import datetime

class AgentReport(object):
    def __init__(self, inAgent, inLoggedInTime, inSignedOnTime, inBreakTime, inIncomingCalls, inAnsweredIncoming,
                 inTalkTimeIncoming,inAbandonedIncomingCalls, inOutgoingCallsExternal):
        self.agent = inAgent
        self.loggedInTime = inLoggedInTime
        self.signedOnTime = inSignedOnTime
        self.breakTime = inBreakTime
        self.incomingCalls = inIncomingCalls
        self.answeredIncoming = inAnsweredIncoming
        self.talkTimeIncoming = inTalkTimeIncoming
        self.abandonedIncomingCalls = inAbandonedIncomingCalls
        self.outgoingCallsExternal = inOutgoingCallsExternal


class Report(object):
    def __init__(self, inFileName): # ,inType, inCreated, inPeriod, inWeekly, inDaily, inResolution, inCycle
        self.fileName = inFileName
        self.reportType = None
        self.dateCreated = None
        self.period = None
        self.weekly = None
        self.daily = None
        self.resolution = None
        self.cycle = None
        self.content = []

    def extractDate(self):
        return str(datetime.strptime(self.period[:10],"%m/%d/%Y").strftime("%Y/%m/%d"))

    def addAgentReport(self, data):
        agentReport = AgentReport(data.get('agent',None),data.get('loggedIn',None),data.get('signedOn',None)
                                  ,data.get('breakTime',None),data.get('incomingCalls',None),
                                  data.get('answeredIncoming',None),data.get('talkTime',None),
                                  data.get('abondonedIncomingCalls',None),data.get('outgoingCalls',None))
        self.content.append(agentReport)


def getAllFileData(input_directory):
    reports = []
    inputFiles = [os.path.join(input_directory, f) for f in os.listdir(input_directory) if "master" not in f.lower()]
    for file in inputFiles:
        filePath = pathlib.Path(file)
        if filePath.suffix in [".xlsx"]:
            try:
                xlsx = openpyxl.load_workbook(filePath.absolute())
            except FileNotFoundError as e:
                print(f"{filePath.name} - {e}")
                continue
            except FileExistsError as e:
                print(f"{filePath.name} - {e}")
                continue
            except PermissionError as e:
                print(f"{filePath.name} - Please close open file")
                continue
            except Exception as e:
                print(f"{filePath.name} FILE_ERROR - {e}")
                continue
            try:
                sheet = xlsx["Table"]
            except KeyError as e:
                print(f"{e} - Ensure sheet 'Table' is present in file - {filePath.name}")
                continue
            report = Report(filePath.stem)
            for rowCount, row in enumerate(sheet.rows):
                content = {}
                if rowCount in range(0,9):
                    if "type" in str(row[0].value).lower():
                        report.reportType = str(row[1].value)
                    elif "created" in str(row[0].value).lower():
                        report.dateCreated = str(row[1].value)
                    elif "period" in str(row[0].value).lower():
                        report.period = str(row[1].value)
                    elif "weekly" in str(row[0].value).lower():
                        report.weekly = str(row[1].value)
                    elif "daily" in str(row[0].value).lower():
                        report.daily = str(row[1].value)
                    elif "resolution" in str(row[0].value).lower():
                        report.resolution = str(row[1].value)
                    elif "cycle" in str(row[0].value).lower():
                        report.cycle = str(row[1].value)
                    else: continue
                else:
                    if "log." in str(row[0].value).lower():
                        break
                    content["agent"] = str(row[0].value)
                    content["loggedIn"] = str(row[1].value)
                    content["signedOn"] = str(row[2].value)
                    content["breakTime"] = str(row[3].value)
                    content["incomingCalls"] = str(row[4].value)
                    content["answeredIncoming"] = str(row[5].value)
                    content["talkTime"] = str(row[6].value)
                    content["abondonedIncomingCalls"] = str(row[7].value)
                    content["outgoingCalls"] = str(row[8].value)
                    report.addAgentReport(content)
            reports.append(report)
    return len(inputFiles), reports


def mergeData(output_file, reports):
    filePath = pathlib.Path(output_file)
    try:
        xlsx = openpyxl.load_workbook(filePath.absolute())
    except FileNotFoundError as e:
        print(f"{filePath.name} - Not Found")
        xlsx = openpyxl.Workbook()
        print(f"Created new workbook -> {filePath.name}")
    except FileExistsError as e:
        print(f"{filePath.name} - {e}")
        xlsx = openpyxl.Workbook()
        print(f"{filePath.name} - Created new workbook")
    except PermissionError as e:
        print(f"{filePath.name} - Please close open file")
        return 0,0
    except Exception as e:
        print(f"{filePath.name} FILE_ERROR - {e}")
        return 0,0

    try:
        if not 'Results' in xlsx.sheetnames:
            xlsx.create_sheet('Results')
            sheet = xlsx["Results"]
            sheet.append(["Agent","Date","Logged In Time [totTLogin / All]","Signed On Time [totTSignon / All]",
                          "Break Time [totTPause / All]","Incoming Calls [totNNew<- / Tel]",
                          "Answered Incoming [totNConv<- / Tel]","Talk Time Incoming [totTConv<- / Tel]",
                          "Abandoned Incoming Calls [totNAban<- / Tel]","Outgoing Calls (External) [totNNew->Ext / Tel]"])
        else:
            sheet = xlsx["Results"]
    except KeyError as e:
        print(f"{e} - Ensure sheet 'Results' is present in file - {filePath.name}")
        return 0, 0

    successCounter = 0
    failedCounter = 0
    for report in reports:
        for data in report.content:
            try:
                sheet.append([data.agent,
                              str(report.extractDate()),
                              data.loggedInTime,
                              data.signedOnTime,
                              data.breakTime,
                              int(data.incomingCalls),
                              int(data.answeredIncoming),
                              data.talkTimeIncoming,
                              int(data.abandonedIncomingCalls),
                              int(data.outgoingCallsExternal)
                              ])
                successCounter +=1
            except Exception as e:
                print(e)
                failedCounter +=1
    xlsx.save(filename=filePath)
    return successCounter, failedCounter


def removeInputFile(input_directory):
    for filename in os.listdir(input_directory):
        if "master" in filename.lower(): continue
        file_path = os.path.join(input_directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
        else:
            print(f"{filename} removed from directory - {file_path}")


def run(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--input_directory',
        dest='input_directory',
        required=True,
        help='Directory of weekly input files')
    parser.add_argument(
        '--output_file',
        dest='output_file',
        required=True,
        help='Output file name and location to write results to.')
    known_args, _ = parser.parse_known_args(argv)

    print(f"File upload started from {known_args.input_directory}")
    inputFileCount, reports = getAllFileData(known_args.input_directory)
    print(f"{str(len(reports))}/{str(inputFileCount)} files loaded")
    print(f"Merging results into {known_args.output_file}")
    success, failed = mergeData(known_args.output_file, reports)
    print(f"{str(success)} rows merged successfully and {failed} records failed to merge")
    removeInputFile(known_args.input_directory)

if __name__ == '__main__':
  run()