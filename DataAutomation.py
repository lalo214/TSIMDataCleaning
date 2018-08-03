import csv
import pandas as pd
import openpyxl

# cleans up list elements by removing colon suffix
def cleanTrail(listID):
    for i in range(len(listID)):
        if ":" in listID[i]:
            listID[i] = listID[i][:-3]
        if listID[i].endswith('x'):
            listID[i] = listID[i][:-1]
        if listID[i].endswith('yy'):
            listID[i] = listID[i][:-2]
        if listID[i].endswith(':1'):
            listID[i] = listID[i][:-2]

#pulls transaciton set out of map name
def sliceTSet(mapName):
    if 'ansix12' in mapName:
        tSet = mapName[13:16]
    else:
        tSet = (mapName[13:19]).upper()
    return tSet

#pulls version out of map name
def sliceVersion(mapName):
    if 'ansix12' in mapName:
        version = '00' + mapName[17:21]
    else:
        version = 'D  ' + (mapName[19:21]) + mapName[21:22].upper()
    return version

# fucntion cleans the raw TSIM data and makes updates to backlog. Adds TSIM map name to backlog
def pLogUpdate(csvFile):

    #Dictionary used for cleaned TSIM data
    d = {'Sender':[], 'Transaction Set':[], 'Version':[], 'EDICode':[], 'TSIM Map':[]}

    # while csvFile open
    with open(csvFile) as f:

        #gets rid of headers
        f.readline()
        csv_reader = csv.reader(f,delimiter = ',')

        #iterate through rows in CSV file
        for line_list in csv_reader:
            #create lists using delimited values from SENDER_EXT and RECEIVER_EXT
            #strip gets rid of whitespace
            sender_list = [x.strip() for x in line_list[4].split('|')]
            receiver_list = [x.strip() for x in line_list[5].split('|')]

            # cleans up garbage at end of strings
            cleanTrail(sender_list)
            cleanTrail(receiver_list)

            # nested iteration through both lists to find all combinations
            for sender in sender_list:
                for receiver in receiver_list:
                    #pulls transaction set and version from map name
                    d['Transaction Set'].append(sliceTSet(line_list[7]))
                    d['Version'].append(sliceVersion(line_list[7]))
                    # because the EDICode is comprised of 'SenderID + ReceiverID', the partner must be checked
                    # too see if its inbound or outbound
                    if '_' in line_list[0]:
                        d['Sender'].append(line_list[0])
                        d['EDICode'].append(sender + receiver)
                    else:
                        d['Sender'].append(line_list[2])
                        d['EDICode'].append(receiver + sender)
                    d['TSIM Map'].append(line_list[7])

    # creates dataframe from dictionary
    df = pd.DataFrame(data=d)

    #write DataFrame to a file
    writer = pd.ExcelWriter(csvFile[:4] + 'Codes.xlsx', engine='xlsxwriter')
    df.to_excel(writer)
    writer.save()

    #open product backlog
    backLogFile = openpyxl.load_workbook('bCopy.xlsx')
    sheet = backLogFile['BacklogPlanning']

    #iterate through the product backlog
    for row in sheet.iter_rows(min_row = 0):
        #uses length of a dictionary list to use as index. Compares dicitonary index to backlog
        for i in range(len(d['Sender'])):
            #these three cell values must match for the map to be written to backlog
            if row[9].value == d['Transaction Set'][i] and row[10].value == d['Version'][i] and row[12].value == d['EDICode'][i]:
                row[19].value = d['TSIM Map'][i]
                row[2].value = 'Yes'
                row[3].value = 'Yes'

    backLogFile.save('bCopy.xlsx')

#makes updates based on each of these files
pLogUpdate('aana.csv')
pLogUpdate('ptus.csv')
pLogUpdate('oena.csv')
pLogUpdate('ttus.csv')
