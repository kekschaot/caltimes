# -*- coding: UTF-8 -*-
import datetime
import time
import os
import getpass # gets the username
import sys
import re

version = "1.1_sonntag"

def hlp():
    print """
caltimes.exe
kalkuliert Zeiten aus Textfile.
generiert csv File für Excel

Das Textfile muss folgendes Format haben:
z.B.:  17:21 18.04.2013 1001 Hier kann ich Beschreibungen einfügen 17:27 18.04.2013
!! Hinweis !! im Notepad kann ein Zeitstempel mit F5 eingefügt werden !!!
Zeilen die mit # beginnen genauso wie Leerzeilen werden ignoriert.

'TTT' in der zeile bedeutet es wurde bereits ein ticket angelegt # bewirkt noch nichts
'FFF' in der zeile bedeutet heute ist ein feiertag # verdoppelt die zeit

"""

holidays=[]
try:
    # Try to load holiday dates from holidays.txt line by line each of them is a holyday
    for each in open("holidays.txt","r").readlines():
        b = each.split(".") # "01.05.2014" 
        holidays.append(b[2],b[1],b[0])
except:
    #ass
    print "!!!WARNING!!! COULD NOT READ HOLIDAYS FILE"




def banner():
    print """ caltimes version %s """ % version



# CONFIG
startExcel=True  # True Startet nach dem parsen excel; False macht nichts
multiFile=True # opens multiple files
multiFileDir=["./","*.txt"]

try:
    infd  = open("Zeiten.txt","r") # inputfile
except IOError as e:
    print e
    hlp()
    os.system("pause")
    exit() ## TODO






###### READABLE HACK COULD BE REMOVED IF YOU LIKE...
if os.name == 'nt' :
    os.system("mode 250,150") # hack to make cmd readable
######

##################DO NOT CHANGE ANYTHING BEYOND THIS LINE#######################
def writeToFile(dicEntry):
    #Write call to csv
    outfd.write("%d; %s; %s; %s; %s; %s; %s; %s\n" % (dicEntry['i'],
                                                         dicEntry['store'],
                                                         dicEntry['start'].date(),
                                                         dicEntry['start'].isocalendar()[1],
                                                         dicEntry['start'].time(),
                                                         dicEntry['end'].time(),
                                                         dicEntry['zeit'],
                                                         dicEntry['info']) )
def writeToCMD(dicEntry):
    #Write parsed call to shell
    print "%2d %10s %s KW: %s %s -> %s  ZEIT: %s INFO: %s" % (dicEntry['i'],
                                                              dicEntry['store'],
                                                              dicEntry['start'].date(),
                                                              dicEntry['start'].isocalendar()[1],
                                                              dicEntry['start'].time(),
                                                              dicEntry['end'].time(),
                                                              dicEntry['zeit'],
                                                              dicEntry['info'] )


def writeAllToCMD(parsed,times,anrufe,complett,durchschnitt):
    for day in sorted(parsed.iteritems()):
        for line in parsed[day[0]]:
            writeToCMD(line)
        print "\t\tZeit Tag %s: %s\n" %  (day[0], times[day[0]])
    print "ANRUFE: %d\tGESAMMTZEIT: %s\tDURCHSCHNITT: %s" % (anrufe,complett,durchschnitt)


def writeAllToFile(parsed,times,anrufe,complett,durchschnitt):
    print "writing to csv file:", outfd.name, "...",
    outfd.write(";;;;;;;Erstellt am: %s\n;;;;;;;Von: %s\n\n\n" % (time.ctime(),getpass.getuser(),))

    outfd.write("CALL;STORE;TAG;KW;VON;BIS;DAUER;INFO\n") # writing csv header

    for day in sorted(parsed.iteritems()):
        for line in parsed[day[0]]:
            writeToFile(line)
        outfd.write(";Tag:;%s;Zeit:;%s\n\n\n" %  (day[0], times[day[0]]))
    outfd.write("\n;;ANRUFE;%d;GESAMMT;%s;SCHNITT; %s" % (anrufe,complett,durchschnitt) )



def parse():
    LINES = [] # holds all the parsed lines
    PARSED = {} # dict to hold all the parsed lines by its date
    i = 0
    rawLinesParsed = 0 # used for error message info
    week=0
    complett = datetime.timedelta()

    for each in infd.readlines():
        rawLinesParsed += 1

        """ prepare the string """
        each = each.strip()
        each = each.replace("\t"," ")
        each = each.replace(";",",")
        #each = each.replace("\v"," ")

        
        if len(each) == 0 or each.startswith("#"): # if line is empty or an comment
            continue
        try:

            # Parse 
            (shour,sminute,sday,smonth,syear,store,info,ehour,eminute,eday,emonth,eyear) = tuple(re.findall("(\d{2}):(\d{2}) (\d{2}).(\d{2}).(\d{4})\s+(.*?)\s+(.+?)\s*?(\d{2}):(\d{2}) (\d{2}).(\d{2}).(\d{0,4})",each)[0])
            start = datetime.datetime(int(syear),int(smonth),int(sday),int(shour),int(sminute)) # build datetime obejct
            end = datetime.datetime(int(eyear),int(emonth),int(eday),int(ehour),int(eminute)) # build datetime object
            flags = {}
            if "TTT" in each:
                flags["ticket"] = True
            if "FFF" in each:
                flags["holiday"] = True
            
            i += 1 # count calls

            # Do calulactions...
            zeit = end - start # calculate arbeitszeit

            # Sunday calculation
            # if start == sonntag
            if start.weekday() == 6: # sonntag
                zeit = zeit * 2
            if flags.get("holiday") == True:
                zeit = zeit * 2
            
            complett += zeit # sum up workinhours

            # Sort the lines by date into a dict with lists
            try:
                PARSED[str(start.date())].append(
                            {
                                "i"     : i ,
                                "store" : store ,
                                "start" : start ,
                                "end"   : end ,
                                "zeit"  : zeit ,
                                "info"  : info ,
                            })
            except KeyError as e: # if the key is not jet here
                PARSED[str(start.date())] = [] # create it
                PARSED[str(start.date())].append( # then append the stuff again
                            {
                                "i"     : i ,
                                "store" : store ,
                                "start" : start ,
                                "end"   : end ,
                                "zeit"  : zeit ,
                                "info"  : info ,
                            })


        except (IndexError,ValueError) as e:
            print e
            print "WARING PARSING ERROR AT LINE: ", rawLinesParsed
            pass # error means malformed line...

    durchschnitt = complett / i

    return PARSED, durchschnitt, complett, i


def calTimesDays(PARSED): # generates dict with summed dates
    times = {} # holds the accumulated dates
    for day in PARSED:
        sum=datetime.timedelta(0)
        for line in PARSED[day]:
            sum += line["zeit"]
            times[day] = sum
    return times



def main():
    banner()

    parsed, durchschnitt, complett, anrufe = parse()

    times = calTimesDays(parsed)
    writeAllToCMD(parsed,times,anrufe,complett,durchschnitt)


    try:
        outfd = open("Zeiten.csv","w+") # outputfile für excel
    except IOError as e:
        print e
        if e.errno == 13:

            print "File is still opened in excel, i will NOT try to generate another..."
            startExcel=False
            raw_input()
            exit() ## TODO



    writeAllToFile(parsed,times,anrufe,complett,durchschnitt)



    print "Finished! ",
    outfd.close()
    if startExcel==True:
        print "starting excel."
        from os import system
        system("start %s" % outfd.name)

    print "Press any key to continue..."
    #os.system("pause")
    raw_input() # hack  cmd to stay open on windows


if __name__ == "__main__":
    main()
