import xlwt
import MySQLdb
import time
import datetime
print datetime.datetime.now()
filePath = "D:\\2014\\June\\18-06-2014\\"
country = "Japan"
choice = 'y'

countries = {
             'asia':['Taiwan','Japan','Indonesia','Philippines','Thailand','Turkey','China','India', 'South Korea', 'China', 'Australia'],
             'usa':['USA'],
             'eu':['GR','CY','SK','CZ','HU','DE', 'RU', 'GB', 'IT', 'FR', 'DE']
             }

ips = {
             'asia':['ec2-54-255-106-72.ap-southeast-1.compute.amazonaws.com','tvmao_epg'],
             'usa':['ec2-54-235-209-55.compute-1.amazonaws.com','EPG'],
             'eu':['ec2-54-83-33-83.compute-1.amazonaws.com','rovi']
             }
for r in range(0,1):
        for i in countries:
                if len(filter(lambda x:x == country,countries.get(i))) == 1:
                        region = i

total=[]
                       
databaseips = [ips.get(region)[0],'zdbadmin','z3l4yi23',ips.get(region)[1]]
        
database = MySQLdb.connect (host= databaseips[0], user = databaseips[1], passwd = databaseips[2], db = databaseips[3], charset = 'utf8')

cursor = database.cursor()

queries = {
           "asia": """SELECT DISTINCT p.program_id, p.show_title, p.progtype, c.name FROM channels c join schedules s on c.channel_id=s.channel_id JOIN programs p ON s.program_id = p.program_id WHERE p.show_image IS NULL AND p.country='""" + country +"""' GROUP BY p.show_title""",
           "usa":"""SELECT DISTINCT p.tmsid, p.fulltitle, p.progtype, st.name FROM stations st JOIN schedules s ON st.sourceid = s.sourceid JOIN programinfo p ON s.tmsid = p.tmsid WHERE p.programimage IS NULL AND (p.progtype='TVSHOW' OR p.progtype='MOVIES' OR p.progtype='SPORTS') GROUP BY p.fulltitle""",
           "eu":"""SELECT DISTINCT p.program_id, p.long_title, p.progtype, s.full_name FROM headend h JOIN lineup l ON h.Headend_ID = l.Headend_ID JOIN source s ON l.Source_ID = s.Source_ID JOIN `schedule` sc ON s.Source_ID = sc.Source_ID JOIN program p ON sc.program_id = p.program_id WHERE p.progimage IS NULL AND h.country='""" + country + """' GROUP BY p.long_title"""
           }

cursor.execute(queries.get(region))
data = cursor.fetchall()
print "Data Retrieved + ", len(data)
for i in range(5000):
    sch_count = {
             "asia":"""SELECT count(*) as duplicates from schedules WHERE country='""" + country + """' and program_id='""" + str(data[i][0]) +"""'""",
             "usa":"""SELECT count(*) as duplicates from schedules WHERE tmsid='""" + str(data[i][0]) +"""'""",
             "eu":"""SELECT COUNT(*) AS duplicates FROM headend h JOIN lineup l ON h.Headend_ID=l.Headend_ID JOIN source s ON l.Source_ID=s.source_id JOIN `schedule` sc ON s.source_id=sc.source_id WHERE h.country='""" + country + """' AND program_id='""" + str(data[i][0]) +"""'"""
            }
    print str(len(data)) + " " + str(i)
##    print sch_count.get(region)
    cursor.execute(sch_count.get(region))
    c=cursor.fetchall()
    #print c[0][0]
    a=[data[i][0], data[i][1], data[i][2], data[i][3], c[0][0]]
    #print a 
    total.append(a)
stotal=sorted(total, key=lambda x: x[-1], reverse=True)
wbk=xlwt.Workbook(encoding="UTF-8")
sheet=wbk.add_sheet("Sheet1", cell_overwrite_ok=True)
for j in range(len(total)):
    sheet.row(0).write(0, "Program_id")
    sheet.row(0).write(1, "Program_Title")
    sheet.row(0).write(2, "Progtype")
    sheet.row(0).write(3, "Country")
    sheet.row(0).write(4, "Channel")
    sheet.row(0).write(5,"No of Schedules")
    sheet.row(j+1).write(0, stotal[j][0])
    sheet.row(j+1).write(1, stotal[j][1])
    sheet.row(j+1).write(2, stotal[j][2])
    sheet.row(j+1).write(3, country)
    sheet.row(j+1).write(4, stotal[j][3])
    sheet.row(j+1).write(5, stotal[j][4])
wbk.save(filePath + country + time.strftime("%d-%m-%Y")+ ".xls")
print datetime.datetime.now()
