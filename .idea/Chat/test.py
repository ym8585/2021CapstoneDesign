import json
import urllib.request
from datetime import datetime
from datetime import timedelta
import re
import openpyxl

def collectClip(id, clientId, File):
    #url = "https://api.twitch.tv/kraken/clips/top?channel=" + channel + "&limit=" + str(lim)
    url="https://api.twitch.tv/kraken/videos/" + str(id)
    req = urllib.request.Request(url, headers = {"Client-ID": clientId, "Accept" : "application/vnd.twitchtv.v5+json"})
    u = urllib.request.urlopen(req)
    c = u.read().decode('utf-8')
    js = json.loads(c)

    collectChat(js, clientId, File)


def collectChat(j, clientId, sheet):

   id=j['channel']['_id']
   offset = j['muted_segments']['offset']
   #duration = j['muted_segments']['duration']
   cursor=""

   while(1):
       url2 = "https://api.twitch.tv/kraken/videos/" + str(id) + "/comments?content_offset_seconds=" + str(offset)

       req2 = urllib.request.Request(url2, headers = {"Client-ID": clientId, "Accept" : "application/vnd.twitchtv.v5+json"})
       u2 = urllib.request.urlopen(req2)
       c2 = u2.read().decode('utf-8')
       j2 = json.loads(c2)

       try:
           for number, com in enumerate(j['comments']):
               dateString = j2['comments'][number]['created_at']
               if "." in dateString:
                   dateString = re.sub(r".[0-9]+Z","Z", dateString)
               date = datetime.strptime(dateString, "%Y-%m-%dT%H:%M:%SZ")

               if (int(offset) + int(duration)) < int(j2['comments'][number]['content_offset_seconds']):
                   endCount = 1
                   break
               else:
                   title = j['channel']['name']
                   #game = j['clips'][num]['game']
                   chattime = date +timedelta(hours=9)
                   commenters = j2['comments'][number]['commenter']['name']
                   message = j2['comments'][number]['message']['body']
                   sheet.append([title, chattime,commenters,message])

       except Exception as e:
           print(e)


if __name__ == "__main__":
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(["title","game","Chat Time","commenter","message"])
    #Channel = "flowervin"
    id=1019863132
    clientId = "mi99sqwm78lcz58dl5q18z0i9ea3bg" # Client id 추가
    collectClip(id, clientId, sheet)

    wb.save("C:/Users/joo01/OneDrive/바탕 화면/chat/"+"test_"+"Chat.xlsx")