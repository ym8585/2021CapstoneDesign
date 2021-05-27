import json
import urllib.request
from datetime import datetime
from datetime import timedelta
import re
import openpyxl

def collectVideo(videoID, clinentID, File):
    url = "https://api.twitch.tv/helix/videos?id="+str(videoID)
    req = urllib.request.Request(url, headers = {"Client-ID": clientID})
    u = urllib.request.urlopen(req)
    c = u.read().decode('utf-8')
    js = json.loads(c)

    collectChat(js, clinentID, File)


def collectChat(j, clientID, sheet):
    id = j['data']['id']
    channel = j['data']['user_name']
    duration=j['data']['muted_segments']['duration']
    offset = j['data']['muted_segments']['offset']

    while(1):
        try:
            for number, com in enumerate(j['comments']):
                datString = j['comments'][number]['created_at']
                if "." in datString:
                    datString = re.sub(r".[0-9]+Z","Z", dateString)
                data =  datetime.strptime(dateString, "%Y-%m-%dT%H:%M:%SZ")

                if (int(offset) + int(duration)) < int(j2['comments'][number]['content_offset_seconds']):
                    endCount = 1
                    break
                else:
                    title = j['clips'][num]['game']
                    game = j['clips'][num]['game']
                    chattime = date +timedelta(hours=9)
                    commenters = j2['comments'][number]['commenter']['name']
                    message = j2['comments'][number]['message']['body']
                    sheet.append([title, game, chattime,commenters,message])

        except Exception as e:
            print(e)

        if endCount==1:
            break


if __name__ =="__main__":
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(["title","game","Chat Time","commenter","message"])
    videoID = 1019863132
    clientID = "jj19xbkqbx3l9ca9o3mno5obymdkmb" # Client id 추가
    #방송날짜 받고싶음
    #방송하는 사람 아이디 받고싶음
    collectVideo(videoID, clientID, sheet)

    wb.save("C:/Users/joo01/OneDrive/바탕 화면/chat/"+"test_1"+"_Chat.xlsx")