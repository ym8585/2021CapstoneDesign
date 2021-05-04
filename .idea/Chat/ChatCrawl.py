import json
import urllib.request
from datetime import datetime
from datetime import timedelta
import re
import openpyxl

def collectClip(channel, lim, Clientid, File):
    url = "https://api.twitch.tv/kraken/clips/top?channel=" + channel + "&limit=" + str(lim)
    req = urllib.request.Request(url, headers = {"Client-ID": Clientid, "Accept" : "application/vnd.twitchtv.v5+json"})
    u = urllib.request.urlopen(req)
    c = u.read().decode('utf-8')
    js = json.loads(c)

    collectChat(js, lim, Clientid, File)


def collectChat(j, limit, clientId, sheet):

    for num in range(int(limit)):
        id = j['clips'][num]['vod']['id']
        offset = j['clips'][num]['vod']['offset']
        duration = j['clips'][num]['duration']

        cursor = ""
        count = 0

        while(1):
            try:
                url2 = ""
                if count == 0:
                    url2 = "https://api.twitch.tv/kraken/videos/" + str(id) + "/comments?content_offset_seconds=" + str(offset)
                else:
                    url2 = "https://api.twitch.tv/kraken/videos/" + str(id) + "/comments?cursor=" + str(cursor)
                req2 = urllib.request.Request(url2, headers = {"Client-ID": clientId, "Accept" : "application/vnd.twitchtv.v5+json"})
                u2 = urllib.request.urlopen(req2)
                c2 = u2.read().decode('utf-8')
                j2 = json.loads(c2)
                endCount = 0
                try:
                    for number, com in enumerate(j2['comments']):

                        dateString = j2['comments'][number]['created_at']
                        if "." in dateString:
                            dateString = re.sub(r".[0-9]+Z","Z", dateString)
                        date = datetime.strptime(dateString, "%Y-%m-%dT%H:%M:%SZ")

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

                if endCount == 1:
                    break

                if j2['_next']:
                    cursor = j2['_next']

                count = count + 1

            except Exception as e:
                print(e)

            #print("채팅 가지고 오기 성공")

if __name__ == "__main__":
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(["title","game","Chat Time","commenter","message"])
    Channel = "sunbaKing"
    Limit = 5
    ClientId = "jj19xbkqbx3l9ca9o3mno5obymdkmb" # Client id 추가
    collectClip(Channel, Limit, ClientId, sheet)

    wb.save("C:/Users/joo01/OneDrive/바탕 화면/chat/"+Channel+"Chat.xlsx")