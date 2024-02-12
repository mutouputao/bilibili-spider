import requests, xlrd, xlwt, time

def data(aid1):
    dataUrl = f'https://api.bilibili.com/x/web-interface/view?aid={aid1}'
    try:
        re1 = requests.get(dataUrl,headers=head).json()['data']
    except:
        print('Error')
    return re1


video = xlrd.open_workbook("video.xls")
video1 = xlwt.Workbook(encoding="UTF-8")
worksheet = video1.add_sheet("video")
sh1 = video.sheet_by_index(0)
allHang = sh1.nrows - 1
for hang in range(allHang):
    hang = hang + 1
    aid1 = sh1.cell_value(hang, 0)
    aid1 = int(aid1)
    head = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0"}
    mingzi = ["aid", "title", "view", "danmaku", "reply", "favorite", "coin", "share", "like"]
    for lie in range(9):
        hhh = data(aid1)
        if lie >= 2:
            hhh = hhh['stat']
            hhh = hhh[mingzi[lie]]
            worksheet.write(hang, lie, label=hhh)
        else:
            if lie == 0:
                worksheet.write(hang, lie, label=aid1)
            elif lie == 1:
                title = hhh['title']
                print(title)
                worksheet.write(hang, lie, label=title)
    if hang % 10 == 0:
        time.sleep(3)

tou = ["aid", "视频名称", "播放", "弹幕", "评论", "收藏", "硬币", "分享", "喜欢"]
for end in range(9):
    worksheet.write(0, end, label=tou[end])

worksheet.col(0).width = 13 * 256
worksheet.col(1).width = 50 * 256
video1.save("video1.xls")