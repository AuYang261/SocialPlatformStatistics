import json
import os
import requests
import openpyxl
import asyncio
import time
import random

url_init = "https://weibo.com/ajax/statuses/mymblog?uid=1676317545&page=1&feature=0"
url = "https://weibo.com/ajax/statuses/mymblog?uid=1676317545&page=1&feature=0&since_id={}"
stop_at = time.strptime("2025-01-01", "%Y-%m-%d")

headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Client-Version": "v2.47.138",
    "Priority": "u=1, i",
    "Referer": "https://weibo.com/u/1676317545",
    "sec-ch-ua": '"Chromium";v="142", "Microsoft Edge";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "server-version": "v2025.11.25.2",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36 Edg/142.0.0.0",
    "x-requested-with": "XMLHttpRequest",
    "x-xsrf-token": "C27M2WDfEC0FgEC8OQy3qAaW",
}

data_path = "data/"


def main():
    if os.path.exists("cookie.txt"):
        with open("cookie.txt") as f:
            headers["Cookie"] = f.read().strip()
    else:
        open("cookie.txt", "w").close()
        print("Please fill in the cookie in config.json")
        return
    if not os.path.exists(data_path):
        os.makedirs(data_path)
    if os.path.exists(data_path + "weibo.xlsx"):
        workbook = openpyxl.load_workbook(data_path + "weibo.xlsx")
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(
            ["创建时间", "阅读量", "转发量", "评论量", "点赞量", "内容", "视频", "直播"]
        )
    if os.path.exists(data_path + "since_id.txt"):
        with open(data_path + "since_id.txt") as f:
            since_id = int(f.read())
    else:
        since_id = None
    while True:
        while True:
            if since_id is None:
                response = requests.get(url_init, headers=headers)
            else:
                response = requests.get(url.format(since_id), headers=headers)
            if response.status_code == 200:
                data_json = response.json()
                break
            else:
                print(
                    f"Request failed({response.status_code}). random time later retrying..."
                )
                time.sleep(random.uniform(1, 5))
                print("Retrying...")
        data = data_json["data"]["list"]
        if not data:
            print("No more data")
            break
        since_id = data_json["data"]["since_id"]
        with open(data_path + "since_id.txt", "w") as f:
            f.write(str(since_id))
        for i in data:
            video_flag = False
            live_flag = False
            if i.get("url_struct"):
                for url_item in i["url_struct"]:
                    if "视频" in url_item.get("url_title"):
                        video_flag = True
                        break
                    if "直播" in url_item.get("url_title"):
                        live_flag = True
                        break
            sheet.append(
                [
                    i["created_at"],
                    i["reads_count"],
                    i["reposts_count"],
                    i["comments_count"],
                    i["attitudes_count"],
                    i["text_raw"],
                    "是" if video_flag else "否",
                    "是" if live_flag else "否",
                ]
            )
        workbook.save(data_path + "weibo.xlsx")
        # "Tue Dec 02 11:55:14 +0800 2025"
        # parse date string
        created_at = data[-1]["created_at"]
        print("Last data created at:", created_at)
        created_at_struct = time.strptime(created_at, "%a %b %d %H:%M:%S %z %Y")
        if created_at_struct < stop_at:
            print("Data before {}, stopping...".format(str(stop_at)))
            break

if __name__ == "__main__":
    main()
