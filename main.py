import json
import os
import requests
import openpyxl

url = "https://weibo.com/ajax/statuses/mymblog?uid=1676317545&page={}&feature=0"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
    "Cookie": "",
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
            [
                "created_at",
                "reads_count",
                "reposts_count",
                "comments_count",
                "attitudes_count",
                "text_raw",
            ]
        )
    if os.path.exists(data_path + "page.txt"):
        with open(data_path + "page.txt") as f:
            first_page = int(f.read())
    for page in range(first_page + 1, 10000):
        print("page", page)
        response = requests.get(url.format(page), headers=headers)
        if response.status_code == 200:
            data = response.json()
        else:
            print("Request failed")
            break
        data = data["data"]["list"]
        if not data:
            print("No more data")
            break
        for i in data:
            sheet.append(
                [
                    i["created_at"],
                    i["reads_count"],
                    i["reposts_count"],
                    i["comments_count"],
                    i["attitudes_count"],
                    i["text_raw"],
                ]
            )
        workbook.save(data_path + "weibo.xlsx")
        with open(data_path + "page.txt", "w") as f:
            f.write(str(page))


if __name__ == "__main__":
    main()
