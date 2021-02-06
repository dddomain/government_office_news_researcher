import datetime

today = datetime.date.today()
day = today.day
month = today.month
year = today.year

# 日付にm月y日を含む官庁
mylist = [
    "外務省",
    "文部科学省",
    "厚生労働省",
    "環境省",
    "農林水産省",
    "総務省",
    "経済産業省",
    "国土交通省"
]

def announce(kancho_name, today) :
    print('')
    print(str(format(today, '%Y年%m月%d日')) + "の" + kancho_name + "の新着情報はありません。")
    print('')

# 防衛省の日時バリデーション
def validation(kancho_name, date_text) :
    if kancho_name in mylist:
        if not str(month) + "月" + str(day) + "日" in date_text:
            return False
    elif kancho_name == "法務省":
        return False
    elif kancho_name == "防衛省":
        if not date_text == str(month) + "/" + str(day) :
            return False
    elif kancho_name == "財務省":
        if not str(format(today, '%m月%d日')) in date_text :
            return False
    else:
        return True