from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import datetime

# 関数宣言
def count_elems(css_links):
    count = driver.find_elements_by_css_selector(css_links)
    elem_count = len(count) + 1
    return elem_count

def get_by_nth_child(kancho_name, css_links, elems, latter_str):
    for n in range(1, elems):
        css_links_perfect = css_links + ":nth-child(" + str(n) + latter_str
        print(kancho_name + " No." + str(n) + ": " + css_links_perfect)
        get_data(css_links_perfect, n)

def get_by_nth_of_type(kancho_name, css_links, elems, latter_str):
    for n in range(1, elems):
        css_links_perfect = css_links + ":nth-of-type(" + str(n) + latter_str
        print(kancho_name + " No." + str(n) + ": " + css_links_perfect)
        get_data(css_links_perfect, n)

def get_data(css_links_perfect, n):
    links = driver.find_elements(By.CSS_SELECTOR, css_links_perfect)
    for link in links:
        link_text = link.text
        link_url = link.get_attribute("href")
        data_list.append([kancho_name, date_text, link_text, link_url, n])
    return data_list


# URLリストの読み込み
wb = openpyxl.load_workbook("官庁新着情報URL.xlsx")
ws = wb["Sheet1"]

row_values_list = [] #URLとCSSセレクタを入れる配列

for row in ws.iter_rows(min_row=2): 
    if row[0].value is None:
        break

    value_list = []
    for c in row: # cという変数名に意味はない 各列の値をひとつの要素として、一行ぶんの値を配列にまとめる
        value_list.append(c.value)
    row_values_list.append(value_list) # 多次元配列として行ごとの値の配列を格納する

driver_path = "driver/chromedriver_mac87"
driver = webdriver.Chrome(executable_path=driver_path) # webdriverの作成
driver.implicitly_wait(5) # 要素が見つからなければ5秒待つ設定

data_list = [] # 読み取り結果を格納する

for value in row_values_list:
    kancho_name = value[0]
    kancho_url = value[1]
    css_date = value[2]
    css_links = value[3]
    
    driver.get(kancho_url)

    date_elem = driver.find_element(By.CSS_SELECTOR, css_date)
    date_text = date_elem.text

    if kancho_name == "財務省":
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > dl > dd > a")
    elif kancho_name == "外務省":
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > a")
    elif kancho_name == "法務省":
        elems = count_elems(css_links)
        continue
        get_by_nth_child(kancho_name, css_links, elems, ") > a")
    elif kancho_name == "厚生労働省":
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > a")   
    elif kancho_name == "農林水産省":
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > dd > a")   
    elif kancho_name == "防衛省":
        # 2月でsectionにidがつけられているので、可用性改善が必要
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > span.news__title > a")   
    elif kancho_name == "文部科学省":
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > a")
    elif kancho_name == "経済産業省":
        elems = count_elems(css_links)
        get_by_nth_child(kancho_name, css_links, elems, ") > div.left.txt_box > a")
    elif kancho_name == "総務省":
        elems = count_elems(css_links)
        get_by_nth_of_type(kancho_name, css_links, elems, ") > a")
    elif kancho_name == "環境省":
        elems = count_elems(css_links)
        get_by_nth_of_type(kancho_name, css_links, elems, ") > a")
    elif kancho_name == "国土交通省":
        elems = count_elems(css_links)
        get_by_nth_of_type(kancho_name, css_links, elems, ") > div > p > a")
    else:
        continue

driver.quit()
# print(data_list)
# exit()

# 新しいブックへの記入と保存
wb_new = openpyxl.Workbook()
ws_new = wb_new.worksheets[0]

row_num = 1

for data in data_list :
    ws_new.cell(row_num, 1).value = data[0] # 官庁名
    ws_new.cell(row_num, 2).value = data[1] #日付
    ws_new.cell(row_num, 3).value = data[2] #テキスト
    ws_new.cell(row_num, 4).value = data[3] #リンク
    ws_new.cell(row_num, 5).value = data[4] #n

    row_num += 1

today = datetime.date.today()
wb_new.save(str(format(today, '%Y%m%d')) + "官庁新着情報.xlsx")