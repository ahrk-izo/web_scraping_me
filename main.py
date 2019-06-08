'''
Webスクレイピング
とあるデータを抽出し、CSVおよびExcelに出力する
'''
import re
import time
import pprint
import configparser
import requests
import csv
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.chart import Reference
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd


def get_login_data():
    'Configファイルの読み込み、ログインに関するデータを取得する'
    login_data = {}
    env_cfg = configparser.ConfigParser()
    env_cfg.read('env.cfg')
    login_data['url'] = env_cfg['url']['login']
    login_data['contents_url'] = env_cfg['url']['contents']
    login_data['user_name'] = env_cfg['account']['username']
    login_data['pass'] = env_cfg['account']['pass']
    login_data['elem_login_id'] = env_cfg['element']['login_id']
    login_data['elem_pass_id'] = env_cfg['element']['pass_id']
    login_data['elem_btn_class'] = env_cfg['element']['btn_class']

    return login_data


def web_scraping(login_data):
    'Seleniumでサイトにアクセスし、データを抽出する'

    driver = webdriver.Chrome()
    try:
        driver.get(login_data['url'])
        WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.ID, "top"))
        )
    except:
        driver.quit()
        return

    # ログイン用CSSスタイル要素取得
    elem_user_name = driver.find_element_by_id(login_data['elem_login_id'])
    elem_pass = driver.find_element_by_id(login_data['elem_pass_id'])
    elem_login_btn = driver.find_element_by_class_name(
                                    login_data['elem_btn_class'])
    time.sleep(2)
    # フォームに入力、ボタンクリック
    elem_user_name.send_keys(login_data['user_name'])
    elem_pass.send_keys(login_data['pass'])
    time.sleep(2)
    elem_login_btn.click()
    time.sleep(5)

    # コンテンツページアクセス
    try:
        driver.get(login_data['contents_url'])
        WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located
        )
    except:
        driver.quit()
        return

    # タイトル取得
    elems = driver.find_elements_by_class_name('header-title')
    titles = [elem.text for elem in elems]
    pprint.pprint(titles)

    # 日付取得
    elems = driver.find_elements_by_class_name('metadata-publish-date')
    dates = [elem.text for elem in elems]
    pprint.pprint(dates)

    # 記事
    elems = driver.find_elements_by_class_name('asset-content')
        
    elems = [elem.find_element_by_class_name('journal-content-article') for elem in elems]
    articles = [elem.text for elem in elems]
    # pprint.pprint(articles)
    # print(articles[0])

    # 画像
    # TODO: 画像の取得ができない
    # img_urls = [elem.find_element_by_tag_name('img').get_attribute('src') for elem in elems]
    # pprint.pprint(img_urls)

    # ページ遷移ボタン
    '''
    elem = driver.find_element_by_class_name('lfr-pagination-buttons')
    link_next_page = elem.find_element_by_link_text('次へ')  # 「次へ」のリンクを取得
    link_next_page.click()  # リンクに移動
    '''

    time.sleep(5)

    # CSV書き出し
    output_file = open('output.csv', 'w', newline='', encoding='shift_jis')
    output_writer = csv.writer(output_file)
    output_writer.writerow(['date', 'title'])
    for i, date in enumerate(dates):
        output_writer.writerow([date, titles[i]])
    output_file.close()

    # Excel書き出し
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'diary_data'
    border = Border(top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000'),
                    left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'))

    for i, item in enumerate(['date', 'title'], 1):  # indexは1オリジン
        cell_coordinate = ws.cell(row=1, column=i).coordinate
        ws[cell_coordinate].value = item
        ws[cell_coordinate].border = border

    for i, date in enumerate(dates):
        cell_coordinate = ws.cell(row=2+i, column=1).coordinate
        ws[cell_coordinate].value = date
        ws[cell_coordinate].border = border
        cell_coordinate = ws.cell(row=2+i, column=2).coordinate
        ws[cell_coordinate].value = titles[i]
        ws[cell_coordinate].border = border

    # 幅指定
    cell_column = ws.cell(row=1, column=2).column
    ws.column_dimensions[cell_column].width = 50


    wb.save('output.xlsx')





    driver.quit()


if __name__ == '__main__':

    LOGIN_DATA = get_login_data()
    pprint.pprint(LOGIN_DATA)

    # web_scraping(LOGIN_DATA)

    # '''
    # Excel書き出し
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'diary_data'
    for i, item in enumerate(['date', 'title', 'article_length'], 1):  # indexは1オリジン
        cell_coordinate = ws.cell(row=1, column=i).coordinate
        ws[cell_coordinate].value = item

    for i in range(0, 5):
        cell_coordinate = ws.cell(row=2+i, column=1).coordinate
        ws[cell_coordinate].value = i+11
        cell_coordinate = ws.cell(row=2+i, column=2).coordinate
        ws[cell_coordinate].value = 'aaaa'
        cell_coordinate = ws.cell(row=2+i, column=3).coordinate
        ws[cell_coordinate].value = 100 + i*5
        

    # 幅指定
    cell_column = ws.cell(row=1, column=2).column
    ws.column_dimensions[cell_column].width = 50

    # グラフ
    ref_obj = openpyxl.chart.Reference(ws, min_row=2, min_col=3, max_row=2+4, max_col=3)
    series_obj = openpyxl.chart.Series(ref_obj, title='test graph')
    chart_obj = openpyxl.chart.BarChart()
    # chart_obj.style = 11  # スタイル(なんかかっこいい)
    chart_obj.type = 'bar'  # 横軸
    chart_obj.append(series_obj)
    dates = Reference(ws, min_row=2, min_col=1, max_row=6)  # 軸に使う範囲指定
    chart_obj.set_categories(dates)  # 軸の範囲に指定
    ws.add_chart(chart_obj, 'D2')  # 表示位置指定

    wb.save('output.xlsx')
    # '''

