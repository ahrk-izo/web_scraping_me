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




    driver.quit()


if __name__ == '__main__':

    LOGIN_DATA = get_login_data()
    pprint.pprint(LOGIN_DATA)

    web_scraping(LOGIN_DATA)

