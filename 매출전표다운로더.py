import time

t1 = time.time()

import glob
import json
import os
import subprocess

from datetime import date, timedelta

import xlwings
from selenium import webdriver, common
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager

# region 0-0. 설정.json 파일읽기
print("0-0. 설정.json 파일읽는중", end="")
with open('설정.json', encoding='utf-8') as f:
    config = json.load(f)
print("\r0-0. 설정.json 파일읽기 완료")
# endregion 0-0. 설정.json 파일읽기

# region 0-1. 전역변수 설정
print("0-1. 전역변수 설정중", end="")
all_card_list = config['전체카드목록']
today = date.today()
if config.get('조회연월') is None or config['조회연월'] == '작월':
    last_day_of_last_month = today.replace(day=1) - timedelta(days=1)
    last_year_of_last_month = last_day_of_last_month.year
    last_month = last_day_of_last_month.month
else:
    last_year_of_last_month, last_month = list(map(int, config['조회연월'].split('-')))

print("\r0-1. 전역변수 설정 완료")
# endregion 0-1. 전역변수 설정

# region 0-2. 프로그램 로딩시간
t2 = time.time()
print(f"0-2. 프로그램 로딩시간 : {t2 - t1}초")
# endregion 0-2. 프로그램 로딩시간

# region 1-0. 보안프로그램 설치
print("1-0. 보안프로그램 설치 대기중", end="")
subprocess.call('nos_setup.exe')
print("\r1-0. 보안프로그램 설치 완료")
# endregion 1-0. 보안프로그램 설치

# region 1-1. 다운로드 폴더 체크
print("1-1. 다운로드 폴더 체크", end="")
download_path = config['파일경로']['다운로드'].replace('/', '\\')
os.path.isdir(download_path) or os.makedirs(download_path)

xl_app = xlwings.App(visible=False)
excel_total_card_history = xlwings.Book()

print("\r1-1. 다운로드 폴더 체크 완료")
# endregion 1-1. 다운로드 폴더 체크

# region 1-2. 로그인화면
print("1-2. 로그인 대기", end="")
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_argument("--start-fullscreen")
# options.add_argument("--window-size=1920,1080")
options.add_experimental_option('prefs', {'download.default_directory': download_path})
# FIXME 다음 링크에서 적절한 크롬드라이버 다운로드 : https://googlechromelabs.github.io/chrome-for-testing/
# driver = webdriver.Chrome(service=Service(ChromeDriverManager(version='116.0.5845.96').install()), options=options)
driver = webdriver.Chrome('chrome/chromedriver.exe', options=options)
driver.get('https://www.shinhancard.com/cconts/html/main.html')
time.sleep(1)

try:
    alert = driver.switch_to.alert
    if alert:
        alert.dismiss()
        driver.close()
        assert False, '보안프로그램 재설치 필요'
except common.exceptions.NoAlertPresentException:
    pass

wait_start = time.time()
while True:
    if 'crp/CRP72000N/CRP72000PH00.shc' in driver.current_url:
        break
    if time.time() - wait_start > 300:
        break
    time.sleep(1)

print("\r1-2. 로그인 완료")

history_button = driver.find_element(
    By.CSS_SELECTOR,
    '#contents > div.ly_inner > div.mainarea > div.main_cont_l > div.blue_box > dl:nth-child(2) > dt > a'
)
history_button.click()
time.sleep(3)


# endregion 1-2. 로그인화면


# region 신용카드 매출전표 다운로드
def num2char(num: int):
    start_idx = 1
    char = ''
    while num > 25 + start_idx:
        char += chr(65 + int((num - start_idx) / 26) - 1)
        num = num - (int((num - start_idx) / 26)) * 26
    char += chr(65 - start_idx + (int(num)))
    return char


def download_slip(card_type: str, excel_sum_card_history: xlwings.Book):
    card_group = all_card_list.get(card_type)
    if not card_group:
        print(f'2-0. 설정에서 {card_type} 항목의 카드를 정의하지 않음')
        return excel_sum_card_history

    # region 2-0. 조회화면(1)
    monthly_button = driver.find_element(
        By.CSS_SELECTOR,
        '#contents > div > div:nth-child(2) > div.card_info > div.info_wrap > form > div.conts_in > div > ul > li:nth-child(5) > div > div.conts_in > div > div > div.radio_btn.full.inner > label:nth-child(2) > span'
    )
    monthly_button.click()
    time.sleep(0.5)

    region = driver.find_element(By.XPATH,
                                 '//*[@id="contents"]/div/div[2]/div[2]/div[2]/form/div[4]/div/ul/li[4]/div/div/label[3]/span')
    region.click()
    time.sleep(0.5)

    select_month = Select(driver.find_element(By.CSS_SELECTOR, '#selMonth'))
    select_month.select_by_value(f'{last_year_of_last_month}{last_month:02d}')
    time.sleep(0.5)

    select_card_number = driver.find_element(By.CSS_SELECTOR,
                                             '#contents > div > div:nth-child(2) > div.card_info > div.info_wrap > form > div.conts_in > div > ul > li:nth-child(1) > div > div.radio_btn.full.outer.row2 > label:nth-child(3) > span'
                                             )
    select_card_number.click()
    time.sleep(0.5)
    # endregion 2-0. 조회화면(1)

    for project_name, project_card_list in card_group.items():
        for card_number in project_card_list:
            specific_download_path = f'{download_path}/{project_name}{card_number} {last_year_of_last_month}-{last_month:02d}'.replace(
                '/', '\\')
            os.path.isdir(specific_download_path) or os.makedirs(specific_download_path)

            # region 2-1. 카드검색
            print(f"2-1. {project_name} {card_number} 카드검색중", end="")

            for _ in range(2):
                driver.execute_script('window.scrollTo(0,0)')
                time.sleep(0.5)

            login_extend = driver.find_element(
                By.CSS_SELECTOR,
                '#header > div.head_top > div > div:nth-child(2) > div.head_btn_wrap > button:nth-child(1)'
            )
            login_extend.click()

            available_card_list = driver.find_element(By.CSS_SELECTOR, '#searchOpt3 > div > p.right.card > a')
            available_card_list.click()
            time.sleep(2)

            more_card = driver.find_element(By.CSS_SELECTOR,
                                            '#popup_card > article > div.pop_cont.pop_com_list.card > button')
            while 'display: inline-block;' in more_card.get_attribute('style'):
                more_card.click()
                time.sleep(1)

            total_card_count = len(driver.find_elements(By.CSS_SELECTOR,
                                                        '#popup_card > article > div.pop_cont.pop_com_list.card > div.scroll_cont.card > dl > div > dd'))

            for i in range(total_card_count):
                candidate_card_number = driver.find_element(By.CSS_SELECTOR,
                                                            f'#popup_card > article > div.pop_cont.pop_com_list.card > div.scroll_cont.card > dl > div > dd:nth-child({1 + i}) > a > p.cst_num > span'
                                                            )
                if card_number == candidate_card_number.text[-4:]:
                    candidate_card_number.click()
                    break
            else:
                raise Exception(f'{card_number} 카드번호가 존재하지 않음')

            time.sleep(2)
            for _ in range(3):
                search = driver.find_element(By.CSS_SELECTOR,
                                             '#contents > div > div:nth-child(2) > div.card_info > div.btn_wrap > a')
                search.click()

                time.sleep(0.5)

            print(f"\r2-1. {project_name} {card_number} 카드검색 완료")
            # endregion 2-1. 카드검색

            no_history = driver.find_element(By.CSS_SELECTOR, '#contents > div > div.list_no_result')
            if 'display: none;' in no_history.get_attribute('style'):
                # region 2-2. 카드내역다운로드
                print(f"2-2. {project_name} {card_number} 카드 이용내역 다운로드", end="")
                accounting_form = driver.find_element(By.CSS_SELECTOR,
                                                      '#contents > div > div.conts_box.list_detail > div.hd_title_wrap > div > a')
                accounting_form.click()
                time.sleep(1.5)
                save_history_to_excel = driver.find_element(By.CSS_SELECTOR,
                                                            '#pop_accForm > article > div.pop_cont.acc_form > div > div.t_right > button')
                save_history_to_excel.click()
                time.sleep(0.5)
                close_popup_button = driver.find_element(By.CSS_SELECTOR, '#pop_accForm > article > button')
                close_popup_button.click()
                time.sleep(3)

                while True:
                    if not glob.glob(download_path + '/*.crdownload'):
                        break
                    if time.time() - wait_start > 300:
                        break
                    time.sleep(0.5)

                ls = glob.glob(os.path.join(download_path, '법인이용내역*.xls'))
                sorted_ls = sorted(ls, key=lambda x: os.path.getmtime(x) * (-1))
                file = sorted_ls[0]

                excel_card_history = xlwings.Book(file)

                total_row_count = excel_sum_card_history.sheets.active.used_range.rows.count
                total_column_count = num2char(excel_sum_card_history.sheets.active.used_range.columns.count)
                this_row_count = excel_card_history.sheets.active.used_range.rows.count
                this_column_count = num2char(excel_card_history.sheets.active.used_range.columns.count)

                if f'{total_column_count}{total_row_count}' == 'A1':
                    excel_sum_card_history.sheets.active.range(
                        f'A1:{this_column_count}{this_row_count}').value = excel_card_history.sheets.active.range(
                        f'A1:{this_column_count}{this_row_count}').value
                else:
                    excel_sum_card_history.sheets.active.range(
                        f'A{total_row_count + 1}:{total_column_count}{total_row_count + this_row_count - 1}').value = excel_card_history.sheets.active.range(
                        f'A2:{total_column_count}{this_row_count}').value

                excel_card_history.save(f'{specific_download_path} 카드사용내역.xls')
                excel_card_history.close()
                os.remove(file)

                print(f"\r2-2. {project_name} {card_number} 카드 이용내역 다운로드 완료")
                # endregion 2-2. 카드내역다운로드

                # region 2-3. 매출전표 조회화면(2)
                print(f"2-3. {project_name} {card_number} 카드 매출전표 리스트업", end="")
                # more_wrapper = driver.find_element(By.CSS_SELECTOR,
                #                                    '#contents > div > div.conts_box.list_detail > div:nth-child(4)')
                # while 'display: block;' in more_wrapper.get_attribute('style'):
                used_count = int(driver.find_element(By.CSS_SELECTOR,
                                                     '#contents > div > div.conts_box.card_user > div:nth-child(2) > p.total_num > strong').text)
                for _ in range(used_count // 25):
                    more_button = driver.find_element(By.CSS_SELECTOR,
                                                      '#contents > div > div.conts_box.list_detail > div:nth-child(4) > button')
                    more_button.click()
                    time.sleep(2)

                # used_count = int(driver.find_element(By.CSS_SELECTOR,
                #                                      '#contents > div > div.conts_box.list_detail > div.accord_first.m_none.c_first > div > label > span > span.total').text)
                print(f"\r2-3. {project_name} {card_number} 카드 매출전표 리스트업 완료")
                # endregion 2-3. 매출전표 조회화면(2)

                # region 2-4. 매출전표화면
                print(f"2-4. {project_name} {card_number} 카드 매출전표 다운로드", end="")
                for i in range(used_count):
                    time.sleep(1)

                    used_date = driver.find_element(By.CSS_SELECTOR,
                                                    f'#contents > div > div.conts_box.list_detail > ul > li:nth-child({1 + 2 * i}) > div.check_btn > p:nth-child(2) > span.date').text
                    used_date = f"""{used_date[:10].replace('.', '-')} {used_date[12:14] or 'xx'}h{used_date[15:17] or 'xx'}m"""

                    shop_name = driver.find_element(By.CSS_SELECTOR,
                                                    f'#contents > div > div.conts_box.list_detail > ul > li:nth-child({1 + 2 * i}) > div.check_btn > label > span').text
                    detail_button = driver.find_element(By.CSS_SELECTOR,
                                                        f'#contents > div > div.conts_box.list_detail > ul > li:nth-child({1 + 2 * i}) > a')
                    detail_button.click()
                    time.sleep(1)

                    slip = driver.find_element(By.CSS_SELECTOR,
                                               '#contents > div > div.conts_box.list_detail > ul > li.li.on > div.check_txt > div > p > a:nth-child(1)'
                                               )
                    slip.click()
                    time.sleep(3)

                    driver.execute_script(f"""document.querySelector("#popup_stment > article").style.marginTop=''""")
                    driver.execute_script(
                        f"""document.querySelector("#btn > span:nth-child(2) > button").style.display='none'""")
                    driver.execute_script(
                        f"""document.querySelector("#btn > span:nth-child(4) > button").style.display='none'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > div.pop_cont.pop_stment.check.pop_sales01 > div").style.overflowY = 'visible'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article").style.height = '1280px'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > div.pop_cont.pop_stment.check.pop_sales01 > div").style.height = '1280px'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > div.pop_cont.pop_stment.check.pop_sales01 > div").style.maxHeight = '1280px'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > div.pop_head").style.padding = '0px'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > div.pop_cont.pop_stment.check.pop_sales01").style.padding = '0px'""")
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > button").style.display = 'none'""")
                    driver.execute_script(f"""document.querySelector("#popup_stment").style.padding = '0px'""")

                    tip = driver.find_element(By.CSS_SELECTOR,
                                              '#PrintAreaPop1 > div > table > tbody > tr:nth-child(7) > th').text
                    if tip == "봉사료":
                        driver.execute_script(
                            f"""document.querySelector("#PrintAreaPop1 > div > table > tbody > tr:nth-child(7)").remove()""")
                    time.sleep(1)
                    slip_area = driver.find_element(By.CSS_SELECTOR,
                                                    '#popup_stment > article > div.pop_cont.pop_stment.check.pop_sales01 > div'
                                                    )
                    time.sleep(1)
                    slip_area.screenshot(f'{specific_download_path}/{i:03d}. {used_date} {shop_name}.png')
                    driver.execute_script(
                        f"""document.querySelector("#popup_stment > article > button").style.display = 'block'""")

                    close_button = driver.find_element(By.CSS_SELECTOR, '#popup_stment > article > button')
                    close_button.click()

                    close_detail = driver.find_element(By.CSS_SELECTOR,
                                                       '#contents > div > div.conts_box.list_detail > ul > li.li.on > a')
                    close_detail.click()

                print(f"\r2-4. {project_name} {card_number} 카드 매출전표 다운로드 완료")
                # endregion 2-4. 매출전표화면
            else:
                print(f"2-2. {project_name} {card_number} 카드 이용내역이 없음")
                continue

    return excel_sum_card_history


excel_total_card_history = download_slip('신용카드', excel_total_card_history)
# endregion 신용카드 매출전표 다운로드

# region 연구카드 매출전표 다운로드
for _ in range(2):
    driver.execute_script('window.scrollTo(0,0)')
    time.sleep(0.5)
card_type = driver.find_element(By.CSS_SELECTOR, '#selectedCcd')
card_type.click()
time.sleep(0.5)

research_card = driver.find_element(By.CSS_SELECTOR, '#cusCcdList > li:nth-child(3) > a')
research_card.click()
time.sleep(0.5)

excel_total_card_history = download_slip('연구카드', excel_total_card_history)
excel_total_card_history.save(download_path + f'/과제전체 {last_year_of_last_month}-{last_month:02d} 카드사용내역.xls')
excel_total_card_history.close()
# endregion 연구카드 매출전표 다운로드

xl_app.kill()
driver.execute_script("alert('실행 완료')")
