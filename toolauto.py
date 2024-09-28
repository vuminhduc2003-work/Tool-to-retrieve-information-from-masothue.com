import os
import time
import datetime
import threading
import pandas as pd
from tkinter import Tk, Label, Button, filedialog, Toplevel, Text, Scrollbar, StringVar
from tkinter.ttk import Progressbar
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging

# Configure logging
logging.basicConfig(filename='app.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Khởi tạo DataFrame cho file Excel mới với các cột
columns = ['Mã số thuế', 'Tên đơn vị', 'Đại chỉ', 'Người đại diện', 'Sđt']
df_excel = pd.DataFrame(columns=columns)

# Khởi tạo trình duyệt Chrome
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Biến cờ để kiểm soát quá trình tạm dừng
pause_processing = False
processing_thread = None

# Toàn cục cho thanh tiến trình và thời gian ước đoán
progress_var = None
progress_bar = None

def save_file_excel(file_path):
    """Lưu DataFrame vào file Excel."""
    try:
        df_excel.to_excel(file_path, index=False)
        logging.info(f"Đã lưu file Excel tại: {file_path}")
    except Exception as e:
        logging.error(f"Lỗi khi lưu file: {e}")


def extract_info(xpath_keyword_pairs):
    """Lấy thông tin từ trang web dựa trên từ khóa XPATH."""
    data = {}
    for key, xpath in xpath_keyword_pairs.items():
        try:
            element = WebDriverWait(driver, 2).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )
            data[key] = element.text.strip()
        except Exception as e:
            logging.error(f"Lỗi khi lấy thông tin {key}: {e}")
            data[key] = ""
    return data


def click_close_button():
    """Nhấn nút đóng khi nó xuất hiện."""
    timeout = 10  # Thay đổi giá trị này tùy thuộc vào nhu cầu của bạn
    start_time = time.time()

    while True:
        try:
            close_button = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/div[4]/div/div/div[3]/button'))
            )
            if close_button.is_displayed():
                close_button.click()
                logging.info("Nút đóng đã được nhấn.")
                break
        except Exception as e:
            if time.time() - start_time > timeout:
                logging.error(f"Không thể tìm thấy hoặc nhấn nút đóng trong {timeout} giây: {e}")
                break
            time.sleep(1)


def update_progress_bar(start_time, total_items, processed_items):
    """Cập nhật thanh tiến trình và thời gian ước đoán."""
    elapsed_time = time.time() - start_time
    if processed_items > 0:
        average_time_per_item = elapsed_time / processed_items
        remaining_items = total_items - processed_items
        estimated_time_remaining = average_time_per_item * remaining_items
    else:
        estimated_time_remaining = 0

    progress_percent = (processed_items / total_items) * 100
    progress_var.set(f"{progress_percent:.2f}% hoàn tất - Ước đoán thời gian còn lại: {estimated_time_remaining:.2f} giây")
    progress_bar['value'] = progress_percent
    progress_bar.update()


def process_data(file_path):
    global df_excel
    global pause_processing
    global progress_var
    global progress_bar

    df_csv = pd.read_csv(file_path)
    total_items = len(df_csv)
    processed_items = 0
    start_time = time.time()

    logging.info("Tên các cột trong file CSV: %s", df_csv.columns)

    for index, row in df_csv.iterrows():
        start_item_time = time.time()
        tax_code = str(row.iloc[0]).zfill(10).strip()

        if tax_code.endswith('.0'):
            tax_code = tax_code[:-2]

        if pd.notna(df_excel.loc[df_excel['Mã số thuế'] == tax_code, 'Tên đơn vị']).any():
            logging.info(f"Mã doanh nghiệp {tax_code} đã có dữ liệu, bỏ qua.")
            processed_items += 1
            update_progress_bar(start_time, total_items, processed_items)
            continue

        try:
            driver.get("https://masothue.com/")
            WebDriverWait(driver, 2).until(
                EC.visibility_of_element_located((By.ID, "search"))
            )

            search_box = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.ID, "search"))
            )
            search_box.clear()
            search_box.send_keys(tax_code)
            search_box.send_keys(Keys.RETURN)

            xpath_keyword_pairs = {
                "Tên đơn vị": "//span[@itemprop='name']",
                "Đại chỉ": "//tr[td/i[contains(@class, 'fa-map-marker')]]/td[@itemprop='address']/span",
                "Người đại diện": "//tr[@itemprop='alumni']/td/span[@itemprop='name']",
                "Sđt": "//tr[td/i[contains(@class, 'fa-phone')]]/td[@itemprop='telephone']/span"
            }

            try:
                error_message = WebDriverWait(driver, 2).until(
                    EC.visibility_of_element_located((By.XPATH, '//div[contains(@class, "alert-danger")]'))
                )
                logging.warning(f"Thông báo lỗi: {error_message.text.strip()}. Đang làm mới trang...")
                driver.refresh()
                time.sleep(2)
                continue
            except Exception:
                pass

            info = extract_info(xpath_keyword_pairs)

            if any(value.strip() for value in info.values()):
                df_excel = pd.concat([df_excel, pd.DataFrame([{
                    "Mã số thuế": tax_code,
                    "Tên đơn vị": info["Tên đơn vị"],
                    "Đại chỉ": info["Đại chỉ"],
                    "Người đại diện": info["Người đại diện"],
                    "Sđt": info["Sđt"]
                }], columns=columns)], ignore_index=True)

                save_file_excel('C:/Users/vumin/PycharmProjects/automst/ddd_partial.xlsx')

                end_item_time = time.time()
                duration = end_item_time - start_item_time

                logging.info(f"Đã tìm thấy thông tin cho mã {tax_code}. Thời gian tìm kiếm: {duration:.2f} giây.")
            else:
                logging.info(f"Không tìm thấy thông tin cho mã {tax_code}. Đang lưu mã số thuế và làm mới trang...")
                df_excel = pd.concat([df_excel, pd.DataFrame([{
                    "Mã số thuế": tax_code,
                    "Tên đơn vị": "",
                    "Đại chỉ": "",
                    "Người đại diện": "",
                    "Sđt": ""
                }], columns=columns)], ignore_index=True)

                save_file_excel('C:/Users/vumin/PycharmProjects/automst/ddd_partial.xlsx')

                click_close_button()

                time.sleep(1)

                continue

            processed_items += 1
            update_progress_bar(start_time, total_items, processed_items)

        except Exception as e:
            logging.error(f"Lỗi với mã doanh nghiệp {tax_code}: {e}")

        while pause_processing:
            time.sleep(1)

    logging.info("Đã hoàn tất xử lý tất cả mã số thuế.")
    final_file_path = 'C:/Users/vumin/PycharmProjects/automst/ddd_final.xlsx'
    save_file_excel(final_file_path)
    logging.info(f"Số lượng bản ghi đã lưu: {len(df_excel)}")


def start_processing(file_path):
    global processing_thread
    global pause_processing

    if processing_thread and processing_thread.is_alive():
        return

    processing_thread = threading.Thread(target=process_data, args=(file_path,))
    processing_thread.start()


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    if file_path:
        start_processing(file_path)


def pause_processing_action():
    global pause_processing
    pause_processing = not pause_processing
    if pause_processing:
        logging.info("Đã tạm dừng xử lý. Nhấn nút Tạm dừng để tiếp tục.")
    else:
        logging.info("Đang tiếp tục xử lý.")


def show_saved_data():
    top = Toplevel()
    top.title("Dữ liệu đã lưu")

    text = Text(top, wrap='word', height=20, width=80)
    text.pack(side='left', fill='both', expand=True)

    scrollbar = Scrollbar(top, orient='vertical', command=text.yview)
    scrollbar.pack(side='right', fill='y')

    text.config(yscrollcommand=scrollbar.set)

    text.insert('end', df_excel.to_string(index=False))


def create_window():
    global progress_var
    global progress_bar

    root = Tk()
    root.title("Xử lý file CSV")

    label = Label(root, text="Chọn file CSV để xử lý và lưu kết quả.")
    label.pack(pady=10)

    button_select = Button(root, text="Chọn file CSV", command=select_file)
    button_select.pack(pady=10)

    button_pause = Button(root, text="Tạm dừng/ Tiếp tục", command=pause_processing_action)
    button_pause.pack(pady=10)

    button_show = Button(root, text="Hiển thị dữ liệu đã lưu", command=show_saved_data)
    button_show.pack(pady=10)

    progress_var = StringVar()
    progress_label = Label(root, textvariable=progress_var)
    progress_label.pack(pady=10)

    progress_bar = Progressbar(root, orient='horizontal', length=300, mode='determinate')
    progress_bar.pack(pady=10)

    def update_progress_bar_gui():
        while processing_thread and processing_thread.is_alive():
            total_items = len(df_excel)
            processed_items = df_excel[df_excel['Tên đơn vị'].notna()].shape[0]
            update_progress_bar(start_time, total_items, processed_items)
            time.sleep(1)

    progress_thread = threading.Thread(target=update_progress_bar_gui)
    progress_thread.start()

    root.geometry("400x300")
    root.mainloop()


if __name__ == "__main__":
    try:
        create_window()
    finally:
        driver.quit()
        logging.info("Trình duyệt đã được đóng.")
