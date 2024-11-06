import requests
from bs4 import BeautifulSoup
import os
import re
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

# 設定基礎 URL 和查詢頁面 URL
# 國家教育研究院-全國中小學題庫網
BASE_URL = "https://exam.naer.edu.tw/"

# Step 1
# 複製查詢結果生成的網址 (以下範例為"新北市/國小/一年級"條件所生成的網址)
SEARCH_URL = "https://exam.naer.edu.tw/searchResult.php?page=1&orderBy=lastest&keyword=&selCountry=30&selCategory=41&selTech=0&selYear=&selTerm=&selType=&selPublisher=&chkCourses%5B%5D=53"

# 使用正則表達式替換 page 數值為 page={}
SEARCH_URL_TEMPLATE = re.sub(r'page=\d+', 'page={}', SEARCH_URL)

# Step 2
# 設定存儲資料夾，如果不想分類就同一個名稱用到底也可以。
DOWNLOAD_FOLDER = "下載試卷"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

class ExamDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("國教院題庫下載程式")
        self.root.geometry("600x600")
        self.search_url = None
        self.driver = None
        
        # 取得目前程式運行的目錄位置
        self.current_dir = os.getcwd()
        # 設定預設下載位置（目前目錄下的"下載試卷"資料夾）
        self.default_folder = "下載試卷"
        self.download_folder = os.path.join(self.current_dir, self.default_folder)
        
        # 確保下載資料夾存在
        try:
            os.makedirs(self.download_folder, exist_ok=True)
            print(f"下載資料夾已準備完成：{self.download_folder}")
        except Exception as e:
            print(f"建立下載資料夾時發生錯誤：{str(e)}")
        
        # 建立主要框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 下載位置設定區域
        self.create_download_path_frame()
        
        # 步驟區域
        self.create_steps_frame()
        
        # 下載進度區域
        self.create_progress_frame()
        
        # 日誌區域
        self.create_log_frame()
        
    def create_download_path_frame(self):
        path_frame = ttk.LabelFrame(self.main_frame, text="下載位置設定", padding="5")
        path_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 顯示目前下載位置（使用完整路徑）
        self.path_label = ttk.Label(path_frame, text=f"目前下載位置：{self.download_folder}")
        self.path_label.grid(row=0, column=0, pady=5, padx=5, sticky=tk.W)
        
        # 變更下載位置按鈕
        self.change_path_btn = ttk.Button(
            path_frame,
            text="變更下載位置",
            command=self.change_download_path
        )
        self.change_path_btn.grid(row=0, column=1, pady=5, padx=5)
        
    def change_download_path(self):
        new_path = filedialog.askdirectory(
            title="選擇下載位置",
            initialdir=self.download_folder
        )
        if new_path:  # 如果使用者有選擇資料夾（而不是按取消）
            self.download_folder = new_path
            os.makedirs(self.download_folder, exist_ok=True)
            self.path_label.config(text=f"目前下載位置：{self.download_folder}")
            self.log_message(f"下載位置已更改為：{self.download_folder}")
        
    def create_steps_frame(self):
        steps_frame = ttk.LabelFrame(self.main_frame, text="操作步驟", padding="5")
        steps_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 步驟 1：開啟瀏覽器
        self.open_browser_btn = ttk.Button(
            steps_frame, 
            text="1. 開啟瀏覽器並前往題庫網", 
            command=self.open_browser
        )
        self.open_browser_btn.grid(row=0, column=0, pady=5, sticky=tk.W)
        
        # 步驟 2：確認搜尋結果
        self.confirm_search_btn = ttk.Button(
            steps_frame, 
            text="2. 確認搜尋結果", 
            command=self.confirm_search,
            state=tk.DISABLED
        )
        self.confirm_search_btn.grid(row=1, column=0, pady=5, sticky=tk.W)
        
        # 步驟 3：開始下載
        self.start_download_btn = ttk.Button(
            steps_frame, 
            text="3. 開始下載", 
            command=self.start_download,
            state=tk.DISABLED
        )
        self.start_download_btn.grid(row=2, column=0, pady=5, sticky=tk.W)
        
        # 步驟 4：離開並關閉瀏覽器
        self.exit_btn = ttk.Button(
            steps_frame, 
            text="4. 離開並關閉瀏覽器", 
            command=self.exit_program
        )
        self.exit_btn.grid(row=3, column=0, pady=5, sticky=tk.W)
        
    def create_progress_frame(self):
        progress_frame = ttk.LabelFrame(self.main_frame, text="下載進度", padding="5")
        progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # 設定 progress_frame 的列/欄位權重，使其可以隨視窗調整大小
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var,
            maximum=100,
            length=400  # 改為 400
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
    def create_log_frame(self):
        log_frame = ttk.LabelFrame(self.main_frame, text="操作日誌", padding="5")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, width=60)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text['yscrollcommand'] = scrollbar.set
        
    def log_message(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        
    def open_browser(self):
        try:
            # 設定 Chrome 選項
            options = webdriver.ChromeOptions()
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            
            self.driver = webdriver.Chrome(options=options)
            self.driver.get(BASE_URL)
            self.log_message("已開啟瀏覽器，請選擇搜尋條件並按下搜尋按鈕")
            self.open_browser_btn.config(state=tk.DISABLED)
            self.confirm_search_btn.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("錯誤", f"開啟瀏覽器時發生錯誤：{str(e)}\n請確認是否已安裝 Chrome 瀏覽器")
            
    def confirm_search(self):
        try:
            if not self.driver:
                raise Exception("瀏覽器未開啟")
                
            # 確認使用者是否已完成搜尋
            if messagebox.askyesno("確認", "您是否已完成搜尋條件設定並按下搜尋按鈕？"):
                # 等待搜尋結果出現
                wait = WebDriverWait(self.driver, 10)
                wait.until(EC.presence_of_element_located((By.ID, "total_p")))
                
                # 獲取搜尋URL
                self.search_url = self.driver.current_url
                formatted_url = re.sub(r'page=\d+', 'page={}', self.search_url)
                
                # 只顯示確認網址
                self.log_message(f"確認搜尋網址：{self.search_url}")
                
                # 啟用下載按鈕
                self.start_download_btn.config(state=tk.NORMAL)
                self.confirm_search_btn.config(state=tk.NORMAL)
                
        except Exception as e:
            messagebox.showerror("錯誤", f"確認搜尋結果時發生錯誤：{str(e)}")
            
    def exit_program(self):
        try:
            if self.driver:
                self.driver.quit()
                self.log_message("已關閉瀏覽器")
            self.root.quit()
        except Exception as e:
            messagebox.showerror("錯誤", f"關閉程式時發生錯誤：{str(e)}")
            self.root.quit()
        
    def start_download(self):
        if not self.search_url:
            messagebox.showerror("錯誤", "尚未獲取搜尋網址")
            return
            
        # 在新執行緒中執行下載，避免凍結GUI
        threading.Thread(target=self.download_thread, daemon=True).start()
        self.start_download_btn.config(state=tk.DISABLED)
        
        # 下載完成後重新啟用下載按鈕
        def enable_download_button():
            self.start_download_btn.config(state=tk.NORMAL)
        
        # 設定一個計時器來重新啟用按鈕
        self.root.after(1000, enable_download_button)
        
    def download_thread(self):
        try:
            # 獲取第一頁來確定總頁數
            response = requests.get(self.search_url)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 找到總頁數
            total_pages_element = soup.find('span', id='total_p')
            if not total_pages_element:
                raise Exception("無法找到總頁數")
                
            total_pages = int(total_pages_element.text)
            
            # 處理每一頁
            for page in range(1, total_pages + 1):
                current_url = re.sub(r'page=\d+', f'page={page}', self.search_url)
                self.scrape_page(current_url)
                
                # 更新進度條
                progress = (page / total_pages) * 100
                self.progress_var.set(progress)
                
            messagebox.showinfo("完成", "所有檔案已下載完成！")
            
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
        finally:
            self.start_download_btn.config(state=tk.NORMAL)

    def scrape_page(self, page_url):
        response = requests.get(page_url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        rows = soup.find_all('tr')[1:]  # 跳過表頭
        for row in rows:
            cells = row.find_all('td')
            if len(cells) < 11:
                continue
            
            info = {
                "city": cells[0].text.strip(),
                "school": cells[1].text.strip(),
                "grade": cells[2].text.strip(),
                "year": cells[3].text.strip(),
                "subject": cells[5].text.strip(),
                "type": cells[6].text.strip(),
                "version": cells[7].text.strip(),
                "id": cells[8].text.strip()
            }

            pdf_name = self.parse_filename(info)

            # 下載試卷和答案
            exam_link = cells[9].find('a')
            if exam_link and not exam_link['href'].startswith("mailto:"):
                exam_pdf_url = urljoin(BASE_URL, exam_link['href'])
                self.download_pdf(exam_pdf_url, f"{pdf_name}_試卷.pdf")

            answer_link = cells[10].find('a')
            if answer_link and not answer_link['href'].startswith("mailto:"):
                answer_pdf_url = urljoin(BASE_URL, answer_link['href'])
                self.download_pdf(answer_pdf_url, f"{pdf_name}_答案.pdf")

    def parse_filename(self, info):
        # 生成檔名的邏輯
        return f"{info['city']}_{info['school']}_{info['grade']}_{info['year']}_{info['subject']}_{info['type']}_{info['version']}_{info['id']}"

    def download_pdf(self, pdf_url, file_name):
        response = requests.get(pdf_url)
        if response.status_code == 200:
            # 使用新的下載位置
            file_path = os.path.join(self.download_folder, file_name)
            with open(file_path, 'wb') as f:
                f.write(response.content)
            self.log_message(f"已下載：{file_name}")
        else:
            self.log_message(f"無法下載: {file_name}")

def main():
    root = tk.Tk()
    app = ExamDownloaderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
