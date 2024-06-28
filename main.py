import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from PIL import Image
import base64
import io
import random
import time
import sys
import multiprocessing
import os

def login(username, password, lesson_no, chrome_driver_path, queue):
    global driver
    def queue_print(msg):
        queue.put(msg)

    options = Options()
    options.add_argument("--disable-gpu")

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")

    try:
        service = Service(executable_path=chrome_driver_path)
        service.start()
    except:
        messagebox.showwarning("输入错误", "ChromeDriver路径错误")
        return

    driver = webdriver.Chrome(service=service, options=options)

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })

    driver.get("https://xk.fudan.edu.cn/xk/stdElectCourse.action")

    # Login
    username_input_box = driver.find_element(By.NAME, "username")
    username_input_box.send_keys(username)
    password_input_box = driver.find_element(By.NAME, "password")
    password_input_box.send_keys(password)
    login_button = driver.find_element(By.NAME, "submitBtn")
    login_button.click()

    try:
        error_element = WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='ui-state-error ui-corner-all']//span[contains(text(), 'failure')]"))
        )
        queue_print("用户名或者密码错误")
        driver.quit()
        return
    except:
        queue_print("登录成功")

    # Navigate to course selection page
    try:
        link_element = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='http://xk.fudan.edu.cn/xk/home.action' and text()='点击此处']"))
        )
        link_element.click()
    except Exception as e:
        pass

    driver.get("https://xk.fudan.edu.cn/xk/stdElectCourse!defaultPage.action")

    try:
        time.sleep(1)
        original_window = driver.current_window_handle
        link_element = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@onclick='document.stdElectCourseIndexForm0.submit();']"))
        )
        link_element.click()
        WebDriverWait(driver, 2).until(EC.new_window_is_opened([original_window]))
        all_windows = driver.window_handles
        for window in all_windows:
            if window != original_window:
                driver.switch_to.window(window)
                break
        driver.switch_to.window(original_window)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except Exception as e:
        queue_print(f"无法找到或点击元素: {e}")

    # Search course
    while True:
        try:
            lesson_no_input_box = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.NAME, "electableLesson.no"))
            )
            lesson_no_input_box.send_keys(lesson_no)
            submit_button = driver.find_element(By.ID, "electableLessonList_filter_submit")
            time.sleep(2)
            actions = ActionChains(driver)
            actions.click_and_hold(submit_button).perform()
            time.sleep(0.1)
            actions.release(submit_button).perform()
        except Exception as e:
            queue_print("无法输入课程序号")

        try:
            tbody = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.ID, 'electableLessonList_data'))
            )
            rows = tbody.find_elements(By.TAG_NAME, 'tr')
            if len(rows) == 0:
                queue_print("没有找到课程")
                driver.quit()
                return
            elif len(rows) >= 2:
                queue_print("找到课程序号对应的多个课程,无法定位到具体课程，请重新输入课程序号")
                driver.quit()
                return
        except Exception as e:
            queue_print(f"无法找到课程列表")
            driver.quit()
            return

        try:
            element = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, "//a[@class='lessonListOperator' and @operator='ELECTION' and contains(@onclick, 'electCourseTable.tip.submit')]"))
            )
            element.click()
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
        except Exception as e:
            pass

        try:
            img_element = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "img[src^='data:image/']"))
            )
            img_src = img_element.get_attribute('src')
            img_data = img_src.split(',')[1]
            img_data = base64.b64decode(img_data)
            image = Image.open(io.BytesIO(img_data))
            width, height = image.size
            pixels = image.load()
            target_value = (192, 192, 192, 255)
            x_max = 0
            for x in range(width):
                for y in range(height):
                    if pixels[x, y] == target_value:
                        x_max = x
                        break
            coordinate_script = """
                var rect = arguments[0].getBoundingClientRect();
                return {
                    x: rect.left + window.scrollX,
                    y: rect.top + window.scrollY,
                    width: rect.width,
                    height: rect.height
                };
            """
            slider_bar = driver.find_element(By.CLASS_NAME, "slider-btn")
            image_coordinates = driver.execute_script(coordinate_script, img_element)
            slider_coordinats = driver.execute_script(coordinate_script, slider_bar)
            offset = x_max *(image_coordinates['width'] / width) - (slider_coordinats['x'] - image_coordinates['x'] + 22)
            actions = ActionChains(driver)
            actions.click_and_hold(slider_bar)
            actions.move_by_offset(offset, 0)
            actions.release().perform()
        except Exception as e:
            pass

        try:
            time.sleep(1)
            bar_element = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.ID, "cboxContent"))
            )
            result = bar_element.text.split("\n")
            if result[1] == "操作结果":
                queue_print(result[0])
                break
            else:
                queue_print(result[1])
            close_button = driver.find_element(By.ID, "cboxClose")
            close_button.click()
        except Exception as e:
            queue_print(e)
        driver.refresh()
        time.sleep(random.randint(1, 10))

    driver.quit()

def create_gui():
    def on_submit():
        global process
        global driver
        username = username_entry.get()
        password = password_entry.get()
        lesson_no = lesson_no_entry.get()
        chrome_driver_path = chrome_driver_path_entry.get()
        if username and password and lesson_no and chrome_driver_path:
            text_box.delete('1.0', tk.END)  # 清空文本框
            queue = multiprocessing.Queue()
            process = multiprocessing.Process(target=login, args=(username, password, lesson_no, chrome_driver_path, queue))
            process.start()
            window.after(100, process_queue, queue)
        else:
            messagebox.showwarning("输入错误", "请填写所有字段")

    def process_queue(queue):
        try:
            while True:
                msg = queue.get_nowait()
                text_box.insert(tk.END, msg + "\n")
                text_box.see(tk.END)
        except multiprocessing.queues.Empty:
            pass
        window.after(100, process_queue, queue)

    def on_stop():
        global process
        global driver
        if process and process.is_alive():
            process.terminate()
            process.join()
            if driver:
                driver.quit()
            text_box.insert(tk.END, "进程已终止\n")
            text_box.see(tk.END)

    def on_closing():
        global process
        global driver
        if process and process.is_alive():
            process.terminate()
            process.join()
            if driver:
                driver.quit()
        window.destroy()

    def show_instructions():
        instructions = (
            "请确保已安装谷歌浏览器，并下载与谷歌浏览器版本对应的ChromeDriver。\n"
            "1. 打开谷歌浏览器，在地址栏中输入 `chrome://version` 检查浏览器版本。\n"
            "2. 访问 https://sites.google.com/chromium.org/driver/downloads 下载对应版本的ChromeDriver。\n"
            "3. 将下载的ChromeDriver解压到一个方便的位置，例如 `E:/bin/chromedriver.exe`。\n"
            "4. 将ChromeDriver的路径拷贝到文本框中。\n"
            "点击\"开始\"按钮启动程序，每1-10s间进行一次选课,选课成功后自动停止。\n可点击\"停止\"按钮来终止程序。"
        )
        messagebox.showinfo("使用说明", instructions)

    window = tk.Tk()
    window.title("CourseMaster")

    tk.Label(window, text="学号:").grid(row=0, column=0, padx=10, pady=10)
    username_entry = tk.Entry(window)
    username_entry.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(window, text="密码:").grid(row=1, column=0, padx=10, pady=10)
    password_entry = tk.Entry(window, show="*")
    password_entry.grid(row=1, column=1, padx=10, pady=10)

    tk.Label(window, text="课程序号:").grid(row=2, column=0, padx=10, pady=10)
    lesson_no_entry = tk.Entry(window)
    lesson_no_entry.grid(row=2, column=1, padx=10, pady=10)

    tk.Label(window, text="ChromeDriver路径:").grid(row=3, column=0, padx=10, pady=10)
    chrome_driver_path_entry = tk.Entry(window)
    chrome_driver_path_entry.grid(row=3, column=1, padx=10, pady=10)

    instructions_button = tk.Button(window, text="使用说明", command=show_instructions)
    instructions_button.grid(row=4, columnspan=2, pady=10)

    submit_button = tk.Button(window, text="开始", command=on_submit)
    submit_button.grid(row=5, column=0, pady=10)

    stop_button = tk.Button(window, text="停止", command=on_stop)
    stop_button.grid(row=5, column=1, pady=10)

    text_box = tk.Text(window, height=10, width=50)
    text_box.grid(row=6, columnspan=2, padx=10, pady=10)

    window.protocol("WM_DELETE_WINDOW", on_closing)

    chrome_driver_path_entry.insert(0, "E:/bin/chromedriver.exe")

    window.mainloop()

if __name__ == "__main__":
    multiprocessing.freeze_support()  # 为了在Windows上支持多进程
    process = None
    driver = None
    create_gui()
