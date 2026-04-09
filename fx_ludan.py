# -*- coding: utf-8 -*-
"""
分销录单脚本 fx_ludan.py
逻辑复制自 jinshuju.py 的 process_jinshuju_excel_new：
处理金数据采集完成的 Excel 表格（新版本），实现重单检测、录单处理等功能，使用新的 HHR 系统。
可独立运行：先启动 Chrome 调试端口，再执行本脚本并传入 Excel 路径。
"""

import os
import time
import sys
import json
import logging
from datetime import datetime

import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# 文件选择对话框（选择 xlsx）
try:
    from tkinter import Tk, filedialog
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

# 可选：Windows API / pyautogui（用于窗口置前）
try:
    import win32gui
    import win32con
    WINDOWS_API_AVAILABLE = True
except ImportError:
    WINDOWS_API_AVAILABLE = False

try:
    import pyautogui
    pyautogui_available = True
except ImportError:
    pyautogui_available = False

# 项目关键词匹配：标准词 + 同义词 + 模糊匹配（可选依赖 rapidfuzz）
try:
    from rapidfuzz import process as rapidfuzz_process
    from rapidfuzz import fuzz as rapidfuzz_fuzz
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    RAPIDFUZZ_AVAILABLE = False

# ---------- 全局变量 ----------
fendan_counter = 1
fendan_last_date = None
last_month = None
todaytime = ""
# 项目关键词匹配缓存（标准词列表、同义词字典）
_std_keywords_cache = None
_synonyms_cache = None

# ---------- 常量 ----------
DUPLICATE_PHONE_PREFIX = "电话："
# 程序启动时打开的网址（重单查询 + 录单页面）
DUPLICATE_CHECK_URL = "https://hhr.meb.com/manage/#/customer/recommend"
ORDER_ENTRY_URL = "https://hhr.meb.pw/manage/mobile/#/share?id=1356139172909420544"
# 推广码（录单时使用，默认值）
DEFAULT_PROMO_CODE = "fx-kol-lujyi9"
# 项目词库配置文件名（放在程序执行目录，用户可自行编辑）
XIANGMU_FILE = "xiangmu.txt"
# 新匹配方式：标准关键词列表、同义词映射（与 pipei.txt 一致）
STD_KEYWORDS_FILE = "std_keywords.txt"
SYNONYMS_FILE = "synonyms.json"
FUZZY_THRESHOLD = 80  # 模糊匹配分数阈值 (0-100)
DEFAULT_PROJECT_KEYWORD = "玻尿酸"


def _program_dir():
    """程序所在目录：打包后为 exe 目录，源码运行时为脚本所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _default_chromedriver_path():
    """优先使用程序目录下的 chromedriver.exe / chromedriver。"""
    base = _program_dir()
    for name in ("chromedriver.exe", "chromedriver"):
        p = os.path.join(base, name)
        if os.path.isfile(p):
            return p
    return None


# 环境变量 FX_CHROMEDRIVER 优先；否则使用程序目录下的 chromedriver；再无则用系统 PATH
CHROMEDRIVER_PATH = os.environ.get("FX_CHROMEDRIVER") or _default_chromedriver_path()


def _exe_dir():
    """可执行文件所在目录（打包后双击运行时工作目录常为 System32，需用此路径）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.getcwd()


def _resolve_data_path(filename):
    """优先 exe 同目录的配置，其次 PyInstaller 内置资源（_MEIPASS）。"""
    p = os.path.join(_exe_dir(), filename)
    if os.path.isfile(p):
        return p
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        p2 = os.path.join(sys._MEIPASS, filename)
        if os.path.isfile(p2):
            return p2
    return p


# ---------- 日志（仅命令行输出，不写文件）----------
logger = None


def setup_logging():
    """设置日志仅输出到命令行（stdout）"""
    global logger
    logger = logging.getLogger('fx_ludan')
    logger.setLevel(logging.INFO)
    for handler in logger.handlers[:]:
        try:
            handler.close()
        except Exception:
            pass
        logger.removeHandler(handler)
    console_handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    return logger


def log_info(message):
    if logger:
        logger.info(message)
    else:
        print(f"INFO: {message}")


def log_error(message):
    if logger:
        logger.error(message)
    else:
        print(f"ERROR: {message}")


def log_warning(message):
    if logger:
        logger.warning(message)
    else:
        print(f"WARNING: {message}")


def _is_browser_connection_lost(exc):
    """检测是否为浏览器已断开（无法连接 Chrome/Chromedriver），便于友好退出而非无限重试。"""
    msg = (str(exc) or "").lower()
    if "10061" in msg or "connection refused" in msg or "max retries exceeded" in msg or "积极拒绝" in msg:
        return True
    if "connection aborted" in msg or "connectionreset" in msg or "connection reset" in msg:
        return True
    return False


def log_progress(current, total, message=""):
    """输出数字进度到命令行，格式：[ 3/100 ] 消息"""
    progress_str = f"[ {current}/{total} ]"
    if message:
        full = f"{progress_str} {message}"
    else:
        full = progress_str
    if logger:
        logger.info(full)
    else:
        print(f"INFO: {full}")
    try:
        sys.stdout.flush()
    except Exception:
        pass


def _load_std_keywords_and_synonyms():
    """加载标准关键词列表与同义词映射（仅加载一次，带缓存）。"""
    global _std_keywords_cache, _synonyms_cache
    if _std_keywords_cache is not None and _synonyms_cache is not None:
        return _std_keywords_cache, _synonyms_cache
    std_path = _resolve_data_path(STD_KEYWORDS_FILE)
    syn_path = _resolve_data_path(SYNONYMS_FILE)
    std_list = []
    syn_dict = {}
    if os.path.exists(std_path):
        try:
            with open(std_path, "r", encoding="utf-8") as f:
                std_list = [line.strip() for line in f if line.strip()]
        except Exception as e:
            log_error(f"读取标准关键词失败: {e}，将使用默认词")
    else:
        log_error(f"标准关键词文件不存在: {std_path}，请放置 {STD_KEYWORDS_FILE}")
    if os.path.exists(syn_path):
        try:
            with open(syn_path, "r", encoding="utf-8") as f:
                syn_dict = json.load(f)
        except Exception as e:
            log_error(f"读取同义词映射失败: {e}")
    if not std_list:
        std_list = [DEFAULT_PROJECT_KEYWORD]
    _std_keywords_cache = std_list
    _synonyms_cache = syn_dict
    return _std_keywords_cache, _synonyms_cache


def get_project_keyword(project_info):
    """
    从项目完整内容中匹配出一个标准项目词，供录单时使用。
    匹配策略（与 pipei.txt 一致）：
      1. 精确包含：project_info 中包含的某个标准词，取最长匹配
      2. 同义词映射：project_info 中包含某同义词键，则返回对应标准词，取最长键匹配
      3. 模糊匹配：使用 rapidfuzz（若已安装）对 project_info 与标准词做 partial_ratio / token_sort_ratio
    若均无结果，返回标准词列表第一项（默认词）。
    """
    raw = (project_info or "").strip()
    if not raw:
        std_list, _ = _load_std_keywords_and_synonyms()
        return DEFAULT_PROJECT_KEYWORD

    std_list, synonyms = _load_std_keywords_and_synonyms()

    # 1. 精确包含：标准词出现在原文中，取最长（更具体优先）
    found_std = []
    for std in std_list:
        if std and std in raw:
            found_std.append(std)
    if found_std:
        return max(found_std, key=len)

    # 2. 同义词：原文中包含同义词键，则返回对应标准词，取最长键匹配
    for syn_key in sorted(synonyms.keys(), key=len, reverse=True):
        if syn_key and syn_key in raw:
            return synonyms[syn_key]

    # 3. 模糊匹配（需安装 rapidfuzz：pip install rapidfuzz）
    if RAPIDFUZZ_AVAILABLE and std_list:
        match = rapidfuzz_process.extractOne(
            raw, std_list, scorer=rapidfuzz_fuzz.partial_ratio, score_cutoff=FUZZY_THRESHOLD
        )
        if match:
            return match[0]
        match = rapidfuzz_process.extractOne(
            raw, std_list, scorer=rapidfuzz_fuzz.token_sort_ratio, score_cutoff=FUZZY_THRESHOLD
        )
        if match:
            return match[0]

    # 4. 默认兜底：固定返回玻尿酸
    return DEFAULT_PROJECT_KEYWORD


def get_project_keyword_with_meta(project_info):
    """
    与 get_project_keyword 逻辑一致，但额外返回匹配方式与模糊得分，供离线测试统计用。
    返回: (keyword, match_type, fuzzy_score)
    - match_type: "exact_contains" | "synonym" | "fuzzy_partial" | "fuzzy_token" | "default"
    - fuzzy_score: 仅模糊匹配时有值 (0-100)，否则 None
    """
    raw = (project_info or "").strip()
    std_list, synonyms = _load_std_keywords_and_synonyms()
    if not raw:
        return (DEFAULT_PROJECT_KEYWORD, "default", None)

    # 1. 精确包含
    found_std = []
    for std in std_list:
        if std and std in raw:
            found_std.append(std)
    if found_std:
        return (max(found_std, key=len), "exact_contains", None)

    # 2. 同义词
    for syn_key in sorted(synonyms.keys(), key=len, reverse=True):
        if syn_key and syn_key in raw:
            return (synonyms[syn_key], "synonym", None)

    # 3. 模糊匹配
    if RAPIDFUZZ_AVAILABLE and std_list:
        match = rapidfuzz_process.extractOne(
            raw, std_list, scorer=rapidfuzz_fuzz.partial_ratio, score_cutoff=FUZZY_THRESHOLD
        )
        if match:
            return (match[0], "fuzzy_partial", match[1])
        match = rapidfuzz_process.extractOne(
            raw, std_list, scorer=rapidfuzz_fuzz.token_sort_ratio, score_cutoff=FUZZY_THRESHOLD
        )
        if match:
            return (match[0], "fuzzy_token", match[1])

    # 4. 默认兜底：固定返回玻尿酸
    return (DEFAULT_PROJECT_KEYWORD, "default", None)


def check_and_reset_fendan_counter():
    """检查日期变化，如果是新的一天则重置计数器"""
    global fendan_counter, fendan_last_date
    current_date = datetime.now().strftime('%Y-%m-%d')
    if fendan_last_date is None or fendan_last_date != current_date:
        fendan_counter = 1
        fendan_last_date = current_date
        log_info(f"新的一天开始，分单计数器已重置为: {fendan_counter}")
    else:
        log_info(f"当前日期: {current_date}, 分单计数器: {fendan_counter}")
    return fendan_counter


def check_and_create_monthly_folder():
    """检查月份变化，如果是新的月份则创建文件夹"""
    global last_month
    current_month = datetime.now().strftime('%Y-%m')
    if last_month is None or last_month != current_month:
        folder_name = f"{current_month}"
        if not os.path.exists(folder_name):
            try:
                os.makedirs(folder_name)
                log_info(f"新月份开始，已创建文件夹: {folder_name}")
            except Exception as e:
                log_error(f"创建文件夹失败: {e}")
        else:
            log_info(f"文件夹已存在: {folder_name}")
        last_month = current_month
        return folder_name
    return f"{current_month}"


def force_window_to_foreground(driver, window_title=None):
    """强制将浏览器窗口置于最前端"""
    success = False
    if not window_title:
        try:
            window_title = driver.title
        except Exception:
            window_title = "Chrome"

    def verify_chrome_foreground():
        try:
            if not driver.current_window_handle:
                return False
            driver.execute_script("return document.readyState;")
            current_url = driver.current_url
            if "chrome://" in current_url or "http" in current_url or "https" in current_url:
                return True
            return False
        except Exception:
            return False

    try:
        if pyautogui_available:
            try:
                pyautogui.hotkey('win', 'd')
                time.sleep(0.8)
                pyautogui.hotkey('alt', 'tab')
                time.sleep(1.0)
                if verify_chrome_foreground():
                    driver.maximize_window()
                    time.sleep(0.5)
                    success = verify_chrome_foreground()
            except Exception as e:
                log_info(f"Win+D/Alt+Tab 置前失败: {e}")
        if not success:
            try:
                driver.execute_script("window.focus();")
                time.sleep(0.5)
                success = verify_chrome_foreground()
            except Exception:
                pass
        if not success:
            try:
                driver.maximize_window()
                time.sleep(0.5)
                success = verify_chrome_foreground()
            except Exception:
                pass
        if not success and WINDOWS_API_AVAILABLE:
            try:
                def enum_cb(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd):
                        text = win32gui.GetWindowText(hwnd)
                        if "Chrome" in text and window_title in text:
                            windows.append(hwnd)
                    return True
                chrome_windows = []
                win32gui.EnumWindows(enum_cb, chrome_windows)
                if chrome_windows:
                    hwnd = chrome_windows[0]
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.BringWindowToTop(hwnd)
                    time.sleep(0.5)
                    success = verify_chrome_foreground()
            except Exception as e:
                log_info(f"win32gui 置前失败: {e}")
        return success
    except Exception as e:
        log_error(f"窗口置前异常: {e}")
        return False


def hhrlogin(driver):
    """HHR 登录"""
    log_info("=== 开始执行 HHR 登录流程 ===")
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        account_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'][placeholder='请输入账号']"))
        )
        account_input.clear()
        account_input.send_keys("xxxxx")
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password'][placeholder='请输入密码']"))
        )
        password_input.clear()
        password_input.send_keys("xxxx")
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.el-button--primary span"))
        )
        login_button.click()
        time.sleep(3)
        try:
            add_customer_menu = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//li[@class='el-menu-item' and contains(text(), '添加客户')]"))
            )
            add_customer_menu.click()
            time.sleep(3)
            return True
        except TimeoutException:
            log_warning("未检测到'添加客户'菜单项，但登录可能已成功")
            return True
        except Exception as e:
            log_error(f"点击'添加客户'时出错: {e}")
            return False
    except Exception as e:
        log_error(f"HHR 登录失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def _save_feedback_excel(excel_file_path, df, status_list):
    """将录单结果输出为 录单反馈+日期.xlsx，保存到原表格所在目录。"""
    if len(status_list) != len(df):
        log_warning("状态条数与表格行数不一致，反馈表可能不完整")
    df_feedback = df.copy()
    df_feedback["状态"] = status_list
    out_dir = os.path.dirname(os.path.abspath(excel_file_path))
    date_str = datetime.now().strftime("%Y%m%d")
    out_path = os.path.join(out_dir, f"录单反馈+{date_str}.xlsx")
    try:
        df_feedback.to_excel(out_path, index=False)
        log_info(f"录单反馈已保存: {out_path}")
    except Exception as e:
        log_error(f"保存录单反馈表失败: {e}")


def _create_detail_logger(excel_file_path):
    """创建详细录单日志记录器，返回 (logger对象, 日志文件路径)"""
    out_dir = os.path.dirname(os.path.abspath(excel_file_path))
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    log_filename = f"日志{timestamp}.txt"
    log_path = os.path.join(out_dir, log_filename)

    detail_logger = logging.getLogger(f'detail_{timestamp}')
    detail_logger.setLevel(logging.INFO)
    # 清除之前的处理器
    for handler in detail_logger.handlers[:]:
        handler.close()
        detail_logger.removeHandler(handler)

    file_handler = logging.FileHandler(log_path, encoding='utf-8', mode='w')
    formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setFormatter(formatter)
    detail_logger.addHandler(file_handler)

    return detail_logger, log_path


def _detect_submit_success(driver, timeout=3):
    """
    检测录单是否成功提交。
    检查常见的成功提示：消息通知、成功弹层、toast提示等。
    返回: (是否成功, 检测到的消息文本)
    """
    success_selectors = [
        # Element UI 成功消息
        "div.el-message--success",
        "div.el-message.el-message--success",
        # Element UI Notification 通知
        "div.el-notification__content",
        "div.el-notification.success",
        # 常见的成功提示文字（通过XPath）
        "//*[contains(@class,'success') and (contains(text(),'成功') or contains(text(),'提交'))]",
        "//*[contains(@class,'el-message') and contains(text(),'成功')]",
        "//div[contains(text(),'提交成功')]",
        "//span[contains(text(),'提交成功')]",
        "//p[contains(text(),'提交成功')]",
        # Toast 提示
        "div.toast",
        "div.van-toast",
        "div.weui-toast",
    ]

    error_selectors = [
        # 错误提示
        "div.el-message--error",
        "div.el-message.el-message--error",
        "div.el-notification.error",
        "//*[contains(@class,'error') and contains(text(),'失败')]",
        "//div[contains(text(),'提交失败')]",
        "//div[contains(text(),'错误')]",
        "div.el-form-item__error",
    ]

    start_time = time.time()
    while time.time() - start_time < timeout:
        # 先检查成功提示
        for selector in success_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector) if not selector.startswith("//") else driver.find_elements(By.XPATH, selector)
                if elements:
                    for elem in elements:
                        text = elem.text or elem.get_attribute("textContent") or ""
                        text = text.strip()
                        if text and ("成功" in text or "提交" in text):
                            return True, text
            except Exception:
                pass

        # 检查错误提示
        for selector in error_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector) if not selector.startswith("//") else driver.find_elements(By.XPATH, selector)
                if elements:
                    for elem in elements:
                        text = elem.text or elem.get_attribute("textContent") or ""
                        text = text.strip()
                        if text and ("失败" in text or "错误" in text or "无法" in text):
                            return False, f"错误提示: {text}"
            except Exception:
                pass

        time.sleep(0.2)  # 快速轮询

    return None, "未检测到成功或错误提示"


def _log_detail(detail_logger, phone, status, reason=""):
    """记录详细录单日志"""
    if detail_logger:
        if status == "已录单":
            detail_logger.info(f"号码: {phone} | 状态: {status}")
        elif status == "失败":
            detail_logger.info(f"号码: {phone} | 状态: {status} | 失败原因: {reason}")
        elif status == "重单":
            detail_logger.info(f"号码: {phone} | 状态: {status} | 原因: 系统中已存在该客户")
        elif status == "跳过":
            detail_logger.info(f"号码: {phone} | 状态: {status} | 原因: 号码为空")
        else:
            detail_logger.info(f"号码: {phone} | 状态: {status} | 备注: {reason}")


def _ensure_duplicate_check_tab(driver):
    """确保当前在重单查询页标签，否则切换或打开；若被重定向到登录页则执行登录后再打开重单页。"""
    duplicate_check_url = DUPLICATE_CHECK_URL
    target_window = None
    for window_handle in driver.window_handles:
        driver.switch_to.window(window_handle)
        try:
            if duplicate_check_url in driver.current_url or "customer/recommend" in driver.current_url:
                target_window = window_handle
                break
        except Exception:
            continue
    if target_window:
        driver.switch_to.window(target_window)
        driver.refresh()
    else:
        driver.switch_to.new_window("tab")
        driver.get(duplicate_check_url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)
    current_url = driver.current_url
    if "signIn" in (current_url or ""):
        log_info("检测到登录页，尝试自动登录...")
        if hhrlogin(driver):
            time.sleep(2)
            driver.get(duplicate_check_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)
        else:
            log_error("自动登录失败，请手动登录后重试")


def run_flow1_duplicate_check_and_entry(driver, excel_file_path, df):
    """
    流程一：同一页面先重单查询再录单。
    若为重单则只记录重单并继续下一条；若不重单则在同一页执行录单。不修改原表，全部循环结束后输出 录单反馈+日期.xlsx 到原表所在路径。
    """
    global fendan_counter, todaytime
    duplicate_check_url = DUPLICATE_CHECK_URL
    fendan_counter = check_and_reset_fendan_counter()
    check_and_create_monthly_folder()
    fendan_content = ""

    # 创建详细日志记录器
    detail_logger, detail_log_path = _create_detail_logger(excel_file_path)
    log_info(f"详细录单日志将保存至: {detail_log_path}")

    def _cell(row, name, default=""):
        if name not in df.columns:
            return default
        v = str(row.get(name, default)).strip()
        return default if v == 'nan' or not v else v

    _ensure_duplicate_check_tab(driver)
    status_list = []
    total_rows = len(df)
    log_info(f"【流程一】共 {total_rows} 行数据，开始处理")
    log_info("=" * 50)

    for index, row in df.iterrows():
        current = index + 1
        log_progress(current, total_rows, f"正在处理第 {current} 行...")
        phone_number = _cell(row, "号码")
        if not phone_number:
            log_info(f"  第 {current} 行号码为空，跳过")
            status_list.append("跳过")
            _log_detail(detail_logger, phone_number if phone_number else "空号码", "跳过", "号码为空")
            continue

        try:
            city_info = _cell(row, "城市")
            project_raw = _cell(row, "项目")
            project_info = project_raw.split(")")[0].strip() if ")" in project_raw else project_raw
            wechat = _cell(row, "微信")
        except Exception as e:
            log_error(f"提取Excel数据时出错: {e}")
            wechat = project_info = city_info = ""

        try:
            phone_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='请输入客户手机号']"))
            )
            phone_input.clear()
            phone_input.send_keys(phone_number)
            try:
                driver.execute_script("arguments[0].blur();", phone_input)
            except Exception:
                pass
            log_info(f"已输入电话号码: {phone_number}，已失焦，等待 4 秒待页面反馈是否重单...")
            time.sleep(4)  # 号码框失焦后页面才会触发查重，再等待 3-5 秒显示重单提示
            # 再显式等待错误提示元素出现（最多 5 秒），避免因网络慢误判
            is_duplicate = False
            try:
                duplicate_error = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.el-form-item__error"))
                )
                error_text = (duplicate_error.text or "").strip()
                if "该客户已被推荐" in error_text or "去推荐其他客户吧" in error_text:
                    is_duplicate = True
            except TimeoutException:
                pass
            if is_duplicate:
                log_info(f"检测到重单: {phone_number}，仅记录并继续下一条")
                fendan_content += f"{fendan_counter}\n"
                fendan_content += f"{DUPLICATE_PHONE_PREFIX}{phone_number}\n"
                fendan_content += f"微信：{wechat if wechat else phone_number}\n"
                fendan_content += f"项目：{project_info}\n"
                fendan_content += f"查询城市：{city_info}\n"
                fendan_content += f"目前情况: 想要体验\n\n"
                fendan_counter += 1
                status_list.append("重单")
                _log_detail(detail_logger, phone_number, "重单", "系统中已存在该客户")
                continue
            log_info(f"未检测到重单: {phone_number}，执行录单...")

            status = "失败"  # 默认失败，确认成功后才改为已录单
            fail_reason = ""
            try:
                if wechat:
                    try:
                        wechat_input = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='请输入客户微信号（非必填）']"))
                        )
                        wechat_input.clear()
                        wechat_input.send_keys(wechat)
                    except Exception:
                        pass
                try:
                    name_input = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='请输入客户称呼']"))
                    )
                    name_input.clear()
                    name_input.send_keys("女士")
                except Exception as e:
                    fail_reason = f"输入客户称呼失败: {e}"
                    log_error(fail_reason)
                project_keyword = get_project_keyword(project_info)
                try:
                    project_input = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='试试搜索：瘦脸']"))
                    )
                    project_input.clear()
                    project_input.send_keys(project_keyword)
                    time.sleep(1)
                    first_option = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "li.el-cascader__suggestion-item"))
                    )
                    first_option.click()
                    time.sleep(2)
                except Exception as e:
                    fail_reason = f"输入项目失败: {type(e).__name__}: {e}"
                    log_error(fail_reason)
                    if _is_browser_connection_lost(e):
                        raise
                try:
                    textarea = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "textarea[placeholder*='说明客户情况']"))
                    )
                    textarea.clear()
                    textarea.send_keys(project_info)
                except Exception as e:
                    fail_reason = f"输入咨询内容失败: {type(e).__name__}: {e}"
                    log_error(fail_reason)
                    if _is_browser_connection_lost(e):
                        raise
                try:
                    city_input = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='试试搜索：成都']"))
                    )
                    city_input.clear()
                    city_input.send_keys(city_info)
                    time.sleep(1)
                    first_city_option = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "li.el-cascader__suggestion-item"))
                    )
                    first_city_option.click()
                    time.sleep(2)
                except Exception as e:
                    fail_reason = f"输入城市失败: {type(e).__name__}: {e}"
                    log_error(fail_reason)
                    if _is_browser_connection_lost(e):
                        raise

                # 点击提交按钮并检测结果
                submit_clicked = False
                try:
                    submit_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.el-button.save.el-button--default"))
                    )
                    submit_button.click()
                    time.sleep(2)
                    submit_clicked = True
                    log_info(f"  已点击提交按钮，等待检测成功提示...")
                except Exception:
                    try:
                        submit_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[@class='el-button save el-button--default is-round']//span[contains(text(), '提交')]/.."))
                        )
                        submit_button.click()
                        time.sleep(2)
                        submit_clicked = True
                        log_info(f"  已点击提交按钮(备用选择器)，等待检测成功提示...")
                    except Exception as e2:
                        fail_reason = f"点击提交按钮失败: {type(e2).__name__}: {e2}"
                        log_error(fail_reason)
                        status = "失败"

                if submit_clicked:
                    # 检测成功提示
                    is_success, message = _detect_submit_success(driver, timeout=3)
                    if is_success is True:
                        status = "已录单"
                        log_info(f"  检测到成功提示: {message}")
                    elif is_success is False:
                        status = "失败"
                        fail_reason = message
                        log_error(f"  检测到错误提示: {message}")
                    else:
                        # 未检测到任何提示，保留原状态（可能是成功提示消失太快）
                        # 记录警告但不强制设为失败，保持之前的状态
                        log_warning(f"  未检测到成功或错误提示，可能是提示消失太快。当前状态: {status}")
                        if status == "已录单":
                            log_info(f"  由于之前状态为已录单，保留该状态（可能实际已成功）")

                    # 刷新回到查询页
                    try:
                        driver.get(duplicate_check_url)
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        time.sleep(1)
                    except Exception:
                        pass

            except Exception as e:
                fail_reason = f"录单异常: {type(e).__name__}: {e}"
                log_error(f"流程一录单异常: {fail_reason}")
                status = "失败"
                if _is_browser_connection_lost(e):
                    log_error("检测到浏览器已断开连接，请重新运行脚本并重新登录。")
                    _log_detail(detail_logger, phone_number, "失败", f"浏览器断开: {e}")
                    raise
                try:
                    driver.get(duplicate_check_url)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(1)
                except Exception:
                    pass

            status_list.append(status)
            _log_detail(detail_logger, phone_number, status, fail_reason)
            log_info(f"  第 {index + 1} 行完成，状态: {status}")

        except Exception as e:
            error_msg = f"处理第{index+1}行时出错: {type(e).__name__}: {e}"
            log_error(error_msg)
            status_list.append("失败")
            _log_detail(detail_logger, phone_number, "失败", error_msg)
            if _is_browser_connection_lost(e):
                log_error("检测到浏览器已断开连接，请重新运行脚本并重新登录。")
                raise
            try:
                driver.get(duplicate_check_url)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(1)
            except Exception:
                pass

    if fendan_content:
        try:
            date_str = datetime.now().strftime('%Y-%m-%d')
            folder = last_month or datetime.now().strftime('%Y-%m')
            base_name = f"重单记录_{date_str}"
            # 当日多次运行用计数后缀避免覆盖：重单记录_YYYY-MM-DD.txt, _2.txt, _3.txt ...
            def _next_dup_number(prefix, dir_path):
                if not os.path.isdir(dir_path):
                    return 1
                existing = [f for f in os.listdir(dir_path) if f.startswith(prefix) and f.endswith(".txt")]
                for n in range(1, 9999):
                    cand = f"{prefix}.txt" if n == 1 else f"{prefix}_{n}.txt"
                    if cand not in existing:
                        return n
                return 1
            num = max(_next_dup_number(base_name, os.getcwd()), _next_dup_number(base_name, folder))
            suffix = "" if num == 1 else f"_{num}"
            notepad_file = f"{base_name}{suffix}.txt"
            month_notepad_file = os.path.join(folder, f"{base_name}{suffix}.txt")
            with open(month_notepad_file, 'w', encoding='utf-8') as f:
                f.write("=" * 50 + "\n")
                f.write(fendan_content)
            with open(notepad_file, 'w', encoding='utf-8') as f:
                f.write("=" * 50 + "\n")
                f.write(fendan_content)
            log_info(f"重单内容已保存到: {notepad_file}")
        except Exception as e:
            log_error(f"保存重单记录时出错: {e}")

    _save_feedback_excel(excel_file_path, df, status_list)
    n_ok = sum(1 for s in status_list if s == "已录单")
    n_dup = sum(1 for s in status_list if s == "重单")
    n_skip = sum(1 for s in status_list if s == "跳过")
    n_fail = sum(1 for s in status_list if s == "失败")
    log_info("=" * 50)
    log_info(f"【流程一】处理完成 | 共 {total_rows} 行 | 已录单: {n_ok} | 重单: {n_dup} | 跳过: {n_skip} | 失败: {n_fail}")


def run_flow2_entry_only(driver, excel_file_path, df):
    """
    流程二：无重单查询，逐条在分享页执行录单。全部结束后输出 录单反馈+日期.xlsx 到原表所在路径。
    说明：分享页若存在手机号/号码输入框，会尝试自动填写；若无该字段则仅填称呼、城市、项目、说明、微信。
    """
    order_entry_url = ORDER_ENTRY_URL

    # 创建详细日志记录器
    detail_logger, detail_log_path = _create_detail_logger(excel_file_path)
    log_info(f"详细录单日志将保存至: {detail_log_path}")

    def _cell(row, name, default=""):
        if name not in df.columns:
            return default
        v = str(row.get(name, default)).strip()
        return default if v == 'nan' or not v else v

    orig_win = driver.current_window_handle
    target_window = None
    for window_handle in driver.window_handles:
        driver.switch_to.window(window_handle)
        try:
            if order_entry_url in driver.current_url or "share?id=1356139172909420544" in driver.current_url:
                target_window = window_handle
                break
        except Exception:
            continue
    if target_window:
        driver.switch_to.window(target_window)
        driver.refresh()
    else:
        driver.switch_to.window(orig_win)
        driver.switch_to.new_window('tab')
        driver.get(order_entry_url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)

    status_list = []
    total_rows = len(df)
    log_info(f"【流程二】共 {total_rows} 行数据，开始处理")
    log_info("=" * 50)
    for index, row in df.iterrows():
        current = index + 1
        log_progress(current, total_rows, f"正在处理第 {current} 行（流程二）...")
        phone_number = _cell(row, "号码")
        if not phone_number:
            log_info(f"  第 {current} 行号码为空，跳过")
            status_list.append("跳过")
            _log_detail(detail_logger, phone_number if phone_number else "空号码", "跳过", "号码为空")
            continue

        try:
            city_info = _cell(row, "城市")
            project_raw = _cell(row, "项目")
            project_info = project_raw.split(")")[0].strip() if ")" in project_raw else project_raw
            wechat = _cell(row, "微信")
        except Exception:
            wechat = project_info = city_info = ""

        status = "失败"  # 默认失败，确认成功后才改为已录单
        fail_reason = ""
        try:
            try:
                name_input = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "input#input1"))
                )
                name_input.clear()
                name_input.send_keys("女士")
            except Exception as e:
                fail_reason = f"输入称呼失败: {type(e).__name__}: {e}"
                log_error(fail_reason)
            # 分享页手机号输入框 id="input2"
            if phone_number:
                try:
                    phone_input = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input#input2"))
                    )
                    phone_input.clear()
                    phone_input.send_keys(phone_number)
                except Exception as e:
                    log_warning(f"填写手机号失败: {e}")
            try:
                city_label = driver.find_element(By.XPATH, "//p[contains(text(), '* 城市')]")
                city_input = city_label.find_element(By.XPATH, "./following::input[@readonly='readonly'][1]")
                city_input.click()
                city_input.send_keys(city_info)
                time.sleep(1)
                first_city_option = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "li.el-cascader__suggestion-item"))
                )
                first_city_option.click()
                time.sleep(2)
            except Exception as e:
                fail_reason = f"输入城市失败: {type(e).__name__}: {e}"
                log_error(fail_reason)
                if _is_browser_connection_lost(e):
                    raise
            try:
                project_label = driver.find_element(By.XPATH, "//p[contains(text(), '* 项目')]")
                project_input = project_label.find_element(By.XPATH, "./following::input[@readonly='readonly'][1]")
                project_input.click()
                project_input.send_keys(get_project_keyword(project_info))
                time.sleep(1)
                first_project_option = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "li.el-cascader__suggestion-item"))
                )
                first_project_option.click()
                time.sleep(2)
            except Exception as e:
                fail_reason = f"输入项目失败: {type(e).__name__}: {e}"
                log_error(fail_reason)
                if _is_browser_connection_lost(e):
                    raise
            try:
                textarea = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "textarea#input3-0"))
                )
                textarea.clear()
                textarea.send_keys(project_info)
            except Exception as e:
                fail_reason = f"输入咨询内容失败: {type(e).__name__}: {e}"
                log_error(fail_reason)
                if _is_browser_connection_lost(e):
                    raise
            if wechat:
                try:
                    wechat_input = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "input#wechat"))
                    )
                    wechat_input.clear()
                    wechat_input.send_keys(wechat)
                except Exception:
                    pass

            # 点击提交按钮并检测结果
                submit_clicked = False
                try:
                    submit_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '提交')]"))
                    )
                    submit_button.click()
                    time.sleep(2)
                    submit_clicked = True
                    log_info(f"  已点击提交按钮，等待检测成功提示...")
                except Exception:
                    try:
                        submit_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '提交')]"))
                        )
                        submit_button.click()
                        time.sleep(2)
                        submit_clicked = True
                        log_info(f"  已点击提交按钮(备用选择器)，等待检测成功提示...")
                    except Exception as e2:
                        fail_reason = f"点击提交按钮失败: {type(e2).__name__}: {e2}"
                        log_error(fail_reason)
                        status = "失败"

            if submit_clicked:
                # 检测成功提示
                is_success, message = _detect_submit_success(driver, timeout=3)
                if is_success is True:
                    status = "已录单"
                    log_info(f"  检测到成功提示: {message}")
                elif is_success is False:
                    status = "失败"
                    fail_reason = message
                    log_error(f"  检测到错误提示: {message}")
                else:
                    # 未检测到任何提示，保留原状态（可能是成功提示消失太快）
                    log_warning(f"  未检测到成功或错误提示，可能是提示消失太快。当前状态: {status}")
                    if status == "已录单":
                        log_info(f"  由于之前状态为已录单，保留该状态（可能实际已成功）")

                # 刷新回到录单页
                try:
                    driver.get(order_entry_url)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(1)
                except Exception:
                    pass

        except Exception as e:
            fail_reason = f"录单异常: {type(e).__name__}: {e}"
            log_error(f"流程二录单异常: {fail_reason}")
            status = "失败"
            if _is_browser_connection_lost(e):
                log_error("检测到浏览器已断开连接，请重新运行脚本并重新登录。")
                _log_detail(detail_logger, phone_number, "失败", f"浏览器断开: {e}")
                raise
            try:
                driver.get(order_entry_url)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(1)
            except Exception:
                pass

        status_list.append(status)
        _log_detail(detail_logger, phone_number, status, fail_reason)

    _save_feedback_excel(excel_file_path, df, status_list)
    n_ok = sum(1 for s in status_list if s == "已录单")
    n_skip = sum(1 for s in status_list if s == "跳过")
    n_fail = sum(1 for s in status_list if s == "失败")
    log_info("=" * 50)
    log_info(f"【流程二】处理完成 | 共 {total_rows} 行 | 已录单: {n_ok} | 跳过: {n_skip} | 失败: {n_fail}")


def process_jinshuju_excel_new(driver, excel_file_path, flow_choice=2):
    """
    根据用户选择的流程调用对应函数：流程一（重单查询+录单）或流程二（仅录单）。
    """
    global todaytime
    log_info(f"读取 Excel: {excel_file_path}")

    try:
        df = pd.read_excel(excel_file_path, header=0)
        log_info(f"  成功读取，共 {len(df)} 行，表头: {list(df.columns)}")
    except Exception as e:
        log_error(f"读取Excel文件失败: {e}")
        return

    try:
        log_info("将浏览器窗口置于最前端...")
        if not force_window_to_foreground(driver):
            log_warning("  浏览器窗口置前可能未完全成功，继续执行")
    except Exception as e:
        log_error(f"设置浏览器窗口状态时出错: {e}")

    if flow_choice == 1:
        log_info("执行流程一（重单查询 + 录单）")
        run_flow1_duplicate_check_and_entry(driver, excel_file_path, df)
    else:
        log_info("执行流程二（分享页录单）")
        run_flow2_entry_only(driver, excel_file_path, df)

    log_info("Excel 录单处理结束")


def launch_new_browser(flow_choice, chromedriver_path=None):
    """
    按用户选择的流程打开对应网址。
    重单查询页 duplicate_check_url 需要等待用户登录确认；流程二的分享页 order_entry_url 不需要登录可直接使用。
    """
    try:
        opts = Options()
        # 降低 Chrome 控制台无关日志（GCM、DEPRECATED_ENDPOINT、Authentication Failed 等）
        opts.add_argument("--log-level=3")
        opts.add_argument("--disable-background-networking")
        opts.add_argument("--disable-sync")
        opts.add_argument("--disable-default-apps")
        opts.add_argument("--metrics-recording-only")
        opts.add_argument("--no-first-run")
        if chromedriver_path and os.path.exists(chromedriver_path):
            from selenium.webdriver.chrome.service import Service
            service = Service(executable_path=chromedriver_path, log_path=os.devnull)
            driver = webdriver.Chrome(service=service, options=opts)
        else:
            from selenium.webdriver.chrome.service import Service
            service = Service(log_path=os.devnull)
            driver = webdriver.Chrome(service=service, options=opts)
        driver.set_page_load_timeout(30)
        driver.implicitly_wait(5)
        log_info("正在打开重单查询页面（需登录确认）...")
        driver.get(DUPLICATE_CHECK_URL)
        time.sleep(1)
        if flow_choice == 2:
            log_info("正在打开流程二录单页面（无需登录，可直接使用）...")
            driver.switch_to.new_window("tab")
            driver.get(ORDER_ENTRY_URL)
            log_info("重单查询页与录单页已打开。请在重单查询页完成登录和确认后，回到本窗口按回车继续。")
        else:
            log_info("当前为流程一，仅打开重单查询页。请在该页完成登录和确认后，回到本窗口按回车继续。")
        return driver
    except Exception as e:
        log_error(f"启动浏览器失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def ask_user_select_flow():
    """程序启动后第一时间让用户选择录单流程，返回 1 或 2。"""
    print("\n请选择录单流程：")
    print("  1 - 流程一（登录分销后台录单，可以查重）")
    print("  2 - 流程二（默认表单页录单，无法查重）")
    while True:
        try:
            choice = input("请输入 1 或 2 后按回车: ").strip()
            if choice == "1":
                return 1
            if choice == "2":
                return 2
        except (EOFError, KeyboardInterrupt):
            return None
        print("无效输入，请重新输入 1 或 2。")


def ask_user_select_excel():
    """弹出文件选择对话框，让用户选择要录单的 xlsx 文件。"""
    if not TKINTER_AVAILABLE:
        log_error("未安装 tkinter，无法使用文件选择对话框。请通过命令行传入: python fx_ludan.py 你的文件.xlsx")
        return None
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="选择要录单的 Excel 文件",
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
    )
    root.destroy()
    return path if path and os.path.isfile(path) else None


def main():
    global todaytime
    if getattr(sys, "frozen", False):
        try:
            os.chdir(_exe_dir())
        except Exception:
            pass
    setup_logging()
    todaytime = datetime.now().strftime('%Y%m%d')
    log_info("========== fx_ludan 分销录单脚本 启动 ==========")

    # 步骤 1/5：选择录单流程
    log_info("【步骤 1/5】选择录单流程")
    flow_choice = 1
    '''
    flow_choice = ask_user_select_flow()
    if flow_choice is None:
        log_info("用户取消选择，退出。")
        return
    log_info(f"  已选择: 流程 {flow_choice}")
    '''
    # 步骤 2/5：启动浏览器并打开页面
    log_info("【步骤 2/5】启动浏览器并打开页面")
    driver = launch_new_browser(flow_choice, chromedriver_path=CHROMEDRIVER_PATH)
    if not driver:
        log_error("无法启动浏览器，程序退出。")
        return

    # 步骤 3/5：等待用户登录
    log_info("【步骤 3/5】等待在重单查询页完成登录")
    try:
        input("\n请在重单查询页完成登录和确认后，回到本窗口按 回车 继续...\n")
    except (EOFError, KeyboardInterrupt):
        log_info("用户取消，退出。")
        return

    # 步骤 4/5：选择 Excel 文件
    log_info("【步骤 4/5】选择要录单的 Excel 文件")
    excel_path = ask_user_select_excel()
    if not excel_path:
        log_error("未选择有效的 Excel 文件，程序退出。")
        return
    log_info(f"  已选择: {excel_path}")

    # 步骤 5/5：执行录单
    log_info("【步骤 5/5】开始执行录单")
    log_info("=" * 50)
    try:
        process_jinshuju_excel_new(driver, excel_path, flow_choice)
    except Exception as e:
        if _is_browser_connection_lost(e):
            log_error("浏览器已断开连接，录单已中止。请关闭浏览器后重新运行脚本。")
        else:
            raise
        return
    log_info("=" * 50)
    log_info("========== fx_ludan 执行完毕 ==========")


def _save_feedback_excel_safe(excel_file_path, df, status_list):
    """反馈 Excel 输出时避免同一天重复运行互相覆盖。"""
    if len(status_list) != len(df):
        log_warning("状态条数与表格行数不一致，反馈表可能不完整")
    df_feedback = df.copy()
    df_feedback["状态"] = status_list
    out_dir = os.path.dirname(os.path.abspath(excel_file_path))
    date_str = datetime.now().strftime("%Y%m%d")
    base_name = f"录单反馈+{date_str}"

    existing = set()
    if os.path.isdir(out_dir):
        existing = {
            f for f in os.listdir(out_dir)
            if f.startswith(base_name) and f.endswith(".xlsx")
        }

    num = 1
    while num < 9999:
        filename = f"{base_name}.xlsx" if num == 1 else f"{base_name}_{num}.xlsx"
        if filename not in existing:
            break
        num += 1

    out_path = os.path.join(out_dir, filename)
    try:
        df_feedback.to_excel(out_path, index=False)
        log_info(f"录单反馈已保存: {out_path}")
    except Exception as e:
        log_error(f"保存录单反馈表失败: {e}")


def _detect_submit_success_safe(driver, timeout=3):
    """
    更严格的提交结果判断：只有出现明确成功语义时才返回 True，
    避免把页面上的“提交”按钮或普通文本误判为提交成功。
    """
    success_selectors = [
        "div.el-message--success",
        "div.el-message.el-message--success",
        "div.el-notification--success",
        "div.el-notification.success",
        "//div[contains(.,'提交成功')]",
        "//span[contains(.,'提交成功')]",
        "//p[contains(.,'提交成功')]",
        "div.toast-success",
        "div.van-toast--success",
        "div.weui-toast",
    ]

    error_selectors = [
        "div.el-message--error",
        "div.el-message.el-message--error",
        "div.el-notification.error",
        "//*[contains(@class,'error') and contains(.,'失败')]",
        "//div[contains(.,'提交失败')]",
        "//div[contains(.,'错误')]",
        "div.el-form-item__error",
    ]

    success_keywords = ("成功", "提交成功", "保存成功", "录入成功", "已提交")
    error_keywords = ("失败", "错误", "无法")

    start_time = time.time()
    while time.time() - start_time < timeout:
        for selector in success_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector) if not selector.startswith("//") else driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    text = (elem.text or elem.get_attribute("textContent") or "").strip()
                    classes = (elem.get_attribute("class") or "").lower()
                    tag_name = (elem.tag_name or "").lower()
                    if not text:
                        continue
                    if tag_name == "button" or "el-button" in classes or "button" in classes:
                        continue
                    if any(keyword in text for keyword in success_keywords):
                        return True, text
            except Exception:
                pass

        for selector in error_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector) if not selector.startswith("//") else driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    text = (elem.text or elem.get_attribute("textContent") or "").strip()
                    if text and any(keyword in text for keyword in error_keywords):
                        return False, f"错误提示: {text}"
            except Exception:
                pass

        time.sleep(0.2)

    return None, "未检测到明确的成功或错误提示"


_detect_submit_success = _detect_submit_success_safe
_save_feedback_excel = _save_feedback_excel_safe


_current_failure_list_path = None


def _next_indexed_output_path(out_dir, base_name, ext):
    existing = set()
    if os.path.isdir(out_dir):
        existing = {
            f for f in os.listdir(out_dir)
            if f.startswith(base_name) and f.endswith(ext)
        }
    num = 1
    while num < 9999:
        filename = f"{base_name}{ext}" if num == 1 else f"{base_name}_{num}{ext}"
        if filename not in existing:
            return os.path.join(out_dir, filename)
        num += 1
    return os.path.join(out_dir, f"{base_name}{ext}")


def _create_detail_logger_safe(excel_file_path):
    global _current_failure_list_path
    out_dir = os.path.dirname(os.path.abspath(excel_file_path))
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    log_path = os.path.join(out_dir, f"日志{timestamp}.txt")

    detail_logger = logging.getLogger(f"detail_{timestamp}")
    detail_logger.setLevel(logging.INFO)
    for handler in detail_logger.handlers[:]:
        try:
            handler.close()
        except Exception:
            pass
        detail_logger.removeHandler(handler)

    file_handler = logging.FileHandler(log_path, encoding="utf-8", mode="w")
    formatter = logging.Formatter("%(asctime)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    file_handler.setFormatter(formatter)
    detail_logger.addHandler(file_handler)

    date_str = datetime.now().strftime("%Y%m%d")
    _current_failure_list_path = _next_indexed_output_path(out_dir, f"失败列表+{date_str}", ".txt")
    try:
        with open(_current_failure_list_path, "w", encoding="utf-8") as f:
            f.write("=" * 50 + "\n")
            f.write(f"失败列表 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n")
    except Exception as e:
        log_error(f"保存失败列表文件失败: {e}")
        _current_failure_list_path = None

    return detail_logger, log_path


def _append_failure_item(phone, reason):
    if not _current_failure_list_path:
        return
    try:
        with open(_current_failure_list_path, "a", encoding="utf-8") as f:
            f.write(f"号码: {phone}\n")
            f.write(f"失败原因: {reason or '未识别到具体原因'}\n")
            f.write("-" * 50 + "\n")
    except Exception as e:
        log_error(f"写入失败列表失败: {e}")


def _log_detail_safe(detail_logger, phone, status, reason=""):
    if detail_logger:
        if status == "已录单":
            detail_logger.info(f"号码: {phone} | 状态: {status}")
        elif status == "失败":
            detail_logger.info(f"号码: {phone} | 状态: {status} | 失败原因: {reason}")
            _append_failure_item(phone, reason)
        elif status == "重单":
            detail_logger.info(f"号码: {phone} | 状态: {status} | 原因: 系统中已存在该客户")
        elif status == "跳过":
            detail_logger.info(f"号码: {phone} | 状态: {status} | 原因: 号码为空")
        else:
            detail_logger.info(f"号码: {phone} | 状态: {status} | 备注: {reason}")


def _detect_submit_success_balanced(driver, timeout=2):
    """
    先做 2 秒明确提示检测，再做一次轻量页面校验。
    轻量校验只在表单明显被重置时判定成功，避免误把提交按钮文本当成成功。
    """
    success_selectors = [
        "div.el-message--success",
        "div.el-message.el-message--success",
        "div.el-notification--success",
        "div.el-notification.success",
        "//div[contains(.,'提交成功')]",
        "//span[contains(.,'提交成功')]",
        "//p[contains(.,'提交成功')]",
        "div.toast-success",
        "div.van-toast--success",
        "div.weui-toast",
    ]
    error_selectors = [
        "div.el-message--error",
        "div.el-message.el-message--error",
        "div.el-notification.error",
        "//*[contains(@class,'error') and contains(.,'失败')]",
        "//div[contains(.,'提交失败')]",
        "//div[contains(.,'错误')]",
        "div.el-form-item__error",
    ]
    success_keywords = ("成功", "提交成功", "保存成功", "录入成功", "已提交")
    error_keywords = ("失败", "错误", "无法")

    start_time = time.time()
    while time.time() - start_time < timeout:
        for selector in success_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector) if not selector.startswith("//") else driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    text = (elem.text or elem.get_attribute("textContent") or "").strip()
                    classes = (elem.get_attribute("class") or "").lower()
                    tag_name = (elem.tag_name or "").lower()
                    if not text:
                        continue
                    if tag_name == "button" or "el-button" in classes or "button" in classes:
                        continue
                    if any(keyword in text for keyword in success_keywords):
                        return True, text
            except Exception:
                pass

        for selector in error_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector) if not selector.startswith("//") else driver.find_elements(By.XPATH, selector)
                for elem in elements:
                    text = (elem.text or elem.get_attribute("textContent") or "").strip()
                    if text and any(keyword in text for keyword in error_keywords):
                        return False, f"错误提示: {text}"
            except Exception:
                pass
        time.sleep(0.2)

    try:
        body_text = (driver.find_element(By.TAG_NAME, "body").text or "").strip()
    except Exception:
        body_text = ""
    for keyword in ("提交失败", "失败", "错误", "无法"):
        if keyword in body_text:
            return False, f"错误提示: {keyword}"

    candidate_selectors = [
        "input[placeholder*='手机号']",
        "input[placeholder*='手机号码']",
        "input#input2",
        "textarea#input3-0",
        "textarea[placeholder*='客户情况']",
        "textarea[placeholder*='说明客户情况']",
    ]
    reset_hits = 0
    for selector in candidate_selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
        except Exception:
            elements = []
        for elem in elements:
            try:
                value = (elem.get_attribute("value") or "").strip()
                if value == "":
                    reset_hits += 1
            except Exception:
                pass
    if reset_hits >= 2:
        return True, f"页面轻量校验通过，检测到 {reset_hits} 个关键字段已重置"

    return None, "未检测到明确成功提示，且页面未呈现明显重置状态"


_create_detail_logger = _create_detail_logger_safe
_log_detail = _log_detail_safe
_detect_submit_success = _detect_submit_success_balanced


if __name__ == "__main__":
    main()
