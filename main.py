import sys
import os
import json
import datetime
import time
import asyncio
from urllib.parse import quote_plus

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QCheckBox, QPushButton, QTableWidget, QTableWidgetItem,
    QStatusBar, QProgressBar, QSizePolicy, QHeaderView, QMessageBox,
    QFileDialog, QShortcut, QMenu, QAction, QLabel
)
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEnginePage, QWebEngineProfile, QWebEngineSettings
from PyQt5.QtCore import Qt, QUrl, QTimer, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QKeySequence, QPixmap

# --- Cờ Debug ---
DEBUG_WORKER = False # True để bật log chi tiết từ Worker
DEBUG_MAIN_THREAD_HANDLERS = False # True để bật log từ các handler trên luồng chính

try:
    import openpyxl
    from openpyxl.styles import Font
except ImportError:
    QMessageBox.critical(None, "Lỗi thư viện",
                         "Thư viện openpyxl là bắt buộc nhưng không tìm thấy. "
                         "Vui lòng cài đặt (pip install openpyxl) và thử lại.")
    sys.exit(1)

# --- Các lớp Helper ---
class SearchResult:
    def __init__(self, website="", keyword="", link="", found="Không", full_data="", full_data_summary=""):
        self.website = website
        self.keyword = keyword
        self.link = link
        self.found = found
        self.full_data = full_data
        self.full_data_summary = full_data_summary

# --- Bộ tạo URL Tìm kiếm ---
class SearchEngines: # (Giữ nguyên như trước)
    @staticmethod
    def get_instagram_search_url(keyword):
        return f"https://www.instagram.com/explore/tags/{quote_plus(keyword.replace(' ', ''))}/"
    @staticmethod
    def get_facebook_top_search_url(keyword):
        return f"https://www.facebook.com/search/top/?q={quote_plus(keyword)}"
    @staticmethod
    def get_facebook_videos_search_url(keyword):
        return f"https://www.facebook.com/search/videos/?q={quote_plus(keyword)}"
    @staticmethod
    def get_facebook_posts_search_url(keyword):
        return f"https://www.facebook.com/search/posts/?q={quote_plus(keyword)}"
    @staticmethod
    def get_nike_search_url(keyword):
        return f"https://www.nike.com/vn/search?q={quote_plus(keyword)}"
    @staticmethod
    def get_x_top_search_url(keyword):
        return f"https://x.com/search?q={quote_plus(keyword)}&src=typed_query"
    @staticmethod
    def get_x_media_search_url(keyword):
        return f"https://x.com/search?q={quote_plus(keyword)}&src=typed_query&f=media"
    @staticmethod
    def get_sneaker_news_search_url(keyword):
        return f"https://sneakernews.com/?s={quote_plus(keyword)}"

# --- Search Worker ---
class SearchWorker(QObject):
    finished_task = pyqtSignal(str, str, str, str, str)
    progress_update = pyqtSignal(int, str)
    search_completed_all = pyqtSignal(str)

    def __init__(self, tasks_to_run, parent_app_ref):
        super().__init__()
        self.tasks_to_run = tasks_to_run
        self.parent_app = parent_app_ref
        self._is_running = True
        self.async_loop = None

    def stop(self):
        if DEBUG_WORKER: print("WORKER: stop() called.")
        self._is_running = False
        if self.async_loop and self.async_loop.is_running():
            if DEBUG_WORKER: print("WORKER: Async loop is running, attempting to cancel tasks.")
            # Hủy tất cả các tác vụ đang chạy trong vòng lặp này
            # Ngoại trừ chính tác vụ đang chạy (nếu stop được gọi từ bên trong một tác vụ)
            for task in asyncio.all_tasks(self.async_loop):
                if task is not asyncio.current_task(self.async_loop):
                    if DEBUG_WORKER: print(f"WORKER: Requesting cancellation for task: {task.get_name()}")
                    task.cancel()

    async def _safe_navigate_and_wait(self, url_string, timeout_sec=30):
        if not self._is_running: raise asyncio.CancelledError("Search stopped during navigation prep.")
        qurl = QUrl(url_string)
        future = asyncio.Future()
        callback_id = id(future)
        if DEBUG_WORKER: print(f"WORKER: Navigating to {url_string}. Future ID: {callback_id}")

        # Hàm callback này sẽ được gọi bởi luồng chính
        def future_resolver_callback(success):
            if DEBUG_WORKER: print(f"WORKER_CB (Nav for {url_string}, FutureID {callback_id}): Received {success}. Future done? {future.done()}")
            if not future.done():
                if success:
                    future.set_result(True)
                else:
                    future.set_exception(ConnectionError(f"Navigation failed for {url_string}"))
            # else:
            #     if DEBUG_WORKER: print(f"WORKER_CB (Nav for {url_string}, FutureID {callback_id}): Future was already done.")

        self.parent_app.request_load_url_from_worker.emit(qurl, future_resolver_callback, callback_id)
        try:
            await asyncio.wait_for(future, timeout=timeout_sec)
            if DEBUG_WORKER: print(f"WORKER: Navigation to {url_string} (FutureID {callback_id}) successful.")
        except asyncio.TimeoutError:
            if DEBUG_WORKER: print(f"WORKER: Navigation to {url_string} (FutureID {callback_id}) TIMED OUT.")
            # Đảm bảo future được giải quyết ngay cả khi timeout từ wait_for
            if not future.done(): future.set_exception(TimeoutError(f"Navigation timed out for {url_string}"))
            self.parent_app.request_future_timeout_from_worker.emit(callback_id) # Báo cho main thread biết future này đã timeout
            raise
        except ConnectionError as ce:
            if DEBUG_WORKER: print(f"WORKER: Navigation to {url_string} (FutureID {callback_id}) FAILED (ConnectionError).")
            if not future.done(): future.set_exception(ce) # Thường thì future đã set exception trong callback
            raise
        except asyncio.CancelledError:
            if DEBUG_WORKER: print(f"WORKER: Navigation to {url_string} (FutureID {callback_id}) CANCELLED.")
            if not future.done(): future.set_exception(asyncio.CancelledError("Navigation cancelled by stop signal"))
            self.parent_app.request_future_timeout_from_worker.emit(callback_id) # Báo cho main thread hủy callback nếu cần
            raise
        except Exception as e:
            if DEBUG_WORKER: print(f"WORKER: Navigation to {url_string} (FutureID {callback_id}) FAILED with other exception: {e}")
            if not future.done(): future.set_exception(e)
            raise

    async def _safe_execute_script_and_wait(self, script, timeout_sec=15):
        if not self._is_running: raise asyncio.CancelledError("Search stopped during JS exec prep.")
        future = asyncio.Future()
        callback_id = id(future)
        if DEBUG_WORKER: print(f"WORKER: Executing JS (first 60 chars): {script[:60]}. Future ID: {callback_id}")

        def future_resolver_callback(result):
            if DEBUG_WORKER: print(f"WORKER_CB (JS Result, FutureID {callback_id}): Result type {type(result)}. Future done? {future.done()}")
            if not future.done():
                future.set_result(result) # JS có thể trả về None, đó là kết quả hợp lệ
            # else:
            #     if DEBUG_WORKER: print(f"WORKER_CB (JS Result, FutureID {callback_id}): Future was already done.")

        self.parent_app.request_execute_js_from_worker.emit(script, future_resolver_callback, callback_id)
        try:
            result = await asyncio.wait_for(future, timeout=timeout_sec)
            if DEBUG_WORKER: print(f"WORKER: JS Execution (FutureID {callback_id}) successful.")
            return result
        except asyncio.TimeoutError:
            if DEBUG_WORKER: print(f"WORKER: JS Execution (FutureID {callback_id}) TIMED OUT.")
            if not future.done(): future.set_exception(TimeoutError("JavaScript execution timed out"))
            self.parent_app.request_future_timeout_from_worker.emit(callback_id)
            raise
        except asyncio.CancelledError:
            if DEBUG_WORKER: print(f"WORKER: JS Execution (FutureID {callback_id}) CANCELLED.")
            if not future.done(): future.set_exception(asyncio.CancelledError("JS execution cancelled by stop signal"))
            self.parent_app.request_future_timeout_from_worker.emit(callback_id)
            raise
        except Exception as e: # Bắt các lỗi khác có thể xảy ra khi chờ future
            if DEBUG_WORKER: print(f"WORKER: JS Execution (FutureID {callback_id}) FAILED with other exception: {e}")
            if not future.done(): future.set_exception(e)
            raise

    def _clean_js_result(self, js_result): # (Giữ nguyên)
        if js_result is None: return ""
        if isinstance(js_result, str):
            if js_result.startswith("\"") and js_result.endswith("\"") and len(js_result) > 1:
                js_result = js_result[1:-1]
            js_result = (js_result.replace("\\n", "\n")
                         .replace("\\r", "\r")
                         .replace("\\t", "\t")
                         .replace("\\\"", "\"")
                         .replace("\\\\", "\\"))
        return js_result

    async def _perform_instagram_search(self, keyword):
        if not self._is_running: raise asyncio.CancelledError()
        self.progress_update.emit(0, f"Instagram: Bắt đầu tìm '{keyword}'...")
        base_tag_url = SearchEngines.get_instagram_search_url(keyword)
        found_links_for_keyword = set()

        try:
            await self._safe_navigate_and_wait(base_tag_url)
            await asyncio.sleep(self.parent_app.config_wait_medium) # Chờ trang tag tải + render

            for i in range(self.parent_app.config_insta_scrolls):
                if not self._is_running: raise asyncio.CancelledError()
                if DEBUG_WORKER: print(f"WORKER (Insta): Scrolling tag page ({i+1}/{self.parent_app.config_insta_scrolls}) for '{keyword}'")
                await self._safe_execute_script_and_wait("window.scrollTo(0, document.body.scrollHeight);")
                await asyncio.sleep(self.parent_app.config_wait_medium) # Chờ tải thêm

            js_get_post_links = """ ( /* Giữ nguyên như trước, nhưng đảm bảo trả về [] nếu lỗi */
                (function(){
                    var links = new Set(); /* Dùng Set để tự loại bỏ trùng lặp */
                    var selectors = [
                        'main div article a[href*="/p/"]', 'main div article a[href*="/reel/"]',
                        'main div section a[href*="/p/"]', 'main div section a[href*="/reel/"]',
                        'div[role="main"] a[href*="/p/"]', 'div[role="main"] a[href*="/reel/"]',
                        'a[href*="/p/"]', 'a[href*="/reel/"]' /* Selector chung hơn cuối cùng */
                    ];
                    for (var sel of selectors) {
                        try {
                            var nodes = document.querySelectorAll(sel);
                            for(var i = 0; i < nodes.length; i++){
                                if(nodes[i].href) { links.add(nodes[i].href); }
                            }
                        } catch (e) { /* console.error('Error with selector: ' + sel, e); */ }
                    }
                    return JSON.stringify(Array.from(links).slice(0, 12)); /* Giới hạn 12 links */
                })();
            )"""
            
            post_links_json = await self._safe_execute_script_and_wait(js_get_post_links)
            cleaned_json_str = self._clean_js_result(post_links_json)
            post_links_raw = []
            if cleaned_json_str:
                try:
                    parsed_json = json.loads(cleaned_json_str)
                    if isinstance(parsed_json, list):
                        post_links_raw = parsed_json
                    else:
                        if DEBUG_WORKER: print(f"WORKER (Insta): Parsed JSON for post links is not a list: {parsed_json}")
                except json.JSONDecodeError as json_e:
                    if DEBUG_WORKER: print(f"WORKER (Insta): JSONDecodeError for post links: {json_e}. String: '{cleaned_json_str}'")
            
            if not post_links_raw:
                self.finished_task.emit("Instagram", keyword, base_tag_url, "Không (no posts found)", "Không tìm thấy link bài đăng nào trên trang tag.")
                return

            self.progress_update.emit(10, f"Instagram: Tìm thấy {len(post_links_raw)} link bài đăng cho '{keyword}'.")
            
            item_processed_count = 0
            for post_url in post_links_raw:
                if not self._is_running: raise asyncio.CancelledError()
                if post_url in found_links_for_keyword: continue
                
                item_processed_count +=1
                self.progress_update.emit(10 + int(item_processed_count * 80 / len(post_links_raw)),
                                          f"Instagram: Xử lý bài đăng {item_processed_count}/{len(post_links_raw)}...")
                try:
                    await self._safe_navigate_and_wait(post_url)
                    await asyncio.sleep(self.parent_app.config_wait_medium)

                    js_get_caption = """ ( /* Giữ nguyên hoặc cải thiện selector cho caption */
                        (function() {
                            let caption = "";
                            const selectors = [
                                'h1', // Thường là caption chính
                                'article div[role="button"] + div ul li span', // Bình luận đầu tiên
                                'article div[role="dialog"] div[role="dialog"] ul li span',
                                'span[dir="auto"]', // Các span có text
                                'div[dir="auto"]'
                            ];
                            for (let sel of selectors) {
                                try {
                                    let elements = document.querySelectorAll(sel);
                                    for (let el of elements) {
                                        if (el.innerText && el.innerText.trim().length > 20) { // Ưu tiên text dài và có nghĩa
                                            caption = el.innerText.trim();
                                            // console.log('Found caption with selector: ' + sel, caption);
                                            return caption; // Trả về ngay khi tìm thấy
                                        }
                                    }
                                } catch (e) { /* console.error('Error with selector: ' + sel, e); */ }
                            }
                            // Fallback cuối cùng nếu không tìm thấy gì cụ thể
                            if (!caption && document.body) caption = document.body.innerText;
                            return caption;
                        })();
                    )"""
                    post_text_raw = await self._safe_execute_script_and_wait(js_get_caption)
                    post_text = self._clean_js_result(post_text_raw)

                    found_status = "Có" if keyword.lower() in post_text.lower() else "Không"
                    self.finished_task.emit("Instagram", keyword, post_url, found_status, post_text)
                    found_links_for_keyword.add(post_url)

                    if found_status == "Có" and self.parent_app.cb_long_screenshot.isChecked():
                        if not self._is_running: raise asyncio.CancelledError()
                        self.parent_app.request_scroll_to_keyword_from_worker.emit(keyword)
                        await asyncio.sleep(1.5)
                        if not self._is_running: raise asyncio.CancelledError()
                        self.parent_app.request_capture_screenshot_from_worker.emit("Instagram", keyword, post_url)
                        await asyncio.sleep(1.0)
                
                except asyncio.CancelledError: raise # Lan truyền CancelledError
                except TimeoutError as te_post:
                    self.finished_task.emit("Instagram", keyword, post_url, "Lỗi (Timeout Post)", str(te_post))
                except ConnectionError as ce_post:
                    self.finished_task.emit("Instagram", keyword, post_url, "Lỗi (Nav Post)", str(ce_post))
                except Exception as e_post:
                    self.finished_task.emit("Instagram", keyword, post_url, "Lỗi (Post)", f"{type(e_post).__name__}: {str(e_post)}")
                
                if not self._is_running: raise asyncio.CancelledError()
                await asyncio.sleep(self.parent_app.config_wait_short) # Nghỉ giữa các bài đăng
        
        except asyncio.CancelledError:
            if DEBUG_WORKER: print(f"WORKER (Insta): Search for '{keyword}' CANCELLED.")
            self.finished_task.emit("Instagram", keyword, base_tag_url, "Bị hủy", "Tìm kiếm Instagram bị hủy.")
            # Không raise lại ở đây, _run_all_tasks_async sẽ xử lý
        except TimeoutError as te:
            self.finished_task.emit("Instagram", keyword, base_tag_url, "Lỗi (Timeout)", str(te))
        except ConnectionError as ce:
            self.finished_task.emit("Instagram", keyword, base_tag_url, "Lỗi (Navigation)", str(ce))
        except Exception as e:
            self.finished_task.emit("Instagram", keyword, base_tag_url, "Lỗi (Chung)", f"{type(e).__name__}: {str(e)}")


    async def _perform_generic_search(self, url_generator, website_name, keyword):
        if not self._is_running: raise asyncio.CancelledError()
        self.progress_update.emit(0, f"{website_name}: Bắt đầu tìm '{keyword}'...")
        search_url = url_generator(keyword)
        try:
            await self._safe_navigate_and_wait(search_url)
            await asyncio.sleep(self.parent_app.config_wait_medium)

            # ... (logic cuộn trang và xóa H1 như cũ) ...
            if "Facebook" in website_name or "Nike" in website_name:
                # ... (JS xóa H1) ...
                await self._safe_execute_script_and_wait("/* JS xóa H1 */")


            scroll_count = 0
            if "Facebook" in website_name or "X" in website_name:
                scroll_count = self.parent_app.config_fb_x_scrolls
            elif "Nike" in website_name: scroll_count = 1
            
            for i in range(scroll_count):
                if not self._is_running: raise asyncio.CancelledError()
                # ... (logic cuộn) ...
                await self._safe_execute_script_and_wait(f"window.scrollBy(0, window.innerHeight * {0.7 + i*0.1});")
                await asyncio.sleep(self.parent_app.config_wait_short)


            js_get_data = "document.body.innerText;"
            full_data_raw = await self._safe_execute_script_and_wait(js_get_data)
            full_data = self._clean_js_result(full_data_raw if full_data_raw else "")

            found_status = "Có" if keyword.lower() in full_data.lower() else "Không"
            self.finished_task.emit(website_name, keyword, search_url, found_status, full_data)

            if found_status == "Có" and self.parent_app.cb_long_screenshot.isChecked():
                if not self._is_running: raise asyncio.CancelledError()
                self.parent_app.request_scroll_to_keyword_from_worker.emit(keyword)
                await asyncio.sleep(1.5)
                if not self._is_running: raise asyncio.CancelledError()
                self.parent_app.request_capture_screenshot_from_worker.emit(website_name, keyword, search_url)
                await asyncio.sleep(1.0)

        except asyncio.CancelledError:
            if DEBUG_WORKER: print(f"WORKER (Generic: {website_name}): Search for '{keyword}' CANCELLED.")
            self.finished_task.emit(website_name, keyword, search_url, "Bị hủy", f"Tìm kiếm {website_name} bị hủy.")
        except TimeoutError as te:
            self.finished_task.emit(website_name, keyword, search_url, "Lỗi (Timeout)", str(te))
        except ConnectionError as ce:
            self.finished_task.emit(website_name, keyword, search_url, "Lỗi (Navigation)", str(ce))
        except Exception as e:
            self.finished_task.emit(website_name, keyword, search_url, "Lỗi (Chung)", f"{type(e).__name__}: {str(e)}")

    async def _run_all_tasks_async(self):
        if DEBUG_WORKER: print("WORKER: _run_all_tasks_async started.")
        total_tasks_overall = len(self.tasks_to_run)
        current_task_num_overall = 0

        for task_info in self.tasks_to_run:
            if not self._is_running:
                if DEBUG_WORKER: print("WORKER: Loop detected _is_running=False, breaking.")
                self.progress_update.emit(100, "Tìm kiếm bị dừng bởi người dùng.")
                break
            
            current_task_num_overall += 1
            progress_percent = int(current_task_num_overall * 100 / total_tasks_overall)
            
            keyword = task_info["keyword"]
            platform_name = task_info["platform_name"]
            url_generator = task_info.get("url_generator")

            self.progress_update.emit(progress_percent,
                                      f"Đang xử lý {platform_name} cho '{keyword}' ({current_task_num_overall}/{total_tasks_overall})")
            try:
                if platform_name == "Instagram":
                    await self._perform_instagram_search(keyword)
                else:
                    if url_generator:
                        await self._perform_generic_search(url_generator, platform_name, keyword)
                    else:
                        self.finished_task.emit(platform_name, keyword, "N/A", "Lỗi (Config)", "URL generator không được định nghĩa")
            except asyncio.CancelledError:
                if DEBUG_WORKER: print(f"WORKER: Main task loop caught CancelledError for {platform_name} - '{keyword}'. Breaking.")
                # progress_update đã được phát từ hàm con hoặc ở lần lặp tiếp theo
                break 
            except Exception as e_task_run:
                if DEBUG_WORKER: print(f"WORKER: Main task loop: Unexpected error for {platform_name} - '{keyword}': {e_task_run}")
                self.finished_task.emit(platform_name, keyword, "N/A", "Lỗi (Worker)", f"Lỗi không xác định: {e_task_run}")
            
            if not self._is_running: # Kiểm tra lại sau mỗi task lớn
                if DEBUG_WORKER: print("WORKER: Loop detected _is_running=False after task. Breaking.")
                break
            
            await asyncio.sleep(0.1) 

        if self._is_running: # Nếu vòng lặp hoàn thành tự nhiên
            self.search_completed_all.emit(f"Hoàn tất tìm kiếm. Đã xử lý {current_task_num_overall} tác vụ.")
        else: # Nếu vòng lặp bị ngắt do _is_running = False
            self.search_completed_all.emit(f"Tìm kiếm đã dừng. Đã xử lý {current_task_num_overall} tác vụ.")
        if DEBUG_WORKER: print("WORKER: _run_all_tasks_async finished.")

    def run(self):
        if DEBUG_WORKER: print(f"WORKER: run() method started. Thread ID: {QThread.currentThreadId()}")
        self.async_loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.async_loop)
        try:
            self.async_loop.run_until_complete(self._run_all_tasks_async())
        except Exception as e_loop:
            if DEBUG_WORKER: print(f"WORKER: Critical error in asyncio event loop: {e_loop}")
            self.search_completed_all.emit(f"Lỗi nghiêm trọng trong luồng worker: {e_loop}")
        finally:
            try:
                if DEBUG_WORKER: print("WORKER: Asyncio loop completed or errored. Cleaning up...")
                # Hủy các tác vụ còn lại (nếu có, ví dụ do lỗi không mong muốn làm loop.run_until_complete thoát sớm)
                if self.async_loop.is_running(): # Hiếm khi xảy ra nếu run_until_complete đã chạy
                    pending_tasks = [task for task in asyncio.all_tasks(self.async_loop) if not task.done()]
                    if pending_tasks:
                        if DEBUG_WORKER: print(f"WORKER: Found {len(pending_tasks)} pending tasks during final cleanup. Cancelling them.")
                        for task in pending_tasks: task.cancel()
                        # Chạy loop một lần nữa để xử lý việc hủy
                        self.async_loop.run_until_complete(asyncio.gather(*pending_tasks, return_exceptions=True))
            except Exception as e_final_cleanup:
                if DEBUG_WORKER: print(f"WORKER: Error during final asyncio cleanup: {e_final_cleanup}")
            
            self.async_loop.close()
            self.async_loop = None 
            if DEBUG_WORKER: print(f"WORKER: run() method finished. Thread ID: {QThread.currentThreadId()}")

# --- Cửa sổ Ứng dụng Chính ---
class KeywordSearchApp(QMainWindow):
    request_load_url_from_worker = pyqtSignal(QUrl, object, int) # QUrl, future_resolver_callback, callback_id
    request_execute_js_from_worker = pyqtSignal(str, object, int) # script, future_resolver_callback, callback_id
    request_scroll_to_keyword_from_worker = pyqtSignal(str)
    request_capture_screenshot_from_worker = pyqtSignal(str, str, str)
    request_future_timeout_from_worker = pyqtSignal(int) # callback_id of the future that timed out

    def __init__(self):
        super().__init__()
        # ... (Cấu hình UI và các biến như trước) ...
        self.setWindowTitle("Keyword Search Tool - Python/PyQt5 (v2 - Refined)")
        self.setGeometry(50, 50, 1350, 900)

        self.config_wait_short = 2.0 # Giảm nhẹ thời gian chờ
        self.config_wait_medium = 3.5
        self.config_insta_scrolls = 2
        self.config_fb_x_scrolls = 2

        self.results_list = []
        self.search_worker_obj = None
        self.search_thread = None
        self.stopwatch_start_time = None
        self.elapsed_timer_display = QTimer(self)
        self.elapsed_timer_display.setInterval(1000)
        self.elapsed_timer_display.timeout.connect(self.update_elapsed_time_display_label)
        self.current_search_tasks_list = []
        
        # Dictionary để lưu trữ các hàm callback đang chờ kết quả từ WebEngine
        # Key: callback_id (int), Value: future_resolver_callback (function)
        self.active_web_callbacks = {}

        self._init_ui()
        self._setup_connections()
        self._setup_shortcuts()
        self.load_keywords_from_file()
        self.update_status("Ứng dụng đã khởi tạo. Sẵn sàng.")
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: App initialized. Main Thread ID: {QThread.currentThreadId()}")


    def _init_ui(self): # (Giữ nguyên phần lớn, đảm bảo QLabel cho lbl_elapsed_time_display_widget)
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setMinimumWidth(300)
        left_panel.setMaximumWidth(380)

        self.txt_keyword = QTextEdit()
        self.txt_keyword.setPlaceholderText("Nhập từ khóa, mỗi từ một dòng...")
        left_layout.addWidget(self.txt_keyword)

        self.cb_instagram = QCheckBox("Instagram")
        self.cb_facebook_top = QCheckBox("Facebook Top")
        self.cb_facebook_videos = QCheckBox("Facebook Videos")
        self.cb_facebook_posts = QCheckBox("Facebook Posts")
        self.cb_nike = QCheckBox("Nike")
        self.cb_x_top = QCheckBox("X Top (Twitter)")
        self.cb_x_media = QCheckBox("X Media (Twitter)")
        self.cb_sneaker_news = QCheckBox("Sneaker News")
        self.cb_instagram.setChecked(True)

        search_options_layout = QVBoxLayout()
        checkboxes = [self.cb_instagram, self.cb_facebook_top, self.cb_facebook_videos,
                      self.cb_facebook_posts, self.cb_nike, self.cb_x_top,
                      self.cb_x_media, self.cb_sneaker_news]
        for cb in checkboxes: search_options_layout.addWidget(cb)
        left_layout.addLayout(search_options_layout)

        self.btn_search = QPushButton("Tìm kiếm (Ctrl+S)")
        left_layout.addWidget(self.btn_search)
        self.btn_stop_search = QPushButton("Dừng Tìm kiếm")
        self.btn_stop_search.setEnabled(False)
        left_layout.addWidget(self.btn_stop_search)

        self.btn_login = QPushButton("Đăng nhập vào Nền tảng")
        left_layout.addWidget(self.btn_login)
        self.btn_export_excel = QPushButton("Xuất ra Excel (Ctrl+E)")
        left_layout.addWidget(self.btn_export_excel)
        self.btn_update_keywords = QPushButton("Cập nhật File Từ khóa (Ctrl+U)")
        left_layout.addWidget(self.btn_update_keywords)
        self.btn_clear_cache = QPushButton("Xóa Cache Trình duyệt")
        left_layout.addWidget(self.btn_clear_cache)

        self.cb_auto_xlsx = QCheckBox("Tự động lưu Excel sau khi tìm")
        left_layout.addWidget(self.cb_auto_xlsx)
        self.cb_long_screenshot = QCheckBox("Thử Chụp Ảnh Dài (nếu tìm thấy)")
        self.cb_long_screenshot.setToolTip("Tính năng này có thể không ổn định.")
        left_layout.addWidget(self.cb_long_screenshot)

        left_layout.addStretch(1)
        main_layout.addWidget(left_panel)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        self.web_view = QWebEngineView()
        self.web_view.settings().setAttribute(QWebEngineSettings.JavascriptEnabled, True)
        self.web_view.settings().setAttribute(QWebEngineSettings.ScrollAnimatorEnabled, True)
        self.web_view.settings().setAttribute(QWebEngineSettings.FullScreenSupportEnabled, True)
        self.web_view.settings().setAttribute(QWebEngineSettings.AllowRunningInsecureContent, False)
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.63 Safari/537.36 Edg/102.0.1245.33" # User agent cập nhật
        self.web_view.page().profile().setHttpUserAgent(user_agent)
        self.web_view.load(QUrl("https://www.google.com.vn"))
        right_layout.addWidget(self.web_view, 3)

        self.dgv_results = QTableWidget()
        self.dgv_results.setColumnCount(5)
        self.dgv_results.setHorizontalHeaderLabels(["Website", "Keyword", "Link", "Found", "Data Summary"])
        header = self.dgv_results.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Interactive)
        self.dgv_results.setColumnWidth(4, 300)
        self.dgv_results.setEditTriggers(QTableWidget.NoEditTriggers)
        self.dgv_results.setSelectionBehavior(QTableWidget.SelectRows)
        self.dgv_results.setWordWrap(True)
        right_layout.addWidget(self.dgv_results, 1)

        main_layout.addWidget(right_panel, 1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.lbl_status_main_text = QLabel("Sẵn sàng")
        self.status_bar.addWidget(self.lbl_status_main_text, 1)

        self.lbl_elapsed_time_display_widget = QLabel("Elapsed Time: 00:00:00", self.status_bar) # Đã sửa thành QLabel
        self.status_bar.addPermanentWidget(self.lbl_elapsed_time_display_widget)

        self.progress_bar = QProgressBar(self.status_bar)
        self.status_bar.addPermanentWidget(self.progress_bar)
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumWidth(250)

    def _setup_connections(self): # (Giữ nguyên như trước)
        self.request_load_url_from_worker.connect(self.handle_load_url_request)
        self.request_execute_js_from_worker.connect(self.handle_execute_js_request)
        self.request_scroll_to_keyword_from_worker.connect(self.scroll_to_keyword_on_main)
        self.request_capture_screenshot_from_worker.connect(self.capture_screenshot_on_main)
        self.request_future_timeout_from_worker.connect(self.handle_future_timeout)


        self.btn_search.clicked.connect(self.start_search)
        self.btn_stop_search.clicked.connect(self.stop_current_search)
        self.btn_login.clicked.connect(self.show_login_menu)
        self.btn_export_excel.clicked.connect(self.export_to_excel)
        self.btn_update_keywords.clicked.connect(self.update_keywords_file)
        self.btn_clear_cache.clicked.connect(self.clear_browser_cache)

    def _setup_shortcuts(self): # (Giữ nguyên như trước)
        QShortcut(QKeySequence("Ctrl+S"), self, self.start_search)
        QShortcut(QKeySequence("Ctrl+E"), self, self.export_to_excel)
        QShortcut(QKeySequence("Ctrl+U"), self, self.update_keywords_file)
        QShortcut(QKeySequence("F5"), self, lambda: self.web_view.reload())

    # --- Các hàm xử lý yêu cầu từ Worker và WebEngine ---
    def handle_load_url_request(self, qurl, future_resolver_callback, callback_id):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: handle_load_url_request for {qurl.toString()}. Callback ID: {callback_id}")
        self.active_web_callbacks[callback_id] = future_resolver_callback
        
        page = self.web_view.page()
        # Ngắt kết nối cũ (nếu có) một cách an toàn hơn
        try: page.loadFinished.disconnect()
        except TypeError: pass
        
        # Kết nối slot mới, truyền callback_id để xác định đúng callback
        page.loadFinished.connect(lambda success, cid=callback_id: self._on_load_finished(success, cid))
        page.load(qurl)

    def _on_load_finished(self, success, callback_id):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: _on_load_finished(success={success}) for Callback ID: {callback_id}")
        
        # Ngắt kết nối ngay để tránh gọi lại cho các lần load không mong muốn
        try: self.web_view.page().loadFinished.disconnect(self._on_load_finished) # Cố gắng ngắt kết nối chính xác hàm này
        except TypeError: 
            try: self.web_view.page().loadFinished.disconnect() # Fallback ngắt tất cả
            except TypeError: pass


        if callback_id in self.active_web_callbacks:
            resolver_func = self.active_web_callbacks.pop(callback_id)
            if resolver_func:
                if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Resolving future for Callback ID: {callback_id} with success={success}")
                resolver_func(success) # Gọi hàm để giải quyết Future trong worker
            # else: # Điều này không nên xảy ra nếu quản lý đúng
            #     if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Warning - No resolver_func found for Callback ID: {callback_id} in _on_load_finished")
        # else:
        #     if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Warning - Callback ID: {callback_id} not in active_web_callbacks for _on_load_finished. Success: {success}. (Might be due to timeout or prior resolution)")
        pass


    def handle_execute_js_request(self, script, future_resolver_callback, callback_id):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: handle_execute_js_request. Script: {script[:60]}. Callback ID: {callback_id}")
        self.active_web_callbacks[callback_id] = future_resolver_callback
        
        # Callback nội bộ để xử lý kết quả JS và gọi lại resolver của worker
        def internal_js_result_handler(result, cid=callback_id):
            if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: internal_js_result_handler for Callback ID: {cid}. Result type: {type(result)}")
            if cid in self.active_web_callbacks:
                resolver_func = self.active_web_callbacks.pop(cid)
                if resolver_func:
                    if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Resolving JS future for Callback ID: {cid}")
                    resolver_func(result)
                # else:
                #     if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Warning - No resolver_func found for JS Callback ID: {cid}")
            # else:
            #     if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Warning - JS Callback ID: {cid} not in active_web_callbacks. (Might be due to timeout or prior resolution)")
            pass

        self.web_view.page().runJavaScript(script, internal_js_result_handler)

    def handle_future_timeout(self, callback_id):
        """Được gọi bởi worker khi một future bị timeout, để xóa callback tương ứng trên luồng chính."""
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: handle_future_timeout for Callback ID: {callback_id}")
        if callback_id in self.active_web_callbacks:
            self.active_web_callbacks.pop(callback_id)
            if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Removed timed-out callback for ID: {callback_id}")


    def update_status(self, message): # (Giữ nguyên)
        self.lbl_status_main_text.setText(message)
        # print(f"Status: {message}") # Bỏ comment nếu muốn log status ra console

    def update_progress_bar(self, value, status_text=""): # (Giữ nguyên)
        if not self.progress_bar.isVisible(): self.progress_bar.setVisible(True)
        self.progress_bar.setValue(value)
        if status_text: self.update_status(status_text)

    def update_elapsed_time_display_label(self): # (Giữ nguyên)
        if self.stopwatch_start_time is not None:
            elapsed_seconds = int(time.time() - self.stopwatch_start_time)
            h, m, s = elapsed_seconds // 3600, (elapsed_seconds % 3600) // 60, elapsed_seconds % 60
            self.lbl_elapsed_time_display_widget.setText(f"Elapsed Time: {h:02}:{m:02}:{s:02}")
        else: self.lbl_elapsed_time_display_widget.setText("Elapsed Time: 00:00:00")

    def load_keywords_from_file(self): # (Giữ nguyên)
        try:
            file_path = "keywords.txt"
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8") as f: self.txt_keyword.setPlainText(f.read())
                self.update_status("Đã tải từ khóa từ keywords.txt")
            else: self.update_status("Không tìm thấy keywords.txt. Nhập từ khóa thủ công.")
        except Exception as e: QMessageBox.warning(self, "Lỗi File", f"Lỗi khi tải từ khóa: {e}")

    def update_keywords_file(self): # (Giữ nguyên)
        try:
            file_path = "keywords.txt"
            with open(file_path, "w", encoding="utf-8") as f: f.write(self.txt_keyword.toPlainText())
            self.update_status("Đã cập nhật và lưu từ khóa vào keywords.txt.")
            QMessageBox.information(self, "Đã lưu Từ khóa", "Đã cập nhật và lưu từ khóa.")
        except Exception as e: QMessageBox.critical(self, "Lỗi File", f"Lỗi khi lưu từ khóa: {e}")

    def build_search_tasks_list(self, keywords_list_str): # (Giữ nguyên)
        tasks = []
        for keyword_str in keywords_list_str:
            if not keyword_str.strip(): continue
            kw = keyword_str.strip()
            if self.cb_instagram.isChecked(): tasks.append({"platform_name": "Instagram", "keyword": kw})
            if self.cb_facebook_top.isChecked(): tasks.append({"platform_name": "Facebook Top", "keyword": kw, "url_generator": SearchEngines.get_facebook_top_search_url})
            # ... các checkbox khác ...
            if self.cb_facebook_videos.isChecked(): tasks.append({"platform_name": "Facebook Videos", "keyword": kw, "url_generator": SearchEngines.get_facebook_videos_search_url})
            if self.cb_facebook_posts.isChecked(): tasks.append({"platform_name": "Facebook Posts", "keyword": kw, "url_generator": SearchEngines.get_facebook_posts_search_url})
            if self.cb_nike.isChecked(): tasks.append({"platform_name": "Nike", "keyword": kw, "url_generator": SearchEngines.get_nike_search_url})
            if self.cb_x_top.isChecked(): tasks.append({"platform_name": "X Top", "keyword": kw, "url_generator": SearchEngines.get_x_top_search_url})
            if self.cb_x_media.isChecked(): tasks.append({"platform_name": "X Media", "keyword": kw, "url_generator": SearchEngines.get_x_media_search_url})
            if self.cb_sneaker_news.isChecked(): tasks.append({"platform_name": "Sneaker News", "keyword": kw, "url_generator": SearchEngines.get_sneaker_news_search_url})
        return tasks

    def start_search(self): # (Cải thiện việc quản lý thread và worker)
        if self.search_thread and self.search_thread.isRunning():
            QMessageBox.information(self, "Đang Tìm kiếm", "Một tìm kiếm đang chạy. Đợi hoặc dừng nó.")
            return

        keywords = [k.strip() for k in self.txt_keyword.toPlainText().splitlines() if k.strip()]
        if not keywords:
            QMessageBox.information(self, "Yêu cầu Nhập liệu", "Nhập ít nhất một từ khóa.")
            return

        self.current_search_tasks_list = self.build_search_tasks_list(keywords)
        if not self.current_search_tasks_list:
            QMessageBox.information(self, "Yêu cầu Chọn lựa", "Chọn ít nhất một tùy chọn tìm kiếm.")
            return

        self.btn_search.setEnabled(False)
        self.btn_stop_search.setEnabled(True)
        self.dgv_results.setRowCount(0)
        self.results_list.clear()
        self.active_web_callbacks.clear()

        self.stopwatch_start_time = time.time()
        self.elapsed_timer_display.start()
        self.update_status("Bắt đầu tìm kiếm...")
        self.update_progress_bar(0)

        self.search_worker_obj = SearchWorker(self.current_search_tasks_list, self)
        self.search_thread = QThread(self) # Đặt self làm parent cho QThread
        self.search_worker_obj.moveToThread(self.search_thread)

        # Kết nối tín hiệu từ worker
        self.search_worker_obj.finished_task.connect(self.add_result_to_grid_display)
        self.search_worker_obj.progress_update.connect(self.update_progress_bar)
        self.search_worker_obj.search_completed_all.connect(self.handle_search_completion_event)
        
        # Quản lý vòng đời của thread và worker
        self.search_thread.started.connect(self.search_worker_obj.run)
        self.search_thread.finished.connect(self._on_search_thread_finished) # Slot để dọn dẹp worker
        
        self.search_thread.start()
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Search thread started. Worker: {self.search_worker_obj}, Thread: {self.search_thread}")


    def _on_search_thread_finished(self):
        """Được gọi khi QThread.finished() được phát ra."""
        if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN: _on_search_thread_finished called.")
        if self.search_worker_obj:
            self.search_worker_obj.deleteLater() # Yêu cầu Qt xóa worker một cách an toàn
            self.search_worker_obj = None
        if self.search_thread: # Thread vẫn còn tồn tại ở đây, deleteLater của nó cũng nên được gọi
             self.search_thread.deleteLater() # Yêu cầu Qt xóa thread một cách an toàn
             self.search_thread = None # Chỉ đặt lại nếu cleanup_after_search_processing không làm
        # Các nút bấm và UI khác sẽ được xử lý trong handle_search_completion_event -> cleanup_after_search_processing

    def stop_current_search(self):
        if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN: stop_current_search called.")
        self.update_status("Đang yêu cầu dừng tìm kiếm...")
        self.btn_stop_search.setEnabled(False)
        if self.search_worker_obj:
            self.search_worker_obj.stop() # Gửi tín hiệu dừng cho worker
        # Worker sẽ tự kết thúc, và QThread.finished sẽ được phát ra, kích hoạt _on_search_thread_finished
        # handle_search_completion_event cũng sẽ được gọi bởi worker với thông báo "đã dừng"

    def handle_search_completion_event(self, final_message):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: handle_search_completion_event: {final_message}")
        found_count = sum(1 for r in self.results_list if r.found == "Có")
        completion_msg = f"{final_message} Tìm thấy {found_count} kết quả khớp."

        if self.cb_auto_xlsx.isChecked() and self.results_list:
            self.update_status("Đang tự động lưu kết quả ra Excel...")
            self.auto_save_excel_file()

        self.cleanup_after_search_processing(completion_msg, True)

    def cleanup_after_search_processing(self, final_status_message, show_completion_messagebox):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: cleanup_after_search_processing: {final_status_message}")
        self.elapsed_timer_display.stop()
        self.update_progress_bar(100)
        QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False)) # Ẩn sau 1s

        self.btn_search.setEnabled(True)
        self.btn_stop_search.setEnabled(False)

        # Luồng và worker nên đã được xử lý bởi _on_search_thread_finished
        # hoặc sẽ được xử lý khi tín hiệu finished của thread phát ra.
        # Chỉ cần đảm bảo các tham chiếu được đặt lại nếu chưa.
        if self.search_thread and not self.search_thread.isRunning():
            if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN (Cleanup): Search thread is not running, ensuring it's None.")
            self.search_thread = None # Nó đã được deleteLater hoặc sẽ sớm thôi
            self.search_worker_obj = None # Worker cũng vậy

        self.update_status(final_status_message)
        if show_completion_messagebox:
            QMessageBox.information(self, "Trạng thái Tìm kiếm", final_status_message)
        
        self.active_web_callbacks.clear()

    def add_result_to_grid_display(self, website, keyword, link, found_status, full_data): # (Giữ nguyên)
        summary = (full_data[:200] + "...") if len(full_data) > 200 else full_data
        result = SearchResult(website, keyword, link, found_status, full_data, summary)
        self.results_list.append(result)
        # ... (thêm vào bảng) ...
        row_position = self.dgv_results.rowCount()
        self.dgv_results.insertRow(row_position)
        self.dgv_results.setItem(row_position, 0, QTableWidgetItem(website))
        self.dgv_results.setItem(row_position, 1, QTableWidgetItem(keyword))
        item_link = QTableWidgetItem(link)
        self.dgv_results.setItem(row_position, 2, item_link)
        self.dgv_results.setItem(row_position, 3, QTableWidgetItem(found_status))
        self.dgv_results.setItem(row_position, 4, QTableWidgetItem(summary))
        self.dgv_results.resizeRowsToContents()
        self.update_status(f"Đã thêm: {website} - '{keyword}' ({found_status})")


    def auto_save_excel_file(self): # (Giữ nguyên, nhưng kiểm tra openpyxl)
        if not self.results_list:
            self.update_status("Không có kết quả để lưu ra Excel.")
            return
        if not openpyxl:
            QMessageBox.warning(self, "Lỗi Excel", "Thư viện openpyxl chưa được cài đặt.")
            return
        # ... (logic lưu excel) ...
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            folder_name = "IM_Keyword_Reports_Py"
            folder_path = os.path.join(desktop_path, folder_name)
            os.makedirs(folder_path, exist_ok=True)

            base_filename = "KeywordSearchResults"
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = os.path.join(folder_path, f"{base_filename}_{timestamp}.xlsx")

            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "SearchResults"

            headers = ["Website", "Keyword", "Link", "Found", "Full Data Summary"]
            sheet.append(headers)
            for cell in sheet[1]: cell.font = Font(bold=True)

            for r_idx, r_data in enumerate(self.results_list, start=2):
                sheet.cell(row=r_idx, column=1, value=r_data.website)
                sheet.cell(row=r_idx, column=2, value=r_data.keyword)
                link_cell_val = r_data.link
                link_cell = sheet.cell(row=r_idx, column=3, value=link_cell_val)
                if link_cell_val and (link_cell_val.startswith("http://") or link_cell_val.startswith("https://")):
                    try:
                        link_cell.hyperlink = link_cell_val
                        link_cell.style = "Hyperlink"
                    except Exception: pass # Bỏ qua nếu không đặt được hyperlink
                sheet.cell(row=r_idx, column=4, value=r_data.found)
                sheet.cell(row=r_idx, column=5, value=r_data.full_data_summary)
            
            for col_idx, column_letter in enumerate(['A', 'B', 'C', 'D', 'E'], start=1):
                # ... (logic điều chỉnh chiều rộng cột) ...
                pass # Tạm thời bỏ qua để tránh lỗi không mong muốn với openpyxl versioning

            workbook.save(file_path)
            self.update_status(f"File Excel đã lưu: {os.path.basename(file_path)}")
            QMessageBox.information(self, "Xuất Excel", f"Kết quả đã được lưu vào:\n{file_path}")
        except Exception as e:
            self.update_status(f"Lỗi khi lưu file Excel: {e}")
            QMessageBox.critical(self, "Lỗi Xuất Excel", f"Không thể lưu file Excel: {e}\n{type(e)}")


    def export_to_excel(self): # (Giữ nguyên)
        if not self.results_list:
            QMessageBox.information(self, "Xuất ra Excel", "Không có kết quả để xuất.")
            return
        self.auto_save_excel_file()

    def show_login_menu(self): # (Giữ nguyên)
        menu = QMenu(self)
        actions_data = [
            ("Đăng nhập Instagram", "https://www.instagram.com/accounts/login/"),
            ("Đăng nhập Facebook", "https://www.facebook.com/login/"),
            ("Đăng nhập X (Twitter)", "https://x.com/login")
        ]
        for text, url_str in actions_data:
            action = QAction(text, self)
            action.triggered.connect(lambda checked, u=url_str: self.web_view.load(QUrl(u)))
            menu.addAction(action)
        menu.popup(self.btn_login.mapToGlobal(self.btn_login.rect().bottomLeft()))

    def clear_browser_cache(self): # (Giữ nguyên)
        if self.web_view:
            try:
                self.update_status("Đang xóa cache và cookies WebView...")
                profile = self.web_view.page().profile()
                profile.clearHttpCache()
                profile.cookieStore().deleteAllCookies()
                QMessageBox.information(self, "Đã xóa Cache", "Cache HTTP và Cookies của WebView đã được xóa.")
                self.update_status("Cache và Cookies WebView đã được xóa.")
            except Exception as e: QMessageBox.critical(self, "Lỗi Cache", f"Lỗi khi xóa cache: {e}")
        else: QMessageBox.warning(self, "Lỗi", "WebView chưa được khởi tạo.")

    def scroll_to_keyword_on_main(self, keyword): # (Giữ nguyên hoặc cải thiện JS)
        self.update_status(f"Đang cuộn đến từ khóa '{keyword}'...")
        try:
            sanitized_keyword = json.dumps(keyword.lower())[1:-1]
            js_scroll = f""" /* ... JS scroll ... */ """ # Giữ nguyên JS scroll
            self.web_view.page().runJavaScript(js_scroll, lambda res: print(f"Scroll result: {res}") if DEBUG_MAIN_THREAD_HANDLERS else None)
        except Exception as e: self.update_status(f"Lỗi khi cuộn đến từ khóa: {e}")

    def capture_screenshot_on_main(self, website, keyword, identifier): # (Cải thiện tên file và logic long screenshot)
        self.update_status(f"Đang chụp ảnh màn hình cho {website} - '{keyword}'...")
        try:
            # ... (Logic tạo đường dẫn và tên file như trước) ...
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            folder_name = "IM_Keyword_Screenshots_Py"
            folder_path = os.path.join(desktop_path, folder_name)
            os.makedirs(folder_path, exist_ok=True)

            safe_website = "".join(c if c.isalnum() else "_" for c in website)
            safe_keyword = "".join(c if c.isalnum() else "_" for c in keyword)
            
            clean_identifier = "".join(c if c.isalnum() else "_" for c in identifier.replace("https://","").replace("http://","").replace("/","_"))[:50]

            file_name = f"{safe_website}_{safe_keyword}_{clean_identifier}_{datetime.datetime.now():%H%M%S}.png"
            file_path = os.path.join(folder_path, file_name)


            if self.cb_long_screenshot.isChecked():
                if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Requesting document height for long screenshot: {file_path}")
                self.web_view.page().runJavaScript("Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );",
                                                lambda h_str: self._capture_long_js_callback(h_str, file_path, website, keyword))
            else:
                self._save_normal_screenshot(file_path, website, keyword)
        except Exception as e:
            self.update_status(f"Lỗi khi chụp ảnh màn hình: {e}")


    def _capture_long_js_callback(self, height_str_js, file_path, website, keyword):
        original_size = self.web_view.size()
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: _capture_long_js_callback. Height str: '{height_str_js}'. File: {file_path}")
        try:
            doc_height = 0
            if isinstance(height_str_js, (int, float)):
                doc_height = int(height_str_js)
            elif isinstance(height_str_js, str) and height_str_js.isdigit():
                doc_height = int(height_str_js)
            
            if doc_height > 0:
                view_width = original_size.width()
                max_screenshot_height = 16384 
                capture_height = min(doc_height, max_screenshot_height)

                if capture_height > original_size.height():
                    if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Resizing webview to {view_width}x{capture_height} for long screenshot.")
                    self.web_view.resize(view_width, capture_height)
                    QTimer.singleShot(1200, lambda: self._save_resized_and_restore(file_path, original_size, website, keyword))
                    return
                else:
                    if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Document height ({doc_height}) not larger than view height ({original_size.height()}). Capturing normal.")
                    self._save_normal_screenshot(file_path, website, keyword)
            else:
                if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Invalid document height ({height_str_js}). Capturing normal.")
                self._save_normal_screenshot(file_path, website, keyword)
        except Exception as e:
            self.update_status(f"Lỗi trong callback chụp ảnh dài: {e}")
            self._save_normal_screenshot(file_path, website, keyword)
            self.web_view.resize(original_size)

    def _save_resized_and_restore(self, file_path, original_size, website, keyword):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: _save_resized_and_restore for {file_path}")
        try:
            # QPixmap cần được tạo với kích thước chính xác của widget tại thời điểm render
            current_webview_size = self.web_view.size()
            pixmap = QPixmap(current_webview_size)
            if pixmap.isNull(): # Kiểm tra nếu pixmap không hợp lệ (ví dụ kích thước quá lớn)
                 if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Failed to create QPixmap of size {current_webview_size}. Fallback to grab().")
                 pixmap = self.web_view.grab() # Fallback về grab() nếu render lỗi
            else:
                self.web_view.render(pixmap)
            
            if pixmap.save(file_path, "PNG"):
                self.update_status(f"Ảnh màn hình dài đã lưu: {os.path.basename(file_path)}")
            else:
                self.update_status(f"Lỗi: Không thể lưu ảnh màn hình dài: {os.path.basename(file_path)}")
        except Exception as e:
             self.update_status(f"Lỗi khi lưu ảnh dài đã resize: {e}")
        finally:
            if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: Restoring webview to original size: {original_size}")
            self.web_view.resize(original_size)

    def _save_normal_screenshot(self, file_path, website, keyword):
        if DEBUG_MAIN_THREAD_HANDLERS: print(f"MAIN: _save_normal_screenshot for {file_path}")
        try:
            pixmap = self.web_view.grab()
            if pixmap.save(file_path, "PNG"):
                self.update_status(f"Ảnh màn hình (thường) đã lưu: {os.path.basename(file_path)}")
            else:
                self.update_status(f"Lỗi: Không thể lưu ảnh (thường): {os.path.basename(file_path)}")
        except Exception as e:
            self.update_status(f"Lỗi khi lưu ảnh thường: {e}")

    def closeEvent(self, event):
        if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN: closeEvent triggered.")
        if self.search_thread and self.search_thread.isRunning():
            reply = QMessageBox.question(self, 'Đang thoát...',
                                         "Một tìm kiếm đang chạy. Bạn có muốn dừng và thoát không?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.update_status("Đang dừng các tác vụ và thoát...")
                if self.search_worker_obj:
                    if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN (CloseEvent): Calling worker.stop()")
                    self.search_worker_obj.stop() 
                
                # Không gọi QThread.quit() hoặc wait() trực tiếp ở đây nếu worker
                # có thể mất thời gian để dừng asyncio loop.
                # Worker sẽ tự kết thúc và QThread.finished sẽ được phát ra.
                # Tuy nhiên, vì ứng dụng đang đóng, chúng ta cần đảm bảo nó không bị treo.
                # Một cách là cho một timeout nhỏ rồi accept event.
                # QTimer.singleShot(1000, event.accept) # Chấp nhận đóng sau 1s
                # Hoặc, nếu worker.stop() nhanh chóng làm worker.run() kết thúc:
                if self.search_thread.isRunning():
                     if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN (CloseEvent): Requesting thread quit and waiting briefly.")
                     self.search_thread.quit() # Yêu cầu thoát event loop của thread
                     if not self.search_thread.wait(1500): # Chờ tối đa 1.5s
                         if DEBUG_MAIN_THREAD_HANDLERS: print("MAIN (CloseEvent): Thread did not finish in time. App will close.")
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Tùy chọn: Bật remote debugging cho QWebEngineView
    # os.environ["QTWEBENGINE_REMOTE_DEBUGGING"] = "9223" 
    # Sau đó truy cập localhost:9223 trong trình duyệt Chrome/Edge
    
    main_win = KeywordSearchApp()
    main_win.show()
    sys.exit(app.exec_())
