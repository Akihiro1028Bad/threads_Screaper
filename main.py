import sys
import time
import logging
import json
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_driver():
    chrome_options = Options()
    #chrome_options.add_argument("--headless")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    logging.info("Chromeドライバーを設定しました。ヘッドレスモード: 有効")
    return driver

def wait_and_log(seconds, message):
    logging.info(f"{message}、{seconds}秒間待機します...")
    time.sleep(seconds)

def save_cookies(driver, filename='cookies.json'):
    """ドライバーのクッキーを保存する"""
    with open(filename, 'w') as f:
        json.dump(driver.get_cookies(), f)
    logging.info(f"クッキーを {filename} に保存しました")

def load_cookies(driver, filename='cookies.json'):
    """保存されたクッキーをドライバーにロードする"""
    with open(filename, 'r') as f:
        cookies = json.load(f)
    for cookie in cookies:
        driver.add_cookie(cookie)
    logging.info(f"{filename} からクッキーをロードしました")

def login_to_threads(driver, username, password):
    url = "https://www.threads.net/login/"

    try:
        driver.get(url)
        logging.info(f"ログインページにアクセスしています: {url}")
        
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text'][class*='x1i10hfl'][class*='x1a2a7pz']"))
        )
        username_field.clear()
        username_field.send_keys(username)
        logging.info("ユーザー名を入力しました")
        
        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        password_field.clear()
        password_field.send_keys(password)
        logging.info("パスワードを入力しました")
        
        wait_and_log(5, "ログインボタンをクリックする前に待機")
        
        login_button_xpath = "//div[@role='button' and contains(@class, 'x1i10hfl') and contains(@class, 'x1qjc9v5')]//div[contains(text(), 'Log in') or contains(text(), 'ログイン')]"
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, login_button_xpath))
        )
        driver.execute_script("arguments[0].click();", login_button)
        logging.info("ログインボタンをクリックしました")
        
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        logging.info("ログインが完了し、ページが読み込まれました")

        time.sleep(10)
        
        # ログイン成功後、クッキーを保存
        save_cookies(driver)
        
        return True

    except Exception as e:
        logging.error(f"ログイン中にエラーが発生しました: {e}")


def check_login_status(driver):
    """'Post'または'投稿'要素の存在に基づいてログイン状態を確認する"""
    try:
        # 'Post'または'投稿'要素を探す
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, 
                "//div[contains(@class, 'xc26acl') and contains(@class, 'x6s0dn4') and contains(@class, 'x78zum5') and (contains(text(), 'Post') or contains(text(), '投稿'))]"
            ))
        )
        logging.info("'Post'または'投稿'要素が見つかりました。ログイン状態が確認されました。")
        return True
    except:
        logging.info("'Post'または'投稿'要素が見つかりません。ログアウト状態です。")
        return False

def refresh_session(driver, username, password):
    """セッションを更新する"""
    if not check_login_status(driver):
        logging.info("セッションが切れています。再ログインを試みます。")
        return login_to_threads(driver, username, password)
    return True

def scroll_page(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    logging.info("ページを下部までスクロールしました")
    time.sleep(2)

def get_post_hrefs(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    post_hrefs = []
    elements = soup.find_all('a', class_=['x1i10hfl', 'x1lliihq'], href=True)
    for element in elements:
        href = element['href']
        if '/post/' in href and href not in post_hrefs:
            post_hrefs.append(href)
    return post_hrefs

def access_threads_profile(driver, username, login_username, login_password):
    url = f"https://www.threads.net/@{username}"
    all_post_hrefs = []
    
    try:
        # セッションを更新
        if not refresh_session(driver, login_username, login_password):
            raise Exception("セッションの更新に失敗しました")

        driver.get(url)
        logging.info(f"プロフィールページにアクセスしています: {url}")
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        logging.info(f"@{username} のプロフィールページの読み込みが完了しました")
        
        no_new_links_count = 0
        
        while True:
            previous_count = len(all_post_hrefs)
            new_hrefs = get_post_hrefs(driver.page_source)
            all_post_hrefs.extend([href for href in new_hrefs if href not in all_post_hrefs])
            
            if len(all_post_hrefs) > previous_count:
                logging.info(f"新たに {len(all_post_hrefs) - previous_count} 件の投稿リンクを見つけました。合計: {len(all_post_hrefs)} 件")
                no_new_links_count = 0
            else:
                no_new_links_count += 1
                logging.info(f"新しいリンクが見つかりません。スクロール回数: {no_new_links_count}")
            
            if no_new_links_count >= 3:
                logging.info("3回連続で新しいリンクが見つからなかったため、収集を終了します")
                break
            
            scroll_page(driver)
            time.sleep(2)
        
        logging.info("データ収集が完了しました")
        logging.info(f"収集された一意の投稿リンクの総数: {len(all_post_hrefs)} 件")
        
    except Exception as e:
        logging.error(f"プロフィールページのアクセス中にエラーが発生しました: {e}")
        return []
    
    return all_post_hrefs

def extract_post_datetime(driver):
    try:
        # 親要素を特定
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span.x1lliihq.x1plvlek.xryxfnj.x1n2onr6.x193iq5w.xeuugli.x1fj9vlw.x13faqbe.x1vvkbs.x1s928wv.xhkezso.x1gmr53x.x1cpjm7i.x1fgarty.x1943h6x.x1i0vuye.xjohtrz.xo1l8bm.x12rw4y6.x1yc453h"))
        )
        
        # time要素を特定
        time_element = parent_element.find_element(By.TAG_NAME, "time")
        
        # datetime属性から投稿日時を取得
        datetime_str = time_element.get_attribute('datetime')
        
        # ISO形式の文字列をdatetimeオブジェクトに変換
        post_datetime = datetime.fromisoformat(datetime_str.replace('Z', '+00:00'))
        
        # 日本時間に変換（UTC+9）
        japan_time = post_datetime + timedelta(hours=9)
        
        # 指定のフォーマットに変換
        formatted_datetime = japan_time.strftime("%m月%d日%H時%M分")
        
        return formatted_datetime
    except Exception as e:
        logging.error(f"投稿日時の抽出中にエラーが発生しました: {e}")
        return None


def extract_like_count(driver):
    try:
        # 親要素を特定
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.x6s0dn4.xfkn95n.xly138o.xchwasx.xfxlei4.x17zd0t2.x78zum5.xl56j7k.x1n2onr6.x3oybdh.x12w9bfk.xx6bhzk.x11xpdln.xc9qbxq.x1ye3gou.xn6708d.x14atkfc"))
        )
        
        # 子要素（いいね数を含む要素）を特定
        like_element = parent_element.find_element(By.CSS_SELECTOR, "span.x1lliihq.x1plvlek.xryxfnj.x1n2onr6.x193iq5w.xeuugli.x1fj9vlw.x13faqbe.x1vvkbs.x1s928wv.xhkezso.x1gmr53x.x1cpjm7i.x1fgarty.x1943h6x.x1i0vuye.xmd891q.xo1l8bm.xc82ewx.x1yc453h div.xu9jpxn.x1n2onr6.xqcsobp.x12w9bfk.x1wsgiic.xuxw1ft.x17fnjtu span.x17qophe.x10l6tqk.x13vifvy")
        
        # いいね数を取得
        like_count = like_element.text
        
        return int(like_count) if like_count.isdigit() else 0
    except Exception as e:
        logging.error(f"いいね数の抽出中にエラーが発生しました: {e}")
        return 0

def extract_comment_count(driver):
    try:
        # 親要素を特定
        parent_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.x6s0dn4.xfkn95n.xly138o.xchwasx.xfxlei4.x17zd0t2.x78zum5.xl56j7k.x1n2onr6.x3oybdh.x12w9bfk.xx6bhzk.x11xpdln.xc9qbxq.x1ye3gou.xn6708d.x14atkfc"))
        )
        
        if len(parent_elements) < 2:
            logging.warning("コメント数の要素が見つかりません。0を返します。")
            return 0
        
        # 2番目の親要素を選択
        second_parent = parent_elements[1]
        
        # 子要素（コメント数を含む要素）を特定
        comment_element = second_parent.find_element(By.CSS_SELECTOR, "span.x1lliihq.x1plvlek.xryxfnj.x1n2onr6.x193iq5w.xeuugli.x1fj9vlw.x13faqbe.x1vvkbs.x1s928wv.xhkezso.x1gmr53x.x1cpjm7i.x1fgarty.x1943h6x.x1i0vuye.xmd891q.xo1l8bm.xc82ewx.x1yc453h div.xu9jpxn.x1n2onr6.xqcsobp.x12w9bfk.x1wsgiic.xuxw1ft.x17fnjtu span.x17qophe.x10l6tqk.x13vifvy")
        
        # コメント数を取得
        comment_count = comment_element.text
        
        return int(comment_count) if comment_count.isdigit() else 0
    except Exception as e:
        logging.error(f"コメント数の抽出中にエラーが発生しました: {e}")
        return 0

def extract_image_urls(driver):
    try:
        # 画像を含む親要素を待機して取得
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.x1xmf6yo"))
        )

        # 画像要素を探す
        image_elements = parent_element.find_elements(By.CSS_SELECTOR, "img.xl1xv1r")
        
        # 画像URLを取得
        image_urls = [img.get_attribute('src') for img in image_elements if img.get_attribute('src')]
        
        # 画像がない場合は「なし」を返す
        return image_urls if image_urls else ["なし"]
    except Exception as e:
        logging.error(f"画像URLの抽出中にエラーが発生しました: {e}")
        return ["なし"]

def extract_caption(driver):
    try:
        # キャプションを含む親要素を待機して取得
        parent_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.x1a6qonq.x6ikm8r.x10wlt62.xj0a0fe.x126k92a.x6prxxf.x7r5mf7"))
        )
        
        # h1とh2要素を取得
        h_elements = parent_element.find_elements(By.CSS_SELECTOR, "h1, h2")
        
        # テキスト部分を抽出して結合
        caption_text = ' '.join([elem.text.strip() for elem in h_elements])
        
        return caption_text if caption_text else None
    
    except Exception as e:
        logging.error(f"キャプションの抽出中にエラーが発生しました: {e}")
        return None

def extract_reply_count(driver, max_scroll_attempts=5, scroll_pause_time=2):
    try:
        reply_count = 0
        scroll_attempts = 0
        last_height = driver.execute_script("return document.body.scrollHeight")
        
        while scroll_attempts < max_scroll_attempts:
            try:
                # x78zum5 xdt5ytf クラスを持つdiv要素を探す
                outer_elements = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='x78zum5 xdt5ytf']"))
                )
                
                for outer_element in outer_elements:
                    # 各外部要素内で最初の x1a2a7pz x1n2onr6 クラスを持つdiv要素を探す
                    inner_element = outer_element.find_element(By.CSS_SELECTOR, "div[class='x1a2a7pz x1n2onr6']")
                    
                    # 内部要素内で特定のspan要素を探す
                    try:
                        span_element = inner_element.find_element(By.CSS_SELECTOR, "span[class='x17qophe x10l6tqk x13vifvy']")
                        if span_element.text == "1":
                            reply_count += 1
                    except NoSuchElementException:
                        # span要素が見つからない場合は何もしない
                        pass
                
                # ページ全体をスクロール
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                # 新しいコンテンツがロードされるのを待つ
                time.sleep(scroll_pause_time)
                
                # 新しい高さを取得
                new_height = driver.execute_script("return document.body.scrollHeight")
                
                # スクロールしても高さが変わらなければループを抜ける
                if new_height == last_height:
                    break
                
                last_height = new_height
                scroll_attempts += 1
            
            except (StaleElementReferenceException, NoSuchElementException):
                # 要素が見つからない場合は少し待ってリトライ
                time.sleep(1)
                continue
        
        return reply_count
    except TimeoutException:
        logging.warning("要素の読み込みがタイムアウトしました。返信数は0として扱います。")
        return 0
    except Exception as e:
        logging.error(f"コメント返信数の抽出中に予期せぬエラーが発生しました: {e}")
        return 0

def process_posts(driver, post_hrefs):
    processed_posts = []
    previous_post = None
    
    for href in post_hrefs:
        try:
            full_url = f"https://www.threads.net{href}"
            driver.get(full_url)
            logging.info(f"アクセス中: {full_url}")
            
            # ページの読み込みを待機
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # 投稿日時を抽出
            post_datetime = extract_post_datetime(driver)

            # いいね数を抽出
            like_count = extract_like_count(driver)

            # コメント数を抽出
            comment_count = extract_comment_count(driver)

            # コメント返信数を抽出
            reply_count = extract_reply_count(driver)

            # 画像URLを抽出
            image_urls = extract_image_urls(driver)

            # キャプションを抽出
            caption = extract_caption(driver)
            
            current_post = {
                'url': full_url,
                'datetime': post_datetime,
                'like_count': like_count,
                'comment_count': comment_count,
                'reply_count': reply_count,
                'caption': caption,
                'image_urls': image_urls,
            }
            
            # 前の投稿と比較
            if previous_post and is_duplicate_post(previous_post, current_post):
                logging.info(f"重複投稿をスキップ: {full_url}")
                continue
            
            processed_posts.append(current_post)
            previous_post = current_post
            
            logging.info(f"処理完了: {full_url}")
            logging.info(f"  - 日時: {post_datetime}")
            logging.info(f"  - いいね数: {like_count}")
            logging.info(f"  - コメント数: {comment_count}")
            logging.info(f"  - コメント返信数: {reply_count}") 
            logging.info(f"  - キャプション: {caption}")
            logging.info(f"  - 画像URL: {', '.join(image_urls)}")
            
            # レート制限を回避するための待機
            time.sleep(2)
        
        except Exception as e:
            logging.error(f"投稿の処理中にエラーが発生しました: {full_url}, エラー: {e}")

    return processed_posts

def is_duplicate_post(previous_post, current_post):
    """
    前の投稿と現在の投稿が重複しているかどうかを判断する関数
    """
    return (
        previous_post['datetime'] == current_post['datetime'] and
        previous_post['caption'] == current_post['caption']
    )



def validate_input(input_string, input_type):
    if not input_string:
        raise ValueError(f"{input_type}が空です。有効な値を入力してください。")
    return input_string

if __name__ == "__main__":
    try:
        login_username = validate_input(input("Threadsのログインユーザー名を入力してください: ").strip(), "ユーザー名")
        login_password = validate_input(input("Threadsのパスワードを入力してください: ").strip(), "パスワード")
        target_username = validate_input(input("スクレイピングしたいユーザー名を入力してください: ").strip(), "ターゲットユーザー名")

        driver = setup_driver()
        
        if login_to_threads(driver, login_username, login_password):
            post_hrefs = access_threads_profile(driver, target_username, login_username, login_password)
            
            logging.info(f"収集された一意の投稿リンクの総数: {len(post_hrefs)} 件")
            logging.info("収集されたすべてのリンク:")
            for href in post_hrefs:
                print(href)

             # 投稿の処理
            processed_posts = process_posts(driver, post_hrefs)
            
            logging.info(f"処理された投稿の総数: {len(processed_posts)} 件")
            
            # ここに処理済みの投稿データを使用するコードを追加します
            # 例: save_to_excel(processed_posts)
        else:
            logging.error("ログインに失敗したため、スクレイピングを実行できません。")

    except ValueError as ve:
        logging.error(f"入力エラー: {ve}")
    except Exception as e:
        logging.error(f"予期せぬエラーが発生しました: {e}")
    finally:
        if 'driver' in locals() and driver:
            driver.quit()
        logging.info("すべてのリソースを解放しました。プログラムを終了します。")