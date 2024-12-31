from selenium import webdriver  # 导入浏览器驱动
from selenium.webdriver.common.by import By  # 导入定位器
from selenium.webdriver.common.keys import Keys  # 导入键盘操作
from selenium.webdriver.support.ui import WebDriverWait  # 导入显式等待
from selenium.webdriver.support import expected_conditions as EC  # 导入预期条件
from selenium.webdriver.chrome.service import Service  # 导入Service类
from selenium.webdriver.chrome.options import Options  # 导入Chrome选项
from webdriver_manager.chrome import ChromeDriverManager  # 导入ChromeDriver管理器
import pandas as pd  # 导入pandas用于数据处理
import time  # 导入time模块用于延时
import os  # 导入os模块用于文件操作
from datetime import datetime  # 导入datetime用于处理日期时间

def setup_driver():
    """设置并返回Chrome浏览器驱动"""
    chrome_options = Options()
    chrome_options.add_argument('--start-maximized')  # 最大化窗口
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')  # 禁用自动化标志
    
    # 添加更多反爬虫设置
    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # 添加随机UA
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
    chrome_options.add_argument(f'user-agent={user_agent}')
    
    # 使用ChromeDriverManager自动下载和管理ChromeDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # 执行CDP命令来修改navigator.webdriver标志
    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        'source': '''
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        '''
    })
    
    return driver

def create_content_folder():
    """创建保存内容的文件夹"""
    content_folder = os.path.join(os.path.dirname(__file__), "内容")
    if not os.path.exists(content_folder):
        os.makedirs(content_folder)
    return content_folder

def get_current_time():
    """获取当前时间字符串"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def scrape_articles():
    """爬取微信文章"""
    driver = setup_driver()
    content_folder = create_content_folder()
    articles_data = []  # 存储所有文章数据
    
    try:
        # 打开搜狗微信搜索页面
        driver.get("https://weixin.sogou.com/")
        
        # 等待搜索框出现并输入关键词
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "query"))
        )
        search_box.clear()
        search_box.send_keys("AI")
        
        # 修改这部分：等待搜索按钮出现并点击
        try:
            # 首先尝试使用新的选择器
            search_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn-search"))
            )
        except:
            try:
                # 如果失败，尝试使用其他可能的选择器
                search_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "[uigs='search_article']"))
                )
            except:
                # 最后尝试直接按回车键搜索
                search_box.send_keys(Keys.RETURN)
                time.sleep(2)
        else:
            search_button.click()
            time.sleep(2)
        
        # 确保搜索结果已加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "news-box"))
        )
        
        # 添加检查是否需要验证码的逻辑
        if "验证码" in driver.page_source:
            input("请在浏览器中完成验证码验证，然后按回车继续...")
        
        # 爬取前3页
        for page in range(3):
            print(f"正在爬取第{page + 1}页...")
            
            # 等待文章列表加载
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "news-box"))
            )
            
            # 获取所有文章项
            articles = driver.find_elements(By.CLASS_NAME, "news-box")
            
            # 解析每篇文章的信息
            for article in articles:
                try:
                    title = article.find_element(By.CSS_SELECTOR, "h3 a").text
                    link = article.find_element(By.CSS_SELECTOR, "h3 a").get_attribute("href")
                    summary = article.find_element(By.CLASS_NAME, "txt-info").text
                    source = article.find_element(By.CLASS_NAME, "account").text
                    
                    articles_data.append({
                        "标题": title,
                        "摘要": summary,
                        "链接": link,
                        "来源": source
                    })
                except Exception as e:
                    print(f"解析文章时出错: {e}")
                    continue
            
            # 如果不是最后一页，点击下一页
            if page < 2:
                try:
                    next_page = driver.find_element(By.ID, "sogou_next")
                    next_page.click()
                    time.sleep(10)  # 休眠10秒
                except Exception as e:
                    print(f"翻页时出错: {e}")
                    break
            
        # 将数据保存到Excel文件
        if articles_data:
            filename = f"AI_微信_{get_current_time()}.xlsx"
            filepath = os.path.join(content_folder, filename)
            df = pd.DataFrame(articles_data)
            df.to_excel(filepath, index=False, engine='openpyxl')
            print(f"数据已保存到: {filepath}")
            
    except Exception as e:
        print(f"发生错误: {e}")
    
    finally:
        driver.quit()

if __name__ == "__main__":
    scrape_articles()
