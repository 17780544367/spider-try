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
    # 获取当前脚本的绝对路径
    current_path = os.path.abspath(os.path.dirname(__file__))
    # 在当前目录下创建"内容"文件夹
    content_folder = os.path.join(current_path, "内容")
    
    # 如果文件夹不存在则创建
    if not os.path.exists(content_folder):
        try:
            os.makedirs(content_folder)
            print(f"创建文件夹成功: {content_folder}")
        except Exception as e:
            print(f"创建文件夹失败: {e}")
    else:
        print(f"文件夹已存在: {content_folder}")
    
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
        time.sleep(3)
        
        try:
            # 等待搜索框可见并输入
            search_box = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "query"))
            )
            search_box.clear()
            search_box.send_keys("AI")
            time.sleep(1)
            
            # 直接提交搜索，不需要点击文章标签
            search_box.send_keys(Keys.RETURN)
            time.sleep(5)  # 增加等待时间
            
            # 检查是否需要验证码
            if "验证码" in driver.page_source:
                input("请在浏览器中完成验证码验证，然后按回车继续...")
                time.sleep(3)
            
            # 爬取前3页
            for page in range(3):
                print(f"正在爬取第{page + 1}页...")
                
                # 等待文章列表加载
                try:
                    # 等待文章列表容器加载
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "main-left"))
                    )
                    time.sleep(2)
                    
                    # 滚动页面以加载所有文章
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    
                    # 使用新的选择器获取所有文章
                    articles = driver.find_elements(By.CSS_SELECTOR, ".news-list li")
                    # 如果上面的选择器无法找到文章，尝试备用选择器
                    if not articles:
                        articles = driver.find_elements(By.CSS_SELECTOR, "div.txt-box")
                    
                    print(f"找到{len(articles)}篇文章")
                    
                    # 遍历每篇文章
                    for article in articles:
                        try:
                            # 获取标题和链接
                            title_element = article.find_element(By.CSS_SELECTOR, "h3 a, .tit a")
                            title = title_element.text.strip()
                            link = title_element.get_attribute("href")
                            
                            # 获取摘要
                            try:
                                summary = article.find_element(By.CSS_SELECTOR, ".txt-info, .s-p").text.strip()
                            except:
                                try:
                                    summary = article.find_element(By.CSS_SELECTOR, "p").text.strip()
                                except:
                                    summary = "无摘要"
                            
                            # 获取来源
                            try:
                                source = article.find_element(By.CSS_SELECTOR, ".account, .s2").text.strip()
                            except:
                                source = "未知来源"
                            
                            if title:  # 只添加有标题的文章
                                articles_data.append({
                                    "标题": title,
                                    "摘要": summary,
                                    "链接": link,
                                    "来源": source
                                })
                                print(f"成功抓取文章: {title[:20]}...")
                                
                        except Exception as e:
                            print(f"解析文章时出错: {str(e)}")
                            # 保存当前文章的HTML以便调试
                            try:
                                with open(f"error_article_{len(articles_data)}.html", "w", encoding="utf-8") as f:
                                    f.write(article.get_attribute("outerHTML"))
                            except:
                                pass
                            continue
                    
                    # 如果不是最后一页，点击下一页
                    if page < 2:
                        try:
                            # 滚动到页面底部
                            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                            time.sleep(2)
                            
                            next_button = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.ID, "sogou_next"))
                            )
                            driver.execute_script("arguments[0].scrollIntoView();", next_button)
                            time.sleep(1)
                            next_button.click()
                            time.sleep(5)
                            
                        except Exception as e:
                            print(f"翻页失败: {str(e)}")
                            break
                            
                except Exception as e:
                    print(f"处理第{page + 1}页时出错: {str(e)}")
                    driver.save_screenshot(f"error_page_{page+1}.png")
                    break
            
            # 保存数据
            if articles_data:
                try:
                    filename = f"AI_微信_{get_current_time()}.xlsx"
                    filepath = os.path.join(content_folder, filename)
                    print(f"准备保存文件到: {filepath}")
                    
                    df = pd.DataFrame(articles_data)
                    df.to_excel(filepath, index=False, engine='openpyxl')
                    
                    if os.path.exists(filepath):
                        print(f"文件成功保存到: {filepath}")
                        print(f"共保存{len(articles_data)}条数据")
                    else:
                        print("文件保存失败，未找到生成的文件")
                        
                except Exception as e:
                    print(f"保存数据时出错: {str(e)}")
                    print(f"当前工作目录: {os.getcwd()}")
                    print(f"目标保存路径: {filepath}")
            else:
                print("未获取到任何数据，跳过保存步骤")
                
        except Exception as e:
            print(f"搜索操作失败: {str(e)}")
            driver.save_screenshot("error_screenshot.png")
            print("已保存错误截图")
            
    except Exception as e:
        print(f"发生错误: {str(e)}")
        driver.save_screenshot("error_final.png")
        
    finally:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    scrape_articles()
