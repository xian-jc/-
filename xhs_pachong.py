#差：滑动页面&页面存储（视频×与图片×）&信息存储√
from DrissionPage import ChromiumPage, ChromiumOptions
import time
from urllib.parse import urljoin
import openpyxl
import os
from openpyxl import load_workbook
import random
import uuid


#目前摘取的个数，同样也可以用作标签
temp_data = 0

filepath = "./result_add.xlsx"


def create_excel():
    workbook = openpyxl.Workbook()

    sheet = workbook.active

    sheet['A1'] = '作者网名'
    sheet['B1'] = '帖子类型'
    sheet['C1'] = '帖子的标题'
    sheet['D1'] = '帖子的文案'
    sheet['E1'] = '发帖的时间与ip'
    sheet['F1'] = '点赞数'
    sheet['G1'] = '收藏数'
    sheet['H1'] = '评论数'
    sheet['I1'] = '评论区前九条评论'
    sheet['J1'] = '帖子的视频或者图片链接'

    workbook.save(filepath)


#滑动页面(还会修改的)
def roll_page(tab):
    #取得帖子数没到指定数额时
    #触发：处理完同一行的三个帖子就下滑
    tab.scroll.to_bottom()
    #随机休眠
    #time.sleep(random(1.5,3))

def download_video(video_ele, page):
    try:
        # 确保目录存在
        media_dir = os.path.abspath("./media")
        os.makedirs(media_dir, exist_ok=True)
        
        # 获取视频URL（确保是字符串）
        video = video_ele.ele('tag:video', timeout=3)
        if not video:
            raise ValueError("未找到video元素")
            
        video_url = str(video.attr('src'))  # 强制转换为字符串
        if not video_url or not isinstance(video_url, str):
            raise ValueError("无效的视频URL")
            
        print(f"获取到视频URL: {video_url}")

        # 生成安全的文件名
        if video_url.startswith('blob:'):
            file_name = f"video_{int(time.time())}.mp4"
        else:
            file_name = os.path.basename(video_url.split('?')[0]) or f"video_{int(time.time())}.mp4"
        
        download_path = os.path.join(media_dir, file_name)
        print(f"准备下载到: {download_path}")

        # 处理blob URL
        if video_url.startswith('blob:'):
            js_code = """
            async function fetchBlob(blobUrl) {
                const response = await fetch(blobUrl);
                const blob = await response.blob();
                return await new Promise((resolve) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result.split(',')[1]);
                    reader.readAsDataURL(blob);
                });
            }
            """
            base64_data = page.run_js(f"{js_code} fetchBlob('{video_url}')")
            
            if not base64_data:
                raise ValueError("无法获取Blob数据")
                
            import base64
            with open(download_path, "wb") as f:
                f.write(base64.b64decode(base64_data))
        else:
            # 使用requests下载普通URL
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Referer': str(page.url)  # 确保是字符串
            }
            
            with requests.get(video_url, headers=headers, stream=True, timeout=30) as r:
                r.raise_for_status()
                with open(download_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)

        # 验证下载
        if os.path.exists(download_path) and os.path.getsize(download_path) > 1024:
            print(f"视频下载成功: {download_path}")
            return str(download_path)  # 确保返回字符串
        else:
            raise ValueError("下载文件无效")

    except Exception as e:
        print(f"视频下载失败: {str(e)}")
        if 'download_path' in locals() and os.path.exists(download_path):
            try:
                os.remove(download_path)
            except Exception as e:
                print(f"删除失败文件时出错: {e}")
        return None




#提取页面中的信息(帖子的详情页面)
def store_information(page):
    
    #所有的信息整合的列表
    whole_info = {}

    #整个帖子的详细界面
    whole_post_ele = page.ele('.note-detail-mask')

    
    #图片区/视频区(要修改！！！！)
    media_content = []
    
    #视频
    video_path = None
    #如果不存在这一类元素，怎么可以做到不报错
    video_ele = whole_post_ele.ele('.player-container', timeout = 0)
    #如果有视频
    if video_ele:
        os.makedirs("./media", exist_ok=True)
        video_path = download_video(video_ele, page)
        
        if video_path:
            whole_info['B1'] = "视频"
            whole_info['J1'] = video_path
        else:
            whole_info['B1'] = "视频"
            whole_info['J1'] = "视频下载失败"


    #图片√
    #可以拉取到图片的URL 但是存进文件夹的是其他类型的文件
    all_photo = ""
    photos = whole_post_ele.eles('.img-container', timeout = 0)
    if not photos:
        print("no photos")
    else:
        for photo in photos:
            img = photo.ele('tag:img')
            img_url = img.attr('src') or img.attr('data-src')
        
            if not img_url:
                continue
    
            full_url = urljoin(page.url, img_url)

            if all_photo == "":
                all_photo = full_url
            else:
                all_photo = all_photo + "\n" + full_url

            page.download(full_url, f'./images/post_{temp_data+ 8000}.jpg')

        whole_info['B1'] = "图片"
        whole_info['J1'] = all_photo
    
    
    
    #文字区 √
    message_ele = whole_post_ele.ele('.note-content')
    #标题（暂时只有提取）?元素是包括前面的这些信息还是只包括信息后的内容eg文字类
    title_ele = message_ele.ele('.title', timeout = 0)
    if title_ele:
        print(title_ele.text)
        whole_info['C1'] = title_ele.text
    else:
        whole_info['C1'] = ""


    #文案
    copy_ele = message_ele.ele('.desc').ele('.note-text')
    print(copy_ele.text)
    whole_info['D1'] = copy_ele.text

    #发布时间与地点
    loc_ele = message_ele.ele('.bottom-container')
    print(loc_ele.text)
    whole_info['E1'] = loc_ele.text
    
    
    #作者身份区 √
    author_ele = whole_post_ele.ele('.author-wrapper')
    #作者id
    id = author_ele.ele('.username')
    print(id.text)
    whole_info['A1'] = id.text
    
    #帖子传播度
    like_ele = whole_post_ele.ele('.like-wrapper like-active')
    like_count = like_ele.ele('.count')
    whole_info['F1'] = like_count.text
    collect_ele = whole_post_ele.ele('.collect-wrapper')
    collect_count = collect_ele.ele('.count')
    whole_info['G1'] = collect_count.text
    chat_ele = whole_post_ele.ele('.chat-wrapper')
    chat_count = chat_ele.ele('.count')
    whole_info['H1'] = chat_count.text

    #评论区√
    all_comment = ""
    comment_ele = whole_post_ele.ele('.comments-container', timeout = 0)
    if not comment_ele:
        print("no comment")
    else:

        #评论（前九条）
        for i in range (9):
            i = i + 1
            comment = comment_ele.ele(locator='.parent-comment', index=i, timeout = 0)

            if not comment:
                continue
            else:
                reviewer = comment.ele('.author-wrapper')
                comment_content = comment.ele('.note-text')
                #凭借成评论者：评论的形式
                whole_comment = reviewer.text + " : " + comment_content.text
                print(whole_comment)
                #print(reviewer.text)
                #print(comment_content.text)
                if all_comment == "":
                    all_comment = whole_comment
                else:
                    all_comment = all_comment + "\n" + whole_comment

    whole_info['I1'] = all_comment
    
    return whole_info


def store_in_excel(whole_info):
    data = whole_info

    row_data = [data.get('A1'), data.get('B1'), data.get('C1'), data.get('D1'), data.get('E1'), data.get('F1'), data.get('G1'), data.get('H1'), data.get('I1'), data.get('J1')]

    workbook = load_workbook(filepath)

    sheet = workbook.active

    sheet.append(row_data)

    workbook.save(filepath)


#对于选中的每一个帖子/视频的操作(进入->提取信息->退出)
def post_op(ele_post, page):
    #左键点击进去，注意这里的
    ele_post.click()
    #加载时间
    time.sleep(3)

    store_information(page)

    #store_in_excel(info)

    #点击退出键返回
    close_icon = page.ele('@@tag()=div@@class=close-circle')
    close_icon.click()
    



if __name__ == '__main__':
    #先生成存储爬虫结果的excel文档
    #filepath = "./result.xlsx"
    if not os.path.exists(filepath):
        create_excel()

    #打开浏览器
    browser = ChromiumPage(9333)

    #选择最近生成的页面
    tab = browser.latest_tab
    #打开小红书
    tab.get('https://www.xiaohongshu.com/explore')

    #先加载页面（可以调大一点）防止程序还没有选到对应的元素就进行了下一步
    time.sleep(10)

    while temp_data < 1000 :
        #查找当前页面符合条件的元素
        ele_post = tab.ele(f'xpath://section[@data-index="{temp_data}"]', timeout = 0)

        if ele_post:

            post_op(ele_post, tab)
            time.sleep(3.5)

            temp_data = temp_data + 1

        else:
            roll_page(tab)

'''
#单个
    ele_post = tab.ele(f'xpath://section[@data-index="{temp_data}"]')

    post_op(ele_post, tab)
'''

    