from selenium import webdriver
import time
import pandas as pd
import re
from lxml import etree
import json
import Data_process as pr
from selenium.webdriver.common.by import By


############################################################################################################

login_url = "https://www.icourse163.org/member/login.htm#/webLoginIndex"  # MOOC登录地址
comment_url = 'https://www.icourse163.org/learn/BIT-268001?tid=1207014257#/learn/forumindex'    # MOOC讨论区地址/Python162页讨论区
path = "chromedriver.exe"  # Chromedriver地址
email = "youremail"
password = "yourpassword"
driver = webdriver.Chrome(executable_path=path)

############################################################################################################


def login():
    driver.get(login_url)
    time.sleep(1)
    # 先找到其他方式登录按钮
    jump = driver.find_element_by_class_name('ux-login-set-scan-code_ft_back')
    # 然后点击跳转到其他方式登录界面
    jump.click()

    iframe = driver.find_elements_by_tag_name("iframe")[0]
    driver.switch_to.frame(iframe)

    driver.find_element_by_name("email").send_keys(email)
    time.sleep(1)
    driver.find_element_by_name("password").send_keys(password)
    time.sleep(1)
    driver.find_element_by_id("dologin").click()  # 点击登录
    time.sleep(1)

    driver.switch_to.default_content()  # 退出框架

    cookies = driver.get_cookies()  # 获取cookie,列表形式
    f1 = open('cookies.txt', 'w')
    f1.write(json.dumps(cookies))
    f1.close()
    driver.close()
    return None


def cookie():
    driver.get(login_url)
    f2 = open("cookies.txt")
    cookies = json.loads(f2.read())
    # 使用cookies登录
    for cook in cookies:
        driver.add_cookie(cook)
    # 刷新页面
    driver.refresh()
    time.sleep(2)
    return None


def getCommentPageNum():
    driver.get(comment_url)
    time.sleep(1)  # 没有的话可能获取的内容为空
    content = driver.page_source
    dom = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
    page_list = dom.xpath(
        '//*[@id="courseLearn-inner-box"]/div/div[7]/div/div[2]/div/div[1]/div[2]/div/a//text()')  # 页码列表
    print(page_list)
    page_Num = page_list[-2]  # 最后一个标签的文本内容即总页码
    print('该课程一共有' + page_Num + '页评论')
    return page_Num


def getpageInfo(url_head, pagenum):
    df_all = pd.DataFrame()
    # for page_num in range(0, pagenum):
    for page_num in range(0, pagenum):
        try:
            url = url_head + '?t=0&p={}'.format(page_num + 1)
            print('正在获取第{}页的数据'.format(page_num + 1))  # 打印进度
            tmp = get_comment_detail(url)  # 调用函数
            df_all = df_all.append(tmp, ignore_index=True)
            print('第{}页数据读取完成'.format(page_num + 1))  # 打印进度
        except Exception as e:
            print("Exception1 had happen")
            break
    return df_all


def get_comment_detail(url_head):
    namesList = []  # 发表评论的用户名列表
    commentTime = []  # 用户评论时间
    commentList = []  # 评论主题
    comContent = []  # 评论具体内容
    watch_numList = []  # 评论浏览次数
    reply_numList = []  # 评论回复次数
    vote_num = []  # 投票数
    comment_href = []
    user_href = []
    user_indexList = []  # 评论者个人详情页
    user_infoList = []  # 评论者信息
    user_ID = []  # 评论者ID
    comment_index = []  # 进入评论贴子详情界面的链接
    excellent_certificate = []  # 优秀证书
    pass_certificate = []  # 通过证书
    post_state = []  # 帖子状态，是原始贴还是回复贴

    driver.get(url_head)
    print("成功进入该课程讨论区")
    time.sleep(1)  # 没有的话可能获取的内容为空
    content = driver.page_source
    dom = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
    comment_l = dom.xpath('//*[@id="courseLearn-inner-box"]/div/div[7]/div/div[2]/div/div[1]/div[1]/li')

    # 获取每页评论列表
    for li in comment_l:
        commentList.append(li.xpath('./div/a/text()'))
        comment_href.append(li.xpath('./div/a/@href'))
        watch_numList.append(li.xpath('./p[1]/text()'))
        reply_numList.append(li.xpath('./p[2]/text()'))
        vote_num.append(li.xpath('./p[3]/text()'))
        commentTime.append(li.xpath('./span/span[1]/span[2]/text()'))
        href = li.xpath('./div/a/@href')  # 进入评论回复的地址[#/learn/forumdetail?pid=1002189262]
        replynum = li.xpath('./p[2]/text()')  # 回复数
        commenthref = comment_url[:-18] + href[0]  # 拼接成一个完整的链接
        name = li.xpath('./span/span[1]/span[1]/span/span[2]/a/@title')
        ushref = li.xpath('./span/span[1]/span[1]/span/span[2]/a/@href')
        # 需要注意有些用户是匿名发表
        if not name:
            namesList.append(['匿名'])
            user_indexList.append('匿名')
            ID = '匿名'
            user_infoList.append('匿名')
            excellent_certificate.append('匿名')
            pass_certificate.append('匿名')
        else:
            namesList.append(name)
            for j in ushref:
                s = ' '.join(str(i) for i in j)  # /learn/forumpersonal?uid=1028283590"
                st = re.findall('([0-9]{1,15})', s)  # 此方法返回的是列表["1028283590"]
                ID = ''.join(st)
            excellent_certificate_url = 'https://www.icourse163.org/home.htm?userId=' + str(
                ID) + '#/home/mycert?userId=' + str(ID) + '+&type=1&p=1'
            pass_certificate_url = 'https://www.icourse163.org/home.htm?userId=' + str(
                ID) + '#/home/mycert?userId=' + str(ID) + '+&type=2&p=1'
            driver.get(excellent_certificate_url)
            driver.refresh()
            time.sleep(.5)
            content = driver.page_source
            dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
            mi = dom2.xpath('//*[@id="j-mycert-body"]/div/div[2]/div/div/div/div[1]/text()')
            if not mi:
                excellent_certificate.append('无')
            else:
                excellent_certificate.append(mi)
            driver.get(pass_certificate_url)
            driver.refresh()
            time.sleep(.5)
            content = driver.page_source
            dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
            mi = dom2.xpath('//*[@id="j-mycert-body"]/div/div[2]/div/div/div/div[1]/text()')
            if not mi:
                pass_certificate.append('无')
            else:
                pass_certificate.append(mi)
            print("原始", commentList[-1], excellent_certificate[-1], pass_certificate[-1])
            dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
            mi = dom2.xpath('//*[@id="j-self-content"]/div/div[3]/span/text()')
            if not mi:
                mi1 = dom2.xpath('//*[@id="j-self-content"]/div/div[4]/span/a/text()')
                mi2 = dom2.xpath('//*[@id="j-self-content"]/div/div[4]/span/span[2]/text()')
                mi = mi1 + mi2
                user_infoList.append(mi)
            else:
                user_infoList.append(mi)
        # print('ID:'+ ID)
        user_href.append(ushref)
        user_ID.append(ID)

        driver.get(commenthref)  # 进入该评论
        time.sleep(1)
        content = driver.page_source
        dom = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
        comContent.append(dom.xpath('//*[@id="courseLearn-inner-box"]/div/div[2]/div/div[1]/div/div[2]//text()'))
        print(commenthref)
        page = dom.xpath(
            '//*[@id="courseLearn-inner-box"]/div/div[2]/div/div[4]/div/div[1]/div[2]//text()')  # 回复页码数
        print(page)
        post_state.append('原始')
        if replynum == ['回复：0']:  # 这个应该放在最后面
            None
        else:  # 若回复数不为0，则进入具体评论，获取回复帖子
            page_num = page[-2]
            print('当前回复共有' + page_num + '页')
            page_num = int(page_num)
            for p in range(0, page_num):
                # try:
                content = driver.page_source
                dom = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                com = dom.xpath('//*[@id="courseLearn-inner-box"]/div/div[2]/div/div[4]/div/div[1]/div[1]/div')
                for di in com:
                    reply = di.xpath('./div/div[2]//text()')  # 回复贴的内容
                    vote = di.xpath('./div/div[3]/div[7]/div/p/text()')  # 投票情况，赞/踩数
                    timee = di.xpath('./div/div[3]/div[2]/text()')  # 回复贴的时间
                    userName = di.xpath('./div/div[3]/div[1]/span/span[2]/a/text()')  # 回复贴的用户名
                    comNum = di.xpath('./div/div[3]/a[3]/text()')  # 回复贴的评论数
                    uhref = di.xpath(
                        './div/div[3]/div/span/span[2]/a/@href	')  # 返回的是#/learn/forumpersonal?uid=12068573
                    if userName == []:
                        userName = '匿名'
                        ID = '匿名'
                        uinfo = '匿名'
                        ec = '匿名'
                        pc = '匿名'
                    else:
                        for j in uhref:
                            s = ' '.join(str(i) for i in j)  # /learn/forumpersonal?uid=1028283590"
                            st = re.findall('([0-9]{1,15})', s)  # 此方法返回的是列表["1028283590"]
                            ID = ''.join(st)
                        excellent_certificate_url = 'https://www.icourse163.org/home.htm?userId=' + str(
                            ID) + '#/home/mycert?userId=' + str(ID) + '+&type=1&p=1'
                        pass_certificateurl = 'https://www.icourse163.org/home.htm?userId=' + str(
                            ID) + '#/home/mycert?userId=' + str(ID) + '+&type=2&p=1'
                        driver.get(excellent_certificate_url)
                        driver.refresh()
                        time.sleep(.5)
                        content = driver.page_source
                        dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                        mi = dom2.xpath('//*[@id="j-mycert-body"]/div/div[2]/div/div/div/div[1]/text()')
                        if not mi:
                            ec = '无'
                        else:
                            ec = mi
                        driver.get(pass_certificateurl)
                        driver.refresh()
                        time.sleep(.5)
                        content = driver.page_source
                        dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                        mi = dom2.xpath('//*[@id="j-mycert-body"]/div/div[2]/div/div/div/div[1]/text()')
                        if not mi:
                            pc = '无'
                        else:
                            pc = mi
                        dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                        mi = dom2.xpath('//*[@id="j-self-content"]/div/div[3]/span/text()')
                        if not mi:
                            time.sleep(.5)
                            mi1 = dom2.xpath('//*[@id="j-self-content"]/div/div[4]/span/a/text()')
                            mi2 = dom2.xpath('//*[@id="j-self-content"]/div/div[4]/span/span[2]/text()')
                            mi = mi1 + mi2
                            uinfo = mi
                        else:
                            uinfo = mi
                    print("回复", reply, ec, pc, vote, timee, userName, comNum)

                    post_state.append('回复')
                    namesList.append(userName)
                    user_ID.append(ID)
                    user_infoList.append(uinfo)
                    commentList.append('null')
                    comContent.append(reply)
                    commentTime.append(timee)
                    watch_numList.append('null')
                    reply_numList.append(comNum)
                    vote_num.append(vote)
                    excellent_certificate.append(ec)
                    pass_certificate.append(pc)
                    if comNum != ['评论(0)']:  # 评论数不为0
                        con = di.xpath('./div[2]/div/div/div/div/div/div')
                        le = len(con)
                        for ei in con:  # 会多出一组空数据，跳出最后一次循环，避免写入多余的无用数据
                            if le == 1:
                                break
                            reply1 = ei.xpath('./div/div[2]//text()')  # 评论贴的内容
                            vote1 = ei.xpath('./div/div[3]/div[7]/div/p/text()')  # 投票情况，赞/踩数
                            timee1 = ei.xpath('./div/div[3]/div[2]/text()')  # 评论贴的时间
                            userName1 = ei.xpath('./div/div[3]/div[1]/span/span[2]/a/text()')  # 评论贴的用户名
                            uhref1 = ei.xpath(
                                './div/div[3]/div/span/span[2]/a/@href')  # 返回的是#/learn/forumpersonal?uid=12068573
                            if not userName1:
                                userName1 = '匿名'
                                ID1 = '匿名'
                                uinfo1 = '匿名'
                                ec1 = '匿名'
                                pc1 = '匿名'
                            else:
                                for j in uhref1:
                                    s = ' '.join(str(i) for i in j)  # /learn/forumpersonal?uid=1028283590"
                                    st = re.findall('([0-9]{1,15})', s)  # 此方法返回的是列表["1028283590"]
                                    ID1 = ''.join(st)
                                excellent_certificate_url = 'https://www.icourse163.org/home.htm?userId=' + str(
                                    ID1) + '#/home/mycert?userId=' + str(ID1) + '+&type=1&p=1'
                                pass_certificateurl = 'https://www.icourse163.org/home.htm?userId=' + str(
                                    ID1) + '#/home/mycert?userId=' + str(ID1) + '+&type=2&p=1'
                                driver.get(excellent_certificate_url)
                                driver.refresh()
                                time.sleep(.5)
                                content = driver.page_source
                                dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                                mi = dom2.xpath('//*[@id="j-mycert-body"]/div/div[2]/div/div/div/div[1]/text()')
                                if not mi:
                                    ec1 = '无'
                                else:
                                    ec1 = mi
                                driver.get(pass_certificateurl)
                                driver.refresh()
                                time.sleep(.5)
                                content = driver.page_source
                                dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                                mi = dom2.xpath('//*[@id="j-mycert-body"]/div/div[2]/div/div/div/div[1]/text()')
                                if not mi:
                                    pc1 = '无'
                                else:
                                    pc1 = mi
                                dom2 = etree.HTML(content, etree.HTMLParser(encoding='utf-8'))
                                mi = dom2.xpath('//*[@id="j-self-content"]/div/div[3]/span/text()')
                                if not mi:
                                    time.sleep(.5)
                                    mi1 = dom2.xpath('//*[@id="j-self-content"]/div/div[4]/span/a/text()')
                                    mi2 = dom2.xpath('//*[@id="j-self-content"]/div/div[4]/span/span[2]/text()')
                                    mi = mi1 + mi2
                                    uinfo1 = mi
                                else:
                                    uinfo1 = mi

                            print("评论", reply1, ec1, pc1, vote1, timee1, userName1)

                            post_state.append('评论')
                            namesList.append(userName1)
                            user_ID.append(ID1)
                            user_infoList.append(uinfo1)
                            commentList.append('null')
                            comContent.append(reply1)
                            commentTime.append(timee1)
                            watch_numList.append('null')
                            reply_numList.append('null')
                            vote_num.append(vote1)
                            excellent_certificate.append(ec1)
                            pass_certificate.append(pc1)
                            le = le - 1
                if (page_num - p) != 1:  # 非最后一页时进行翻页
                    driver.get(commenthref)
                    time.sleep(1)
                    for i in range(0, p + 1):
                        jump = driver.find_element_by_xpath(  # 通过Xpath来找到元素"下一页"的位置
                            '//*[@id="courseLearn-inner-box"]/div/div[2]/div/div[4]/div/div[1]/div[2]/div/a[11]')  # success nice!
                        time.sleep(.5)
                        jump.click()  # 不同回复页面网页地址相同，只能通过点击下一页进行切换

                time.sleep(1)
               
    tmp = pd.DataFrame({
        '状态': post_state,
        '用户名': namesList,
        'ID': user_ID,
        '用户身份': user_infoList,
        '评论主题': commentList,
        '评论内容': comContent,
        '评论时间': commentTime,
        '浏览次数': watch_numList,
        '回复次数': reply_numList,
        '投票数': vote_num,
        '优秀证书': excellent_certificate,
        '合格证书': pass_certificate,
    })
    return tmp


if __name__ == '__main__':
    # login()
    cookie()
    pagenum = getCommentPageNum()
    pagenum = int(pagenum)
    df_all = getpageInfo(comment_url, pagenum)
    df_all.to_excel('test.xlsx')
    print('数据写入完成，正在处理格式')
    pr.Process('test.xlsx', 'test_DP.xlsx')     # 数据处理
    print('数据处理完成')
    time.sleep(3)
    driver.quit()
