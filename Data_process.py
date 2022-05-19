import xlrd
import pandas as pd


# 格式帖子状态数据（保持原样）1
def extractstate(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    state = []
    for i in range(1, nrows):  # 第0行为表头
        mi = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = mi[1]  # 取出表中列数据
        state.append(result)
    return state


# 格式处理用户名数据2
def extractusername(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    usernameList = []
    username_List = []
    for i in range(1, nrows):  # 第0行为表头
        username = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = username[2]  # 取出表中列数据
        usernameList.append(result)
    for i in usernameList:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        name = ''.join(name_list)
        username_List.append(name)
    return username_List


# 格式处理用户ID数据（保持原样即可）3
def extractuserid(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    userIDList = []
    for i in range(1, nrows):  # 第0行为表头
        id = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = id[3]  # 取出表中列数据
        userIDList.append(result)
    return userIDList


# 格式处理用户身份数据4
def extractuserinfor(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    userinfoList = []
    user_infoList = []
    for i in range(1, nrows):  # 第0行为表头
        userinfo = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = userinfo[4]  # 取出表中列数据
        userinfoList.append(result)
    for i in userinfoList:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count(' ')
        for i in range(0, count):
            name_list.remove(' ')
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        count = name_list.count(",")
        for i in range(0, count):
            name_list.remove(",")
        name = ''.join(name_list)
        user_infoList.append(name)
    return user_infoList


# 格式处理评论主题数据5
def extractcommentlist(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    commentList = []
    comment_List = []
    for i in range(1, nrows):  # 第0行为表头
        comment = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = comment[5]  # 取出表中列数据
        commentList.append(result)
    for i in commentList:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        name = ''.join(name_list)
        comment_List.append(name)
    return comment_List


# 格式处理评论内容数据6
def extractcommentcon(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    commentList = []
    comment_List = []
    for i in range(1, nrows):  # 第0行为表头
        comment = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = comment[6]  # 取出表中列数据
        commentList.append(result)
    for i in commentList:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        name = ''.join(name_list)
        comment_List.append(name)
    return comment_List


# 格式处理评论时间数据7
def extractcommenttime(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    commentTime = []
    commentTimeList = []
    for i in range(1, nrows):  # 第0行为表头
        time = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = time[7]  # 取出表中列数据
        commentTime.append(result)
    for i in commentTime:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        count = name_list.count('发表')
        for i in range(0, count):
            name_list.remove('发表')
        name = ''.join(name_list)
        commentTimeList.append(name)
    return commentTimeList


# 格式处理浏览次数数据8
def extractwatch_num(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    watch_num = []
    watch_numList = []
    for i in range(1, nrows):  # 第0行为表头
        areaAFpiece = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = areaAFpiece[8]  # 取出表中列数据
        watch_num.append(result)
    for i in watch_num:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        name = ''.join(name_list)
        watch_numList.append(name)
    return watch_numList


# 格式处理回复次数数据9
def extractreply_num(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    reply_num = []
    reply_numList = []
    for i in range(1, nrows):  # 第0行为表头
        areaAFpiece = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = areaAFpiece[9]  # 取出表中列数据
        reply_num.append(result)
    for i in reply_num:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        name = ''.join(name_list)
        reply_numList.append(name)
    return reply_numList


# 格式处理用户个人主页数据（保持原样即可）8//未使用
def extractuser_index(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    userindex_List = []
    for i in range(1, nrows):  # 第0行为表头
        userindex = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = userindex[8]  # 取出表中列数据
        userindex_List.append(result)
    return userindex_List


# 格式处理投票数据10
def extractvote(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    reply_num = []
    reply_numList = []
    for i in range(1, nrows):  # 第0行为表头
        areaAFpiece = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = areaAFpiece[10]  # 取出表中列数据
        reply_num.append(result)
    for i in reply_num:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        name = ''.join(name_list)
        reply_numList.append(name)
    return reply_numList


# 格式处理优秀证书数据11
def extractexcellentcertificate(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    excellentcertificate = []
    excellent_certificate = []
    for i in range(1, nrows):  # 第0行为表头
        certi = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = certi[11]  # 取出表中列数据
        excellentcertificate.append(result)
    for i in excellentcertificate:
        name = ''.join(str(a) for a in i)
        name_list = list(name)
        count = name_list.count('[')
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        count = name_list.count("\\")
        for i in range(0, count):
            name_list.remove("\\")
        count = name_list.count('n')
        for i in range(0, count):
            name_list.remove('n')
        name = ''.join(name_list)
        excellent_certificate.append(name)
    return excellent_certificate


# 格式处理合格证书数据12
def extractpasscertificate(commentpath):
    data = xlrd.open_workbook(commentpath, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号
    ncols = table.ncols  # 获取列号
    passcertificate = []
    pass_certificate = []
    for i in range(1, nrows):  # 第0行为表头
        certi = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = certi[12]  # 取出表中列数据
        passcertificate.append(result)
    for i in passcertificate:
        name = ''.join(str(a) for a in i)
        name_list = list(name)  # 字符串转List
        count = name_list.count('[')  # 删除’[],\n‘
        for i in range(0, count):
            name_list.remove('[')
        count = name_list.count(']')
        for i in range(0, count):
            name_list.remove(']')
        count = name_list.count("'")
        for i in range(0, count):
            name_list.remove("'")
        count = name_list.count("\\")
        for i in range(0, count):
            name_list.remove("\\")
        count = name_list.count('n')
        for i in range(0, count):
            name_list.remove('n')
        name = ''.join(name_list)
        # print(name)
        pass_certificate.append(name)
    return pass_certificate


# 对数据进行格式处理形成新的pd数据
def dataProcess(commentpath):
    state = extractstate(commentpath)
    watch_numList = extractwatch_num(commentpath)
    reply_numList = extractreply_num(commentpath)
    # userindex_List=extractuser_index(commentpath)
    userIDList = extractuserid(commentpath)
    usernameList = extractusername(commentpath)
    user_infoList = extractuserinfor(commentpath)
    comment_theme = extractcommentlist(commentpath)
    comment_con = extractcommentcon(commentpath)
    commentTimeList = extractcommenttime(commentpath)
    vote_num = extractvote(commentpath)
    excellent_certificate = extractexcellentcertificate(commentpath)
    pass_certificate = extractpasscertificate(commentpath)
    tf = pd.DataFrame({
        '状态': state,
        '用户名': usernameList,
        'ID': userIDList,
        '用户身份': user_infoList,
        '评论主题': comment_theme,
        '评论内容': comment_con,
        '评论时间': commentTimeList,
        '浏览次数': watch_numList,
        '回复次数': reply_numList,
        '投票数': vote_num,
        # '用户主页网址:': userindex_List,
        '优秀证书': excellent_certificate,
        '合格证书': pass_certificate,

    })
    return tf


def Process(before, after):  # before = 'test.xlsx'
    tf = dataProcess(before)
    tf.to_excel(after)
    return None


if __name__ == '__main__':
    Process('test5_10_1.xlsx', 'test5_10_1_DP.xlsx')
