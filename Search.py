import Ranking
import pandas as pd
import itchat


def main():
    while 1:
        print('=' * 20)
        print('请输入指令：')
        cmd = input()
        print('=' * 20)
        # -------------------
        if cmd == '说明':
            # ---------------
            print('程序说明\n1.输入"search"进入抓取分支：复制进关键词(以"#"结束，回车分隔)\n2.输入"query"进入多关键词搜索分支\n3.title/domin排名全查询('
                  '以"title:"/"domain:"开始)\n4.输入"微信登录"，扫码后发送"查询结果"给文件传输助手\n5.')
        # -------------------
        if 'search' in cmd:
            # ---多关键词分支
            print('- 多关键词抓取分支 -')
            print('- 以#结束 -')
            keywords = []
            # 多行输入关键词
            while 1:
                k = input()
                if k == '#':
                    print('即将抓取"{k}"...请确认(y/n)'.format(k=keywords))
                    if input() == 'y':
                        print('- 开始抓取 -')
                        Ranking.scrape_multi(keywords)
                        print('- 抓取完毕 -')
                        break
                    else:
                        break
                else:
                    keywords.append(k)
        if cmd == 'query':
            print('请选择：1.标题匹配/2.域名匹配')
            choose = input()
            if choose == '1':
                print('- 多关键词标题排名匹配分支 -')
                print('- 以#结束 -')
                match_title = []
                while 1:
                    k = input()
                    if k == '#':
                        print('即将匹配"{k}"...请确认(y/n)'.format(k=match_title))
                        if input() == 'y':
                            print('- 开始匹配 -')
                            frame = pd.read_excel('./data/keyword-multi.xlsx')
                            match_title_function(frame, match_title)
                            print('- 匹配完毕 -')
                            break
                        else:
                            break
                    else:
                        match_title.append(k)
            # -------------------------------
            if choose == '2':
                print('- 多关键词域名排名匹配分支 -')
                print('- 以#结束 -')
                match_domain = []
                while 1:
                    k = input()
                    if k == '#':
                        print('即将匹配"{k}"...请确认(y/n)'.format(k=match_domain))
                        if input() == 'y':
                            print('- 开始匹配 -')
                            frame = pd.read_excel('./data/keyword-multi.xlsx')
                            match_domain_function(frame, match_domain)
                            print('- 匹配完毕 -')
                            break
                        else:
                            break
                    else:
                        match_domain.append(k)
            # ---------------
        if 'title:' in cmd:
            key = cmd.replace('title:', '').strip()
            try:
                frame = pd.read_excel('./data/keyword-multi.xlsx')
                Search_Keyword_Multi(frame, key)
            except Exception:
                print('- 文件不存在 -')
        # ------------------------
        if 'domain:' in cmd:
            key = cmd.replace('domain:', '').strip()
            try:
                frame = pd.read_excel('./data/keyword-multi.xlsx')
                Search_Domain_Multi(frame, key)
            except Exception:
                print('- 文件不存在 -')
        if cmd == '微信登录':
            itchat.auto_login(enableCmdQR=2)
            itchat.run()
        if cmd == 'exit':
            break


@itchat.msg_register('Text')
# 自动回复
def text_reply(msg):
    if msg['Text'] == '查询结果':
        # print('- 按任意键发送查询Results -')
        # input()
        itchat.send_file('./data/Results.xlsx', 'filehelper')
        print('- 文件已发送 -')
    elif msg['Text'] == '退出微信':
        itchat.logout()
        print('- 微信已退出 -')
    else:
        itchat.send_msg('发送"查询结果"获取Excel结果\n发送"退出微信"退出登录', 'filehelper')
        print('发送"查询结果"获取Excel结果\n发送"退出微信"退出登录')


# ----------------匹配函数----------------------------
def match_title_function(frame, match_title):
    title_list = []
    domain_list = []
    rank_list = []
    Keyword = []
    i = 0
    for item in match_title:
        tmpframe = frame[frame.keyword == frame.keyword.unique()[i]]
        i += 1
        for index, row in tmpframe.iterrows():
            try:
                if item in row.title:
                    print('关键词 "', item, '"在{k}的排名为:'.format(k=row['keyword']), row['rank'])
                    Keyword.append(row.keyword)
                    rank_list.append(row['rank'])
                    title_list.append(row.title)
                    domain_list.append(row.domain)
            except Exception:
                pass
                # ------------------------
    tmp_dict = {
        'Keyword': Keyword,
        'rank': rank_list,
        'title': title_list,
        'domain': domain_list,
    }
    # -----------------------
    tmp_frame = pd.DataFrame(tmp_dict)
    tmp_frame.to_excel('./data/Results.xlsx')
    print(tmp_frame)


def match_domain_function(frame, match_domain):
    title_list = []
    domain_list = []
    rank_list = []
    Keyword = []
    i = 0
    for item in match_domain:
        tmpframe = frame[frame.keyword == frame.keyword.unique()[i]]
        i += 1
        for index, row in tmpframe.iterrows():
            try:
                if item in row.domain:
                    print('域名 "', item, '"在{k}的排名为:'.format(k=row['keyword']), row['rank'])
                    Keyword.append(row.keyword)
                    rank_list.append(row['rank'])
                    title_list.append(row.title)
                    domain_list.append(row.domain)
            except Exception:
                pass
                # ------------------------
    tmp_dict = {
        'Keyword': Keyword,
        'rank': rank_list,
        'title': title_list,
        'domain': domain_list,
    }
    # -----------------------
    tmp_frame = pd.DataFrame(tmp_dict)
    tmp_frame.to_excel('./data/Results.xlsx')
    print(tmp_frame)


# -------------------------------------------


def Search_Keyword_Multi(frame, keyword):
    title_list = []
    domain_list = []
    rank_list = []
    Keyword = []
    for index, row in frame.iterrows():
        try:
            if keyword in row.title:
                print('关键词 "', keyword, '"在{k}的排名为:'.format(k=row['keyword']), row['rank'])
                Keyword.append(row.keyword)
                rank_list.append(row['rank'])
                title_list.append(row.title)
                domain_list.append(row.domain)
        except Exception:
            pass
        # ------------------------
    tmp_dict = {
        'Keyword': Keyword,
        'rank': rank_list,
        'title': title_list,
        'domain': domain_list,
    }
    # -----------------------
    tmp_frame = pd.DataFrame(tmp_dict)
    tmp_frame.to_excel('./data/Results.xlsx')
    print(tmp_frame)


def Search_Domain_Multi(frame, word):
    title_list = []
    domain_list = []
    rank_list = []
    Keyword = []
    for index, row in frame.iterrows():
        try:
            if word in row['domain']:
                print('域名 "', word, '"在{k}的排名为:'.format(k=row['keyword']), row['rank'])
                Keyword.append(row['keyword'])
                rank_list.append(row['rank'])
                title_list.append(row['title'])
                domain_list.append(row['domain'])
        except Exception:
            pass

        # ------------------------
    tmp_dict = {
        'Keyword': Keyword,
        'rank': rank_list,
        'title': title_list,
        'domain': domain_list,
    }
    # -----------------------
    tmp_frame = pd.DataFrame(tmp_dict)
    tmp_frame.to_excel('./data/Results.xlsx')
    print(tmp_frame)


if __name__ == '__main__':
    main()
