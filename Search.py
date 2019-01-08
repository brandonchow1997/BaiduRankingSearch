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
        if cmd == '百度排名程序说明':
            # ---------------
            print('程序说明\n1.输入关键词:(以#开始，多个关键词以逗号[,]分隔)\n2.title/domin排名查询\n(以title+:/domain+:开始)')
        # -------------------
        if '#' in cmd:
            # ---多关键词分支
            if ',' in cmd:
                print('进入多关键词分支')
                keywords = cmd.replace('#', '').strip().split(',')
                print('即将抓取"{k}"...请确认(y/n)'.format(k=keywords))
                if input() == 'y':
                    print('开始抓取')
                    Ranking.scrape_multi(keywords)
                else:
                    print('- 请重试 -')
                    pass
            # ---------------
        if 'title+:' in cmd:
            key = cmd.replace('title+:', '').strip()
            try:
                frame = pd.read_excel('./data/keyword-multi.xlsx')
                Search_Keyword_Multi(frame, key)
            except Exception:
                print('- 文件不存在 -')
        # ------------------------
        if 'domain+:' in cmd:
            key = cmd.replace('domain+:', '').strip()
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
        print('- 按任意键发送查询Results -')
        input()
        itchat.send_file('./data/Results.xlsx', 'filehelper')
        print('- 文件已发送 -')
    if msg['Text'] == '退出微信':
        itchat.logout()
        print('- 微信已退出 -')


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
