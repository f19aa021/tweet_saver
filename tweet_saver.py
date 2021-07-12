import json
import time
from datetime import datetime, timedelta, timezone
from requests_oauthlib import OAuth1Session
import pandas as pd
import openpyxl
import schedule
import smtplib
from email.mime.text import MIMEText
from email.utils import formatdate
import matplotlib.pyplot as plt

from keys import Keys

class GetTweet:
    CK = Keys.CK
    CS = Keys.CS
    AT = Keys.AT
    AS = Keys.AS

    twitter_api = OAuth1Session(CK, CS, AT, AS)

    url = 'https://api.twitter.com/1.1/search/tweets.json?tweet_mode=extended'
    # APIドキュメント
    # https://developer.twitter.com/en/docs/twitter-api/v1/tweets/search/api-reference/get-search-tweets

    # APIに必要なパラメーターの定義
    params =  {
        'lang': 'ja', 'result_type': 'recent'
    }
    
    JST = timezone(timedelta(hours=+9), 'JST')
    
    output_names = []
    num_tweet = []

    output_table = []

    # グラフ表示に必要なデータ
    plt.style.use('dark_background')
    fig = plt.figure('num_of_tweet / hour', figsize=(10, 5.8))
    
    ax = fig.add_subplot(111)
    ax.grid(True)
    ax.set_ylim(0, 180)
    lines, = ax.plot([], [], alpha=0.8, marker='8', color='#68DEEA')

    # 最大ツイート取得数を定義
    def __init__(self, count):
        self.params['count'] = count

    # 入力ファイルを「検索したいワードの配列」に変換
    def input_file(self, file_name):
        df_input = pd.read_excel(file_name)
        self.search_words = df_input.iloc[:, 0].values

    # 現時刻と１時間後の時刻を取得
    # 出力ファイルの名前やパスの定義
    def set_datetime(self):
        now_time = datetime.now(self.JST)
        one_hour_ago_time = now_time - timedelta(hours=1)

        now_time_str = now_time.strftime('%Y-%m-%d_%H:%M:%S_JST')
        one_hour_ago_time_str = one_hour_ago_time.strftime('%Y-%m-%d_%H:%M:%S_JST')
        
        self.params['since'] = one_hour_ago_time_str
        self.params['until'] = now_time_str

        self.output_name = now_time.strftime('%Y%m%d%H%M')

        self.output_names.append(now_time.strftime('%Y-%m-%d_%H:%M:%S'))

        # unix系
        # self.output_path = 'output_files/' + self.output_name + '.xlsx'
        # win
        self.output_path = 'output_files\\' + self.output_name + '.xlsx'

    # APIを使用してtweetを取得し任意のデータを出力ファイル用の２次元配列に追加
    def search(self):
        for search_word in self.search_words:
            self.params['q'] = search_word

            res = self.twitter_api.get(self.url, params=self.params)

            if res.status_code != 200:
                continue

            tweets = json.loads(res.text)
            for tweet in tweets['statuses']:
                line = []
                
                if 'retweeted_status' in tweet:
                    continue

                check_flag = None
                has_images = 'media' in tweet['entities']
                has_images_str = '有' if has_images else '無'
                full_text = tweet['full_text']
                screen_name = tweet['user']['screen_name']
                tweet_id = tweet['id_str']
                tweet_url = 'https://twitter.com/' + screen_name + '/status/' + tweet_id
                
                line = [check_flag, has_images_str, full_text, tweet_url]
                self.output_table.append(line)

    # 出力ファイル用の２次元配列をDataFrameに書き込みファイルを出力
    # 再度同ファイルを開き列幅の調整
    def output_file(self):
        df_output = pd.DataFrame(self.output_table, columns=['確認フラグ', '画像の有無', '本文', 'URL'])
        df_output.to_excel(self.output_path, header=False, index=False)

        wb_output = openpyxl.load_workbook(self.output_path)
        ws_1 = wb_output.worksheets[0]
        ws_1.column_dimensions['C'].width = 100
        ws_1.column_dimensions['D'].width = 50
        wb_output.save(self.output_path)

        print('Has been printed by ' + self.output_name)

        # その時間に取得できたツイート数を記録する
        self.num_tweet.append(len(self.output_table))

        # 出力ファイル用の２次元配列を空にする
        self.output_table = []

    # メールの作成
    def create_mail(self, from_addr, to_addr, subject, msg):
        body_msg = MIMEText(msg)
        body_msg['Subject'] = subject
        body_msg['From'] = from_addr
        body_msg['To'] = to_addr
        body_msg['Date'] = formatdate()
        return body_msg

    # メールの送信
    def send_mail(self, from_addr, to_addr, body_msg):
        smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpobj.ehlo()
        smtpobj.starttls()
        smtpobj.ehlo()
        smtpobj.login(from_addr, '120sa91F')
        smtpobj.sendmail(from_addr, to_addr, body_msg.as_string())
        smtpobj.close()

    # 時間ごとの取得ツイート数のグラフを表示（作成中）
    def display_graph(self):
        x = self.output_names
        y = self.num_tweet

        self.lines.set_data(x[-5:], y[-5:])
        self.ax.set_xlim(len(x)-1, len(x)+3)
        plt.pause(6)

def save(get_tweet):
    get_tweet.set_datetime()
    get_tweet.search()
    get_tweet.output_file()
    # get_tweet.display_graph()

def main():
    get_tweet = GetTweet(count=180)
    get_tweet.input_file(file_name='input.xlsx')
    # 標準処理
    try:
        save(get_tweet)
        # saveを一時間ごとに実行
        schedule.every(1).hours.do(save, get_tweet)
        while True:
            schedule.run_pending()
            time.sleep(1)
    # エラー処理
    except Exception as e:
        # エラーをコマンドラインに表示
        print(e)

        # 現時刻取得
        stop_time = datetime.now(get_tweet.JST)
        stop_time_str = stop_time.strftime('%Y-%m-%d_%H:%M:%S_JST')

        # メールの情報定義
        from_addr = 'f19aa021@gmail.com'
        # to_addr = 'f19aa021@gmail.com,f19aa021@chuo.ac.jp'
        to_addr = 'f19aa021@gmail.com'
        subject = 'ツイート自動収集_システム停止通知'

        # メールの文を定義したファイルを開きメールを作成
        with open('./mail_content.txt') as f:
            msg = f.read().format(stop_time_str, e)
            body_msg = get_tweet.create_mail(from_addr, to_addr, subject, msg)

        # メールを送信
        get_tweet.send_mail(from_addr, to_addr.split(','), body_msg)

if __name__ == '__main__':
    main()