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

class GetTweet:
    # 初期化
    # 現時刻と１時間後の時刻を取得
    # APIの認証とパラメーターの定義
    # 出力ファイルの名前やパスの定義
    def __init__(self, count):
        JST = timezone(timedelta(hours=+9), 'JST')
        now_time = datetime.now(JST)
        one_hour_ago_time = now_time - timedelta(hours=1)

        now_time_str = now_time.strftime('%Y-%m-%d_%H:%M:%S_JST')
        one_hour_ago_time_str = one_hour_ago_time.strftime('%Y-%m-%d_%H:%M:%S_JST')

        CK = 'NKu57EZlWMKtNHAws3OWY0NHb'
        CS = 'fKkG9o22A3lObWBF2mNRoxBWzhhjgWhTpHkpHyn8kUttVMvqrT'
        AT = '1348856236062560256-ofOP557oWNFsfKkXG7zNisw8n5s7Dy'
        AS = 'SqNyZv9Lsejn4imIeELHGE7u2200oAvWwpTjfInUQ3m8G'

        self.twitter_api = OAuth1Session(CK, CS, AT, AS)
        
        self.url = 'https://api.twitter.com/1.1/search/tweets.json?tweet_mode=extended'
        # APIドキュメント
        # https://developer.twitter.com/en/docs/twitter-api/v1/tweets/search/api-reference/get-search-tweets
        self.params =  {
            'lang': 'ja', 'count': count, 'result_type': 'recent',
            'since': one_hour_ago_time_str, 'until': now_time_str}

        self.output_table = []

        self.output_name = now_time.strftime('%Y%m%d%H%M')
        # unix系
        self.output_path = 'output_files/' + self.output_name + '.xlsx'
        # win
        # self.output_path = 'output_files\\' + self.output_name + '.xlsx'

    # 入力ファイルを「検索したいワードの配列」に変換
    def input_file(self, file_name):
        df_input = pd.read_excel(file_name)
        self.search_words = df_input.iloc[:, 0].values

    # APIを使用してtweetを取得し任意のデータを出力ファイル用の２次元配列に追加
    def search(self):
        for search_word in self.search_words:
            self.params['q'] = search_word
            res = self.twitter_api.get(self.url, params=self.params)

            if res.status_code == 200:
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

# メールの作成
def create_mail(from_addr, to_addr, subject, msg):
    body_msg = MIMEText(msg)
    body_msg['Subject'] = subject
    body_msg['From'] = from_addr
    body_msg['To'] = to_addr
    body_msg['Date'] = formatdate()
    return body_msg

# メールを送信
def send_mail(from_addr, to_addr, body_msg):
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.ehlo()
    smtpobj.starttls()
    smtpobj.ehlo()
    smtpobj.login(from_addr, '120sa91F')
    smtpobj.sendmail(from_addr, to_addr, body_msg.as_string())
    smtpobj.close()

def save():
    get_tweet = GetTweet(count=180)

    get_tweet.input_file(file_name='input.xlsx')
    get_tweet.search()
    get_tweet.output_file()
    print('Has been printed by ' + get_tweet.output_name)

def main():
    try:
        save()
        # saveを一時間ごとに実行
        schedule.every(1).hours.do(save)
        while True:
            schedule.run_pending()
            time.sleep(1)
    except Exception as e:
        print(e)

        JST = timezone(timedelta(hours=+9), 'JST')
        stop_time = datetime.now(JST)
        stop_time_str = stop_time.strftime('%Y-%m-%d_%H:%M:%S_JST')

        from_addr = 'f19aa021@gmail.com'
        # to_addr = 'f19aa021@gmail.com,f19aa021@chuo.ac.jp'
        to_addr = 'f19aa021@gmail.com'
        subject = 'ツイート自動収集_システム停止通知'
        with open('./mail_content.txt') as f:
            msg = f.read().format(stop_time_str, e)
            body_msg = create_mail(from_addr, to_addr, subject, msg)
        send_mail(from_addr, to_addr.split(','), body_msg)

if __name__ == '__main__':
    main()