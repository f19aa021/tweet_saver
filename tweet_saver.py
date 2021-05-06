import json
import time
from datetime import datetime, timedelta, timezone
from requests_oauthlib import OAuth1Session
import pandas as pd
import openpyxl
import schedule

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

        CK = 'xxxxxxxxxxxxxxxxxxxxx'
        CS = 'xxxxxxxxxxxxxxxxxxxxx'
        AT = 'xxxxxxxxxxxxxxxxxxxxx'
        AS = 'xxxxxxxxxxxxxxxxxxxxx'

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
        # self.output_path = 'output_files/' + output_name + '.xlsx'
        # win
        self.output_path = 'output_files\\' + self.output_name + '.xlsx'

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

def save():
    get_tweet = GetTweet(count=180)

    get_tweet.input_file(file_name='input.xlsx')
    get_tweet.search()
    get_tweet.output_file()
    print('Has been printed by ' + get_tweet.output_name)

def main():
    save()
    # 5分ごとに「タスク実行」を出力
    schedule.every(1).hours.do(save)

    # タスク監視ループ
    while True:
        # 当該時間にタスクがあれば実行
        schedule.run_pending()
        # 1秒スリープ
        time.sleep(1)

if __name__ == '__main__':
    main()