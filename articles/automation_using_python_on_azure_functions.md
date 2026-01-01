---
title: "（執筆中）Azure Functions + Python(openpyxl)で毎月のExcel集計を自動化した"
emoji: "📊"
type: "tech"
topics: ["azurefunctions", "python", "openpyxl", "excel"]
published: true
---

# 何やったの

業務で月に1度ユーザの利用集計をExcelにまとめており、その作業に約3時間かかっていたのを0にしました。

Pythonファイルに3時間分の苦悩(後述)を900行程で詰め込み、それをAzure FunctionsのTimer Trigger機能で自動的に月一実行し、出力結果であるExcelファイルをAzure Blob Storageの指定したコンテナに格納します。

今後は月に1度そのコンテナを見に行って、作成されているExcel(実体はBLOB)ファイルをダウンロードするだけで集計表が手に入る状態になりました。最高。

GitHubにPythonファイルやrequirements.txtなどをまとめたプロジェクトを上げています。

リンク張る

# 背景

集計にかかっていた3時間の作業の質も良くありませんでした。

- 機械的な作業の繰り返しで退屈
  - 自分の成長にも繋がらない
- 手動なので人的ミスが起こり放題

具体的には以下の流れで行っていました。

1. 社内のGitLabプロジェクト(リポジトリ)から、実行したいSQLファイル(複数)をコピー
    - SQLを複数ファイルに分けている理由は、それらのSQLの中でも接続先したいサーバが異なっているものがあるため、それぞれコネクションを別にしてAzure SQL Databaseと接続する必要があり、それらを分かりやすく分離するため
    - 本記事と同じAzure SQL DatabaseではなくSQL Server(オンプレ)を使用する場合は、今回使用した`pandas.read_sql()`メソッドのエラー回避[^1]のためにも分離する必要はありそう

2. [SSMS](https://learn.microsoft.com/ja-jp/ssms/install/install)(SQL Server Management Studio)に貼り付け、接続先であるAzure SQL Databaseに対して実行

3. 実行結果をコピーして集計Excelファイルに一覧表シートとして貼り付け

4. 2と3をSQLファイルの数だけ行う
    - 接続サーバ/DBを変える場合は別のコネクション(≒別のSQLファイル)を確立する必要がある

5. 一覧表シートを見るとNULLになっている箇所がある
    - これはあまりよろしくないですが、DBの仕様上そうなっていました
    - 例えば以下のような状況です

| 企業コード | 企業名 | 社員コード | 社員名 |
| :--- | :--- | :--- | :--- |
| 1　|　株式会社A　	 | A | 田中太郎　|
| 2　|　**NULL**    | B | 加藤花子　|
| 3　|　株式会社C　	 | C | **NULL** |
| 4　|　株式会社D　	 | D | 吉田勝 |

6. NULLを埋めるために別DBでクエリ実行し、エクセルの別シートに一旦貼り付けておく
    - 例えば以下のようなクエリになります

```sql:get_company_name.sql
USE [Correct-Company-Info-DB];

SELECT
  company_code AS [企業コード],
  company_name AS [企業名]
FROM [Company-Table];
```

```sql:get_employee_name.sql
USE [Correct-Employee-Info-DB];

SELECT
  employee_code AS [社員コード],
  employee_name AS [社員名]
FROM [Employee-Table];
```

7. それによってできた「NULLでない、正しい企業名/社員名が入ったエクセルシート」を、「NULLが混在している、データをそのまま貼り付けた状態のエクセルシート」から参照し、NULLを置き換える
    - Excelの[`VLOOKUP`](https://support.microsoft.com/ja-jp/office/vlookup-%E9%96%A2%E6%95%B0-0bbc8083-26fe-4963-8ab8-93a18ad188a1)関数を使用
    - 余談ですが、このVLOOKUPの第四引数を省略してあいまい検索がかかってしまい、意図しない結果が入っていたことに1カ月後気づくという初歩的な[失態](https://x.com/clumsy_ug/status/2003330818111156456?s=46)を犯しました（泣）

8. 以上で完成した各一覧表シートを基に、列ごとの合計値や平均値をとった結果を記載するサマリシートを複数作成
    - 厳密には、すでに作成されている各サマリシートに1列追加して、その月の分として新しい値を入力していく
    - Excelの`SUM`, `AVERAGE`関数を使用

9.  各サマリシートを参照している各グラフシートの、データ範囲を1列分拡張する
    - ひと月ごとに1列サマリシートの列が追加されていくため

10. 完成

[^1]: [pandas](https://pandas.pydata.org/)の[`read_sql()`](https://pandas.pydata.org/docs/reference/api/pandas.read_sql.html)メソッドは、第一引数のSQLテキストを読み取り、実行までしてくれる便利なものですが、読み取れるSQLステートメントの数が最大1つまでという制約があります。例えば以下のコードはエラーになります。`USE`, `SELECT`という2つのステートメントがあるからですね。
    ```sql
    USE [Sample-Database];
    SELECT * FROM [Sample-Table];
    ```

---

これを部分的にでも可能な限り自動化したいと思い少しずつ試してみたら、最終的に完全自動化をすることができ、3時間が0になった(やったね)ので、その方法を説明します。

# プログラムの内容


## Pythonファイル
作ったPythonファイルはこちらです。

リンク張る。

## メイン関数とTimer Trigger

Azure Functions上で実行するのはmain_process関数としています。

```python:function_app.py
def main_process():
    ...
```

これをAzure Functionsの機能として提供されているTimer Triggerというものの最新である[バージョン2(v2)の記法](https://learn.microsoft.com/ja-jp/azure/azure-functions/functions-bindings-timer?tabs=python-v2%2Cisolated-process%2Cnodejs-v4&pivots=programming-language-python#example)を使って、自動で定期実行します。今回の例では毎月20日の午前4時にしています。

（Timer Triggerという機能がAzure Functionsにあるということは知っていましたが、GUI上でポチポチ設定するのではなく決まった書式でコードに落とし込むことで利用できる機能であるということを初めて知りました）

```python:function_app.py
# schedule: "秒 分 時 日 月 曜日"
# Asia/Tokyoで毎月20日の4:00に実行したい -> UTCだと9時間前なので19日の19:00になる
@app.schedule(
        schedule="0 0 19 19 * *",
        # schedule="0 0 */2 * * *",  # デバッグ用: 2時間おきに実行
        # schedule="0 0 * * * *",  # デバッグ用: 1時間おきに実行
        arg_name="myTimer",
        run_on_startup=False,
        use_monitor=False
) 
def monthly_processing(myTimer: func.TimerRequest) -> None:
    logging.info('Python timer trigger function started.')
    
    try:
        main_process()
        logging.info('処理が正常に終了しました。')
    except Exception as e:
        logging.error(f'処理中にエラーが発生しました: {e}')
        # エラーを再送出してAzure側で失敗として記録させる
        raise
```

`schedule="0 0 19 19 * *",`は[CRON](https://wa3.i-3-i.info/word11748.html)式というもので、[こちら](https://learn.microsoft.com/ja-jp/azure/azure-functions/functions-bindings-timer?)に公式の書き方が載っています。

[UTC](https://wa3.i-3-i.info/word11831.html)として実行されてしまうので、[JST](https://wa3.i-3-i.info/word18620.html)として実行したい時間から逆算して計算する必要があります。
https://www.jisakeisan.com/?t1=utc&t2=jst

また、後述しますがPythonファイル内でも`datetime.now()`を使っている箇所があり、これだとUTC時間で認識されてしまうので、好きなタイムゾーンを指定できる[ZoneInfo](https://docs.python.org/ja/3/library/zoneinfo.html)というPythonライブラリを`datetime.now()`の引数で使用することで正確に`Asia/Tokyo`のタイムゾーンで認識できるようにしました。

## 環境変数の取得

ではmain_process内の処理を追っていきます。

まず、Azureの環境変数を取得しつつ存在しない場合はエラーをメッセージと共にraiseするという関数を`get_env_or_raise`という名前で作成しています。

それを利用しながら、接続するサーバー/DBが計3種類あるのでそれぞれ別で接続情報を取得しています。

```python:function_app.py
import os

def get_env_or_raise(key: str) -> str:
    """Azureの環境変数から指定されたキーの値を返す。存在しなかったら例外を投げる。"""

    value = os.getenv(key)
    if not value:
        raise EnvironmentError(f"環境変数 {key} が設定されていません。Azure Functions の Application Settings に追加してください。")
    return value

def main_process():
    # ---環境変数を取得---

    # Azure Blob Storageのコンテナ内にある、各SQLファイルのURL
    query_url_per_employee_1 = get_env_or_raise('QUERY_URL_PER_EMPLOYEE_1')
    query_url_per_company_1 = get_env_or_raise('QUERY_URL_PER_COMPANY_1')
    query_url_per_employee_2 = get_env_or_raise('QUERY_URL_PER_EMPLOYEE_2')
    query_url_per_company_2 = get_env_or_raise('QUERY_URL_PER_COMPANY_2')
    query_url_per_employee_3 = get_env_or_raise('QUERY_URL_PER_EMPLOYEE_3')
    query_url_per_company_3 = get_env_or_raise('QUERY_URL_PER_COMPANY_3')

    # NULLになる企業名/社員名を置き換えるために使用される正しいデータに関連するものは"get"や"getname"をつけている
    query_url_get_employeename_1 = get_env_or_raise('QUERY_URL_GET_EMPLOYEENAME_1')
    query_url_get_companyname_1 = get_env_or_raise('QUERY_URL_GET_COMPANYNAME_1')
    query_url_get_employeename_2 = get_env_or_raise('QUERY_URL_GET_EMPLOYEENAME_2')
    query_url_get_companyname_2 = get_env_or_raise('QUERY_URL_GET_COMPANYNAME_2')

    # SQLファイルのURLと、それに対応するベース名(シート名の一部として後に使用する)

    # 実データ取得用
    SQL_FILES = [
        (query_url_per_employee_1, "1(企業ごと)"),
        (query_url_per_company_1, "1(社員ごと)"),
        (query_url_per_employee_2, "2(企業ごと)"),
        (query_url_per_company_2, "2(社員ごと)"),
        (query_url_per_employee_3, "3(企業ごと)"),
        (query_url_per_company_3, "3(社員ごと)"),
    ]

    # NULLを置き換えるための正しい名前が入ったデータ取得用
    GETNAME_SQL_FILES = [
        (query_url_get_employeename_1, "1(企業名取得)"),
        (query_url_get_companyname_1, "1(社員名取得)"),
        (query_url_get_employeename_2, "2(企業名取得)"),
        (query_url_get_companyname_2, "2(社員名取得)"),
    ]

    driver = get_env_or_raise('AZURE_SQL_DRIVER')
    
    # 1
    server_1 = get_env_or_raise('SERVER_1')
    database_1 = get_env_or_raise('DATABASE_1')
    username_1 = get_env_or_raise('USERNAME_1')
    password_1 = get_env_or_raise('PASSWORD_1')

    conn_str_1 = (
        f'DRIVER={driver};'
        f'SERVER={server_1};'
        f'DATABASE={database_1};'
        f'UID={username_1};'
        f'PWD={password_1};'
        'Encrypt=yes;'
        'TrustServerCertificate=no;'
        'Connection Timeout=120;'  # Azure SQL Databaseへの接続確立の制限時間
    )

    # 2
    server_2 = get_env_or_raise('SERVER_2')
    database_2 = get_env_or_raise('DATABASE_2')
    username_2 = get_env_or_raise('USERNAME_2')
    password_2 = get_env_or_raise('PASSWORD_2')

    conn_str_2 = (
        f'DRIVER={driver};'
        f'SERVER={server_2};'
        f'DATABASE={database_2};'
        f'UID={username_2};'
        f'PWD={password_2};'
        'Encrypt=yes;'
        'TrustServerCertificate=no;'
        'Connection Timeout=120;'
    )

    # 3
    server_3 = get_env_or_raise('SERVER_3')
    database_3 = get_env_or_raise('DATABASE_3')
    username_3 = get_env_or_raise('USERNAME_3')
    password_3 = get_env_or_raise('PASSWORD_3')

    box_conn_str = (
        f'DRIVER={driver};'
        f'SERVER={server_3};'
        f'DATABASE={database_3};'
        f'UID={username_3};'
        f'PWD={password_3};'
        'Encrypt=yes;'
        'TrustServerCertificate=no;'
        'Connection Timeout=120;'
    )

    # 1で、企業名/社員名がNULLだった場合に利用する接続情報
    getname_server_1 = get_env_or_raise('GETNAME_SERVER_1')
    getname_database_1 = get_env_or_raise('GETNAME_DATABASE_1')
    getname_username_1 = get_env_or_raise('GETNAME_USERNAME_1')
    getname_password_1 = get_env_or_raise('GETNAME_PASSWORD_1')

    getname_conn_str_1 = (
        f'DRIVER={driver};'
        f'SERVER={getname_server_1};'
        f'DATABASE={getname_database_1};'
        f'UID={getname_username_1};'
        f'PWD={getname_password_1};'
        'Encrypt=yes;'
        'TrustServerCertificate=no;'
        'Connection Timeout=120;'
    )

    # 2で、企業名/社員名がNULLだった場合に利用する接続情報
    getname_server_2 = get_env_or_raise('GETNAME_SERVER_2')
    getname_database_2 = get_env_or_raise('GETNAME_DATABASE_2')
    getname_username_2 = get_env_or_raise('GETNAME_USERNAME_2')
    getname_password_2 = get_env_or_raise('GETNAME_PASSWORD_2')

    getname_conn_str_2 = (
        f'DRIVER={driver};'
        f'SERVER={getname_server_2};'
        f'DATABASE={getname_database_2};'
        f'UID={getname_username_2};'
        f'PWD={getname_password_2};'
        'Encrypt=yes;'
        'TrustServerCertificate=no;'
        'Connection Timeout=120;'
    )

    # 3ではNULLにならないため省略
    
    ...
```


# 下書きメモ

- flex consumptionだとremote buildという便利な機能があり、手元でpythonをビルドする必要がなく、またrequirements.txtさえ用意しておけばリモートでそれを自動でインストールしてくれるっぽく、凄く楽にデプロイ成功して良かった。

- 最初matplotlibで図を作ろうとしていたが、既存のexcelの図をopenpyxlで取得していじるところまでやっちゃえば良いことに気づいて楽にグラフを拡張できてよかった。スタイル引き継げるのが良い

- 最初はローカルでこういう風に実行して試してた。毎回sql走って実行を5~10分待つのが面倒だったが、本番に近い状態で正常に実行されるか常に確認しておきたかったので我慢した

- 最後の方は急いでいてリファクタしてないのだが、ご愛嬌

- 余談だがexcelなどのmsアプリのファイルは内部でxmlになっていること、だからこそ .xlsx / .docx / .pptx のxはxmlのxとして使われていること、を知った。で、グラフのタイトルが消えてしまうときとかにxmlの中身を見に行ってxmlが存在しないからxmlとは違って内部キャッシュを恐らくexcelはもっていてそこが消えてしまったんだろうとかそういうアタリをつけられて、楽しかった。実際に1文字消す、戻す、とやってxmlにそれが認識されていることを確認し、再度実行したらちゃんと直ったので、良かった。
 
 - 集計表がおかしいとか、ある時点での集計表をもう一度見たいとか、集計表を見る営業サイドだけでなく開発サイド(主に私)も過去の集計表が見たいということで、月ごとの集計表はファイル名の末尾に`_old`をつけて一応スナップショットとして保存しておくことにしている

- 単語のリンクをハイパーリンクでのせるべきところはのせとく
