---
title: "Azure Functions + Python(openpyxl)で毎月のExcel集計を自動化した"
emoji: "📊"
type: "tech"
topics: ["azurefunctions", "python", "openpyxl", "excel"]
published: true
---

# 何やったの

業務で月に1度ユーザの利用集計をExcelにまとめており、その作業に約3時間かかっていたのを0にしました。

Pythonファイルに3時間分の作業内容(後述)を800行弱で詰め込み、それをAzure FunctionsのTimer Trigger機能で自動的に月一実行し、出力結果であるExcelファイルをAzure Blob Storageの指定したコンテナに格納します。

今後は月に1度そのコンテナを見に行って、作成されているExcel(実体はBLOB)ファイルをダウンロードするだけで集計表が手に入る状態になりました。やったね。

GitHubにPythonファイルやrequirements.txtなどをまとめたリポジトリを上げています。

https://github.com/clumsy-ug/excel-operations-and-azurefunc-sample/tree/main

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
    - 余談ですが、このVLOOKUPの第四引数を省略してあいまい検索がかかってしまい、意図しない結果が入っていたことに1カ月後の集計時に気づくという[失態](https://x.com/clumsy_ug/status/2003330818111156456?s=46)を犯しました（泣）

8. 以上で完成した各一覧表シートを基に、列ごとの合計値や平均値をとった結果を記載するサマリシートを複数作成
    - 厳密には、すでに作成されている各サマリシートに1列追加して、その月の分として新しい値を入力していく
    - Excelの`SUM`, `AVERAGE`関数を使用

9.  各サマリシートを参照している各グラフシートの、データ範囲を1列分拡張する
    - 毎月サマリシートに1列追加されていくため

10. 完成

[^1]: [pandas](https://pandas.pydata.org/)の[`read_sql()`](https://pandas.pydata.org/docs/reference/api/pandas.read_sql.html)メソッドは第一引数のSQLテキストを読み取り実行までしてくれる便利なものですが、読み取れるSQLステートメントの数が最大1つまでという制約があります。例えば以下のコードは`USE`, `SELECT`という2つのステートメントがあるためエラーになります。
    ```sql
    USE [Sample-Database];
    SELECT * FROM [Sample-Table];
    ```

---

これを部分的にでも可能な限り自動化したいと思い少しずつ試してみたら、最終的に完全自動化をすることができ3時間が0になったので、その方法を説明します。

# プログラムの解説

## Pythonファイル
Pythonファイルの完成版はこちらです。

https://github.com/clumsy-ug/excel-operations-and-azurefunc-sample/blob/main/function_app.py

サンプルとして抽象化しているため実際の値や機密情報に関わる部分はマスクしています。

## メイン関数とTimer Trigger

Azure Functions上で実行するのはmain_process関数としています。

```python:function_app.py
def main_process():
    ...
```

これをAzure Functionsで提供されているTimer Triggerという機能の最新である[バージョン2(v2)の記法](https://learn.microsoft.com/ja-jp/azure/azure-functions/functions-bindings-timer?tabs=python-v2%2Cisolated-process%2Cnodejs-v4&pivots=programming-language-python#example)を使って、自動で定期実行します。今回の例では毎月20日の午前4時にしています。

（Timer Triggerという機能がAzure Functionsにあるということは知っていましたが、GUI上でポチポチ設定するのではなく決まった書式でコードに落とし込むことで利用できる機能であるということを初めて知りました）

公式の例に則り、以下のように書くことで定期実行されるようにします。

```python:function_app.py
import azure.functions as func

# アプリケーションの初期化
app = func.FunctionApp()

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
f
ちなみにローカルで動かしたいときは以下のように書いて実行します。

```python:function_app.py
if __name__ == '__main__':
    try:
        main_process()
        logging.info('処理がすべて終了しました')
    except Exception as e:
        logging.info(f"main_process実行中にエラー発生: {e}")
        raise
```

---

なお、今回使用するライブラリは以下です。

```txt:requirements.txt
requests==2.32.5
pyodbc==5.3.0
pandas==2.3.3
azure-storage-blob==12.27.1
openpyxl==3.1.5
azure-functions==1.24.0
```

同じバージョンで一括インストールするには以下を実行します。
```bash
pip install requests==2.32.5 pyodbc==5.3.0 pandas==2.3.3 azure-storage-blob==12.27.1 openpyxl==3.1.5 azure-functions==1.24.0
```

## 1. 環境変数の取得

ではmain_process内の処理を1つずつ追っていきましょう。

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

    ...

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

    conn_str_3 = (
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

今回1, 2, 3という数字がコード内に出てきますが、これはあるサービスのシステムを3種類に分けていることでサーバーやDBが異なっており、それが3種類あるという意味になります。（1と2はサーバーが違う、1と3はサーバーが同じだがDBが違う）

実際はサービス名が入りますが、今回は便宜上1, 2, 3として分けています。


## 2. 実データを取得→DFに変換

```python:function_app.py
import requests
import pandas as pd
import pyodbc
import logging

def download_sql_text(sql_url: str) -> list[str]:
    """指定URLからSQLテキストをダウンロードして返す。"""

    # timeoutはblobからsqlファイルをダウンロード完了するまでの制限時間
    response = requests.get(sql_url, timeout=600)
    response.raise_for_status()
    response.encoding = response.encoding or 'utf-8'
    return response.text

def execute_sql_to_df(
    conn_str: str,
    sql_text: str
) -> pd.DataFrame:
    """pyodbc経由でSQLを実行してpandas.DataFrameを返す。"""

    # timeoutはクエリ実行から完了までの制限時間
    # autocommitはfalseでok（SELECTのみであればsqlのcommit()不要なので）
    with pyodbc.connect(conn_str, autocommit=False, timeout=600) as conn:
        df = pd.read_sql(sql_text, conn)
        df = df.fillna('NULL')
    
    return df

def main_process():

    ...

    # ベース名とそれに対応するdataframeを保持するタプルを、0個以上保持するリスト
    basename_df_list: list[tuple[str, pd.DataFrame]] = []

    # データ取得のループ
    for sql_url, base_name in SQL_FILES:
        logging.info(f"SQL_FILESループ -> {sql_url}")
        
        try:
            sql_text = download_sql_text(sql_url)
        except Exception as e:
            logging.error(f"エラー: download_sql_text()に失敗しました: {sql_url}")
            logging.error(f"エラー内容: {e}")
            continue

        if '1' in base_name:
            conn_str = conn_str_1
        elif '2' in base_name:
            conn_str = conn_str_2
        else:
            conn_str = conn_str_3

        try:
            df = execute_sql_to_df(conn_str, sql_text)
        except Exception as e:
            logging.error(f"エラー! execute_sql_to_df()に失敗しました: {base_name}")
            logging.error(f"エラー内容: {e}")
            basename_df_list.append((f"エラー_{base_name}", None))
            continue

        basename_df_list.append((base_name, df))
        logging.info('現在のループ内の処理正常終了')

        ...
```

先ほど作成したSQL_FILESの中をループしています。

まず`download_sql_text()`でSQLそのものをstringとして受け取るためにダウンロードし、それを基に実際に`execute_sql_to_df()`でSQL実行とdataframeへの変換まで行います。

今回、SQLの実行結果を毎回Excelに直接書き込みすると恐らくオーバーヘッドが凄いので、中間状態としてpandasのDataFrameとしてデータを保持しておき、加工が終わったらそれを一気にExcelに書き込む、という形をできる限り行っています。

また実行によるログを見るために[Application Insights](https://learn.microsoft.com/ja-jp/azure/azure-monitor/app/app-insights-overview)を有効にしているのですが、そこにログとして出力させるには`logging.info()`をする必要があり、`logging`も`import`しています。

`basename_df_list`はこの後にNULLを置き換えたり、実際にExcelシートとして作り上げていく時に使用します。

`basename_df_list[0]`はシート名の一部として、`basename_df_list[1]`はそのシートのデータを表現しているDataFrameとして使用します。

なのでそのbasename_df_listを作り上げるためのループになります。

## 3. 正しい名前が入ったデータを取得→DFに変換

```python:function_app.py
def main_process():

    ...

    # 名前情報を格納する4つのdfを格納するリスト
    """
    ループの流れ
    1回目: 企業コード(1), 企業名(1)
    2回目: 社員コード(1), 社員名(1)
    3回目: 企業コード(2), 企業名(2)
    4回目: 社員コード(2), 社員名(2)
    """
    names_df_list: list[pd.DataFrame] = []

    # 名前取得のループ
    for sql_url, base_name in GETNAME_SQL_FILES:
        logging.info(f"GETNAME_SQL_FILESループ -> {sql_url}")

        # 実行したいsql文をダウンロード
        try:
            sql_text = download_sql_text(sql_url)
        except Exception as e:
            logging.error(f"エラー! download_sql_text()に失敗しました: {sql_url} : {e}")

        if '1' in base_name:
            conn_str = getname_conn_str_1
        elif '2' in base_name:
            conn_str = getname_conn_str_2

        # dfのNULLになっている名前を置き換えるための名前情報が含まれたdfを作成
        try:
            names_df = execute_sql_to_df(conn_str, sql_text)
            names_df_list.append(names_df)
        except Exception as e:
            logging.error(f"エラー! execute_sql_to_df()に失敗しました: {sql_url}")
            logging.error(f"エラー内容: {e}")

        ...
```

今度はGETNAME_SQL_FILESの中をループしますが、やっていることは同じです。

作られるのが実データのようにNULLが混在したものではなく、しっかりと名前が入ったデータです。

それをDataFrameとして保持しています。

SQL実行結果であるDataFrameを、names_df_listにどんどん追加しています。

## 4. NULLを置換

先ほどのGETNAME_SQL_FILES内のループはまだ続いています。

データは揃ったので、実際にNULLを置き換えます。
（以前までExcel内でVLOOKUP関数を使っていた箇所です）

```python:function_app.py
def main_process():
    
    ...

    for sql_url, base_name in GETNAME_SQL_FILES:
        
        ...    

        # dfのNULLをnames_dfによって置き換え
        try:
            # 1回目のループ: 1(企業名取得)
            if sql_url == GETNAME_SQL_FILES[0][0]:
                df = basename_df_list[0][1]
                replace_null(df, names_df_list, 0, '企業コード', '企業名')
        except Exception as e:
            logging.error(f"エラー! 1巡目のreplace_null()に失敗しました: {sql_url}")
            logging.error(f"エラー内容: {e}")

        try:
            # 2回目のループ: 1(企業名&社員名取得)
            if sql_url == GETNAME_SQL_FILES[1][0]:
                df = basename_df_list[1][1]
                replace_null(df, names_df_list, 0, '企業コード', '企業名')
                replace_null(df, names_df_list, 1, '社員コード', '社員名')
        except Exception as e:
            logging.error(f"エラー! 2巡目のreplace_null()に失敗しました: {sql_url}")
            logging.error(f"エラー内容: {e}")

        try:
            # 3回目のループ: 2(企業名取得)
            if sql_url == GETNAME_SQL_FILES[2][0]:
                df = basename_df_list[2][1]
                replace_null(df, names_df_list, 2, '企業コード', '企業名')
        except Exception as e:
            logging.error(f"エラー! 3巡目のreplace_null()に失敗しました: {sql_url}")
            logging.error(f"エラー内容: {e}")

        try:
            # 4回目のループ: 2(企業名&社員名取得)
            if sql_url == GETNAME_SQL_FILES[3][0]:
                df = basename_df_list[3][1]
                replace_null(df, names_df_list, 2, '企業コード', '企業名')
                replace_null(df, names_df_list, 3, '社員コード', '社員名')
        except Exception as e:
            logging.error(f"エラー! 4巡目のreplace_null()に失敗しました: {sql_url}")
            logging.error(f"エラー内容: {e}")

        logging.info('現在のループ内の処理終了')
```

main_process内の処理はほぼ`replace_null`関数を実行して、`names_df_list`内の各`names_df`を参照することで`df`のnullを置き換えているだけです。

そのため`replace_null`の中身を見てみましょう。

```python:function_app.py
def replace_null(
    df: pd.DataFrame,
    names_df_list: list[pd.DataFrame],
    names_df_order: int,
    code_column: str,
    name_column: str
) -> None:
    """1と2の企業名/社員名 のNULLになっている箇所を正確な名前に置き換える"""

    for n in range(len(df)):
        if df.at[n, name_column] == 'NULL':
            code_that_have_null_name = df.at[n, code_column]
            names_df = names_df_list[names_df_order]

            # bool_series_whether_match_code は [False, False, True, False] のような pd.Series
            bool_series_whether_match_code = names_df[code_column] == code_that_have_null_name
            del_indexes = names_df.index[bool_series_whether_match_code]
            
            if len(del_indexes) == 0:
                logging.info(f'NULLの名前を持つ{code_column} {code_that_have_null_name} が名前取得用のdfには存在しなかったため、dfの該当セルがある行は削除します。')
                # 1で削除された社員は物理削除になる仕様で正しい接続先DBにも存在しなくなっているので、集計対象外としてレコードごと削除
                df.drop(n, inplace=True)
            else:
                # DBの仕様上1つしかないのが確定しているので[0]と断定してOK
                del_index = del_indexes[0]
                correct_name = names_df.at[del_index, name_column]
                df.at[n, name_column] = correct_name
```

処理の流れは以下です。

- df(実データ)内を1行ずつループ
- dfの現在のループ行の該当の列(企業名か社員名が入っている列)の値がNULLだった場合のみ処理する
  - そうでない場合は何もしない
- dfの現在のループ行の企業コードor社員コードを、`code_that_have_null_name`として保持
- dfに対応する正しい情報を持つdfを、`names_df`として保持
- 先ほどdf内でNULLの値を持っていたコードと一致するかどうかを、names_dfの各コードで検査し、それぞれの結果がbool値になったものをpandas.Series型の`bool_series_whether_match_code`に保持
- そのSeries内でtrueになっている部分のindexを、`pandas.Index`型の`del_indexes`として保持
  - DBの仕様上、各コードは一意であるので、trueは返るとしても1つだけ
- `len(del_indexes)`が0であった場合、trueが1つもなかった、つまりdfではNULLの名前を持つコードがあるがそのコードがnames_dfには存在しないという場合、仕様上サービス上でそのデータの物理削除が行われたということになるので、そもそも集計対象外にするということで`df.drop`で行削除して終了
- `len(del_indexes)`が0でなかった場合、trueが1つ以上(実際は1つであることが確定している)あった、かつそのコードがnames_dfにちゃんと存在しているということになるので、NULLを置換

という、単純なロジックでNULLを置換しています。

恐らくもっと効率化したコードにできると思います。今回はあえて愚直で素直なアルゴリズムを自力で実装してみたくなったのでそうしました。

これで、`VLOOKUP`でやっていたNULL置換部分を自動化することに成功しました。

## 5. コンテナクライアントの初期化

```python:function_app.py
from azure.storage.blob import BlobServiceClient

def main_process():

    ...

    # Azureとの接続関連
    CONNECTION_STRING = get_env_or_raise('CONNECTION_STRING')
    CONTAINER_NAME = get_env_or_raise('CONTAINER_NAME')
    EXCEL_FILE_NAME_PREFIX = get_env_or_raise('EXCEL_FILE_NAME_PREFIX')

    # クライアントの初期化
    blob_service_client = BlobServiceClient.from_connection_string(
        CONNECTION_STRING,
        connection_timeout=600,  # 接続確立までの待機秒数
        read_timeout=600,  # データ読み込みの待機秒数
        retry_total=5  # 失敗時のリトライ回数
    )
    container_client = blob_service_client.get_container_client(CONTAINER_NAME)

    ...
```

今回、最終出力は.xlsxにしたいのですが、それを作るために素材として利用する前月分の.xlsxを読み取ったり、新しい月の分を書き込んだり、完成したものをアップロードするために、Azure Blob Storageのコンテナに接続する必要があります。

接続文字列などを環境変数から取得し、それを利用してコンテナに接続するためのクライアントを`container_client`という名前で初期化します。

## 6. ブロブクライアントの初期化

```python:function_app.py
def main_process():

    ...

    # 最新(先月20日時点)のエクセルを見つける
    # それより前の月のエクセルはファイル名に_oldをつけてアーカイブ扱いしているので、_oldがないものが最新ということになる
    blobs = list(container_client.list_blobs(name_starts_with=EXCEL_FILE_NAME_PREFIX))

    # listに格納されたエクセルの中から最新のものを見つける
    for blob in blobs:
        if '_old' not in blob.name:
            latest_excel_blob_name = blob.name
            break

    # Blobからエクセルファイルを仮想メモリにダウンロード
    # pd.read_excel()だけで済ませてしまうと、エクセルとして持ってくるのではなく中身のデータしか持ってこないからデザインが消えたり色々な問題があるのでバイトで扱う。その後にピンポイントでデータを置き換えたいところでだけpandas使用していく
    # get_blob_clientはgetというよりcreateが実態に近い
    blob_client = container_client.get_blob_client(latest_excel_blob_name)

    ...
```

まず、集計月の1月前のExcel、つまり現時点では最新のExcelのファイル名を`latest_excel_blob_name`として取得します。

そしてそのファイル名のblobを操作する窓口であるクライアントを、`blob_client`として作成します。

コメントにも書いた通りですが、以下が注意点です。

- `get_blob_client()`ではなく`pd.read_excel()`だけで済ませてしまうと、エクセルとして持ってくるのではなく中身のデータしか持ってこないためデザインが消えたり色々な問題があるのでバイトで扱う
  - その後にピンポイントでデータを置き換えたいところでだけpandasを使用
- get_blob_clientはgetというよりcreateが実態に近い

## 7. Workbook作成

```python:function_app.py
from openpyxl import load_workbook
import io

def main_process():

    ...

    # メモリ上に空の仮想ファイルを作成
    # コンピュータのディスク(HDD/SSD)ではなくメモリ(RAM)に作成される
    download_stream = io.BytesIO()

    # クラウドからデータをダウンロードし、仮想ファイルに流し込む(書き込む)
    blob_client.download_blob().readinto(download_stream)

    # 読み取り位置(カーソル)が最後になっているので、先頭(0バイト目)に戻す
    # これをやらないと pd.read_excel(download_stream) などやってもデータが空と判断されてエラーになる
    download_stream.seek(0)

    # excelの基になるbookを作成
    book = load_workbook(download_stream)
    output_stream = io.BytesIO()

    ...
```

以下を行っています。

- 空の仮想ファイルを`BytesIO`型(バイト列 / バイナリ)の`download_stream`として作成
- 先ほど作成したblob_clientを使って実際にblobの中身(データ)を取得し、それをdownload_streamに流し込む
- `download_stream`の読み取り位置(カーソル)を先頭に戻す
  - これをやらないと`download_stream`の末尾つまりデータが無いところから読み込もうとしてしまいエラーになる
- [openpyxl](https://pypi.org/project/openpyxl/)の`load_workbook`関数の引数にバイト列の`download_stream`を入れ実行することで、返り値として`openpyxl.Workbook`型の`book`を取得
    - [こちら](https://note.nkmk.me/python-openpyxl-usage/)の説明にあるような`openpyxl.load_workbook('data/src/sample.xlsx')`という形で`load_workbook`の引数にはExcelファイルの物理パスを入れても良いですが、ファイルライクなオブジェクト(`BytesIO`などのバイナリストリーム)も受け付けるように作られているのでバイナリを入れても問題ないです
- 最後に、最終出力用のバイナリストリームであるoutput_streamも初期化しておく

blobをダウンロードしてそれをわざわざバイナリ(BytesIO)という中間表現に一旦落とし込んで、今度はそれをopenpyxlの`Workbook`という形に変換する、という手間をしています。
まぁBLOBもBinary Large Objectの略なのでBinary、つまり同じじゃんという感じではあるんですが、`readinto()`を使用して`BytesIO`というpythonの型に明示的に変換しているので、AzureのblobとpythonのBytesIOという2つのバイナリはやはり厳密には違う構造だと言えそうです。

一見、お目当てのblobを`download_blob`で取得して終わりに出来たら楽そうですが、残念ながらできません。

`BlobClient.download_blob()`の返り値は`StorageStreamDownloader[bytes]`というもので、これはExcelのデータそのものでも、バイナリでもなく、ダウンロードを管理する専用のオブジェクトでしかないからです。

```python:_blob_client.py
class BlobClient(StorageAccountHostsMixin, StorageEncryptionMixin):
    ...
    @overload
    def download_blob(
        self, offset: Optional[int] = None,
        length: Optional[int] = None,
        *,
        encoding: None = None,
        **kwargs: Any
    ) -> StorageStreamDownloader[bytes]:
        ...
```

なので後からお目当てのblob(excel)ファイルを操作したい場合、`StorageStreamDownloader`として状態を保持していてもしょうがない(操作できない)ので、`openpyxl.Workbook`という状態(型)にしておきたい訳ですが、そのためには今回のケースだと一旦バイナリという状態を挟む必要があるので、まず`StorageStreamDownloader`をバイナリストリームにし、それを更に`Workbook`にしている、という流れになります。

## 8. 一覧表作成

今回操作対象のExcelファイルのシート情報は以下になっています。

- 一覧表6シート
  - データを貼り付けただけ
- サマリ6シート
  - 一覧表を基に、列単位で`SUM`や`AVERAGE`を取る
- グラフ6シート
  - サマリを基に、月ごとの推移を折れ線グラフで描画
- 計18シート

このうちの一覧表の作成に入ります。
実際には、既存のExcelファイルの一覧表シートのヘッダー以外の行を全て削除し、そこに新しい値(SQL実行結果)を張り付けます。

```python:function_app.py
from zoneinfo import ZoneInfo
from openpyxl.utils.dataframe import dataframe_to_rows

def main_process():

    ...

    # ---一覧表(6シート)の作成---
    # 今の年月(新規作成するシート名に使用)
    now_year_dot_month = datetime.now(ZoneInfo('Asia/Tokyo')).strftime('%Y.%m')
    now = datetime.now(ZoneInfo('Asia/Tokyo'))

    # 1か月前の年月(削除するシート名の判別に使用)
    prev_date = pd.Timestamp(datetime.now(ZoneInfo('Asia/Tokyo'))) - pd.DateOffset(months=1)
    prev_year_dot_month = prev_date.strftime('%Y.%m')

    for basename, df in basename_df_list:
        new_sheetname = f"【{now_year_dot_month}】{basename}"
        old_sheetname = f"【{prev_year_dot_month}】{basename}"

        ws = book[old_sheetname]

        # シート名を変更(年月日を今月のものに変える)
        ws.title = new_sheetname

        # 既存のシートのデータ(ヘッダーである3行目までを除く。そこまではそのままで良い)を削除(4行目から、現在の最大行数分だけ削除)
        ws.delete_rows(4, amount=ws.max_row)

        # データの書き込み
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)

        match basename:
            case '1(企業ごと)':
                # A1セルの値の年月日を現在のものに変える
                ws['A1'].value = f'1(企業ごと)利用集計({now.year}年{now.month}月{now.day}日時点累計数)'
                # 書式設定(左揃え、右揃え、桁区切り)の適用
                logging.info(f'{basename}のapply_column_style()を実行中...')
                apply_column_style(ws, ['企業コード', '企業名'])
            case '1(社員ごと)':
                ws['A1'].value = f'1(社員ごと)利用集計({now.year}年{now.month}月{now.day}日時点累計数)'
                logging.info(f'{basename}のapply_column_style()を実行中...')
                apply_column_style(ws, ['企業コード', '企業名', '社員コード', '社員名'])
            case '2(企業ごと)':
                ws['A1'].value = f'2(企業ごと)利用集計({now.year}年{now.month}月{now.day}日時点累計数)'
                logging.info(f'{basename}のapply_column_style()を実行中...')
                apply_column_style(ws, ['企業コード', '企業名'])
            case '2(社員ごと)':
                ws['A1'].value = f'2(社員ごと)利用集計({now.year}年{now.month}月{now.day}日時点累計数)'
                logging.info(f'{basename}のapply_column_style()を実行中...')
                apply_column_style(ws, ['企業コード', '企業名', '社員コード', '社員名'])
            case '3(企業ごと)':
                # 2列しかなくスタイル適用処理が必要ない
                ws['A1'].value = f'3(企業ごと)利用集計({now.year}年{now.month}月{now.day}日時点累計数)'
            case '3(社員ごと)':
                ws['A1'].value = f'3(社員ごと)利用集計({now.year}年{now.month}月{now.day}日時点累計数)'
                logging.info(f'{basename}のapply_column_style()を実行中...')
                # 企業名は取得する必要なし（集計の都合上）
                apply_column_style(ws, ['企業コード', '企業コード2', '社員コード', '社員名'])
```

以下のようなことをやっています。

### 時刻関連の変数を作成
- Azure(クラウド)上ではdatetime.nowはUTCとして実行されてしまうため、回避策としての`ZoneInfo`を使用
- シート名判別に使用する`prev_year_dot_month`と、シート名作成(変更)に使用する`now_year_dot_month`を作成
  - `prev_year_dot_month`を算出するための`prev_date`も作成
    - `pandas.Timestamp`と`pandas.DateOffset`を使用して実現
- 各一覧表シートのA1セルに使用するnowを作成
- 

### `basename_df_list`内をループ
- basename_df_listは先ほどの [# 2. 実データを取得→DFに変換](https://zenn.dev/yg_kita/articles/automation_using_python_on_azure_functions#2.-%E5%AE%9F%E3%83%87%E3%83%BC%E3%82%BF%E3%82%92%E5%8F%96%E5%BE%97%E2%86%92df%E3%81%AB%E5%A4%89%E6%8F%9B) で作成したリスト
- 各一覧表シートのシート名を変更
  - 例: `【2025.11】1(企業ごと)` → `【2025.12】1(企業ごと)`
  - Workbook[シート名]の形でシート(ワークシート / `ws`)が取得できます。
    - 型は`_WorksheetOrChartsheetLike`というものですがほぼ`Worksheet`と同じだと思います
- ヘッダーである3行目までを除き、それ以降の4行目から最後の行までを削除
  - それによってヘッダーはスタイルなどもそのまま残すことができる
  - ただ4行目以降の値に関してはセルのスタイルが失われるので、後述の`apply_column_style`関数でスタイルを適用する
- dfを行単位でループし、ワークシートの4行目以降にその行データを追加(挿入)
  - なぜちゃんと4行目から追加されるかというと、`openpyxl`の`Worksheet.append()`の挙動は「現在データが存在している最終行の次の行から追加する」という[仕様](https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.worksheet.worksheet.html#openpyxl.worksheet.worksheet.Worksheet.append)だから
    - > Appends a group of values at the bottom of the current sheet.
- あとはシートごと(basenameごと)に、左揃えにしたいカラム名が違う、というようにスタイルの適用のさせ方が微妙に違うので分岐させています
  - が、やっていることはほぼ同じで、A1セルを最新の日付で更新したあと`apply_column_style`を実行しています
    - この関数について以下詳解します

---
`apply_column_style`の中身は以下です。

```python:function_app.py
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment

def apply_column_style(
    ws: Worksheet,
    left_align_cols: list[str]
) -> None:
    """
    【追加】指定されたシートのデータ行(4行目以降)に対して書式設定を行う。
    3行目をヘッダーとして列名を判定する。
    指定列は左揃え、それ以外は右揃え + 3桁カンマ区切りにする。
    
    Args:
        ws: 対象のWorksheet
        left_align_cols: 左揃えにする列名のリスト
    """

    # 列インデックスと列名のマッピングを作成 (1始まり)
    left_col_indices = set()
    right_col_indices = set()

    # ヘッダー行のセルを読み込む
    for cell in ws[3]:
        # なぜか最初のシートだけNoneが4個認識されてしまうので弾く
        if not cell.value:
            break

        col_name = str(cell.value)

        # 指定された列名なら左揃えリストへ
        if col_name in left_align_cols:
            left_col_indices.add(cell.column)
        else:
            right_col_indices.add(cell.column)
        
    # 行ごとにスタイル適用(列で一気にやろうとしたらなぜか効かなかったため)
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for cell in row:
            if cell.column in left_col_indices:
                cell.alignment = Alignment(horizontal='left')
            elif cell.column in right_col_indices:
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '#,##0'
```

ヘッダの中でも左揃えにしたい列と右揃えにしたい列があります。
その列名を第二引数に配列として渡しています。
第一引数は変更を加えるワークシートそのものです。

そしてそれらの列ごとに、行ループ->セルループ と二重ループし、セルごとに書式を整えています。
右揃えにするセルは3桁ごとにカンマを打つ数字として扱いたいので、その書式指定も行っています。

列単位で一気に書式設定できれば楽かつ直感的ですが、なぜか効かなかったためこのやり方にしています。

## 9. サマリ作成

```python:function_app.py
from openpyxl import Workbook  # load_workbookの返り値bookの型を表現するときにのみ使う

def main_process():

    ...

    # ---サマリ(6シート)の作成---
    logging.info('「【サマリ】1(企業)」に1列追加中...')
    add_new_column_to_summarysheet_about_number_of_company(
        basename_df_list[1][1],
        book,
        '【サマリ】1(企業)',
    )

    logging.info('「【サマリ】1(社員)」に1列追加中...')
    add_new_column_to_summarysheet_about_number_of_employee(
        basename_df_list[0][1],
        basename_df_list[1][1],
        book,
        '【サマリ】1(社員)',
    )

    logging.info('「【サマリ】2(企業)」に1列追加中...')
    add_new_column_to_summarysheet_about_number_of_company(
        basename_df_list[3][1],
        book,
        '【サマリ】2(企業)',
    )

    logging.info('「【サマリ】2(社員)」に1列追加中...')
    add_new_column_to_summarysheet_about_number_of_employee(
        basename_df_list[2][1],
        basename_df_list[3][1],
        book,
        '【サマリ】2(社員)',
    )

    logging.info('「【サマリ】3(企業)」に1列追加中...')
    add_new_column_to_summarysheet_about_number_of_company(
        basename_df_list[5][1],
        book,
        '【サマリ】3(企業)'
    )

    logging.info('「【サマリ】3(社員)」に1列追加中...')
    add_new_column_to_summarysheet_about_number_of_employee(
        None,
        basename_df_list[5][1],
        book,
        '【サマリ】3(社員)'
    )

    ...
```

`add_new_column_to_summarysheet_about_number_of_employee`, `add_new_column_to_summarysheet_about_number_of_employee`関数は既存のサマリシート(計6シート)に新しい列を追加し、そこに集計月の新しい値を入れていく処理をしています。

それぞれの関数の実装は以下です。

```python:function_app.py
from datetime import datetime
from openpyxl.utils import get_column_letter

date_of_execution = datetime.now(ZoneInfo('Asia/Tokyo')).strftime('%Y/%m/%d')

def add_new_column_to_summarysheet_about_number_of_company(
    target_df: pd.DataFrame,
    book: Workbook,
    summary_sheetname: str,
) -> None:
    """【サマリ】1,2,3(企業)を作成"""

    try:
        # 参照渡しなので、これ以降summary_sheetを変更したらbookを変更したことにもなる
        summary_sheet = book[summary_sheetname]
        
        # 最後の列の番号を取得する (例: c列までデータがあれば、max_columnは3になる)
        last_col_number = summary_sheet.max_column
        # その次の列の番号(ここに今月分の値を代入していく)
        target_col_number = last_col_number + 1

        target_letter = get_column_letter(target_col_number)

        # 列の幅は25.75で固定（実際はなぜか25.17になる）
        summary_sheet.column_dimensions[target_letter].width = 25.75

        set_value_and_copy_style(summary_sheet, 4, target_col_number, date_of_execution)

        total_number_of_employees = target_df['社員数'].sum()
        set_value_and_copy_style(summary_sheet, 5, target_col_number, total_number_of_employees)

        if '3' in summary_sheetname:
            xxx = target_df['xxx'].sum()
            # ...
        else:
            yyy = target_df['yyy'].sum()
            # ...

        # ...

        # サマリシートの各テーブル範囲を1列増やす
        expand_table_range(summary_sheet)
    except Exception as e:
        logging.error(f"「{summary_sheetname}」シートへの書き込みもしくは値の計算に失敗しました: {e}")

def add_new_column_to_summarysheet_about_number_of_employee(
    target_df_user: pd.DataFrame | None,  # 3では使用しないためNoneで呼び出す
    target_df_office: pd.DataFrame,
    book: Workbook,
    summary_sheetname: str,
) -> None:
    """【サマリ】1,2,3(社員)を作成"""

    try:
        summary_sheet: Worksheet = book[summary_sheetname]

        last_col_number = summary_sheet.max_column
        target_col_number = last_col_number + 1

        target_letter = get_column_letter(target_col_number)
        summary_sheet.column_dimensions[target_letter].width = 25.75

        set_value_and_copy_style(summary_sheet, 4, target_col_number, date_of_execution)

        total_number_of_assigned_tasks = target_df_office['担当業務数'].mean()
        set_value_and_copy_style(summary_sheet, 5, target_col_number, total_number_of_assigned_tasks)

        if '3' in summary_sheetname:
            xxx = target_df['xxx'].sum()
            # ...
        else:
            yyy = target_df_user['yyy'].sum()
            # ...

        # ...

        expand_table_range(summary_sheet)
    except Exception as e:
        logging.error(f"「{summary_sheetname}」シートへの書き込みもしくは値の計算に失敗しました: {e}")
```

現在値の入っている最終列の文字(`E`, `F`など)が何なのかがわかればその次の列から値を挿入していくことができるので、`get_column_letter`を使用して`target_letter`として保持しています。

そしてdfの該当列の`sum()`や`mean()`をして、その結果を挿入するために`set_value_and_copy_style`をしています。

また、サマリはいくつかのテーブルとして作成しているので、最後に`expand_table_range`で1列分テーブル範囲を拡張しています。

まず`set_value_and_copy_style`を見てみましょう。

```python:function_app.py
from copy import copy

def set_value_and_copy_style(
    summary_sheet: Worksheet,
    row: int,
    col: int,
    value: int | float  # sum()の返り値はAnyという仕様らしいが今回はint,もしくはfloatと断言して良いはず
) -> None:
    """指定したセルに値を書き込み、すぐ左の列(col-1)のセルから書式(フォント、罫線、塗りつぶし、表示形式、配置)をコピー"""

    # セルへの値の書き込み(Setter的な使い方)
    cell = summary_sheet.cell(row=row, column=col, value=value)

    # セルの値の取得(Getter的な使い方。書き込みは行われない)
    source_cell = summary_sheet.cell(row=row, column=col - 1)
    
    if source_cell.has_style:
        cell.font = copy(source_cell.font)  # フォント
        cell.border = copy(source_cell.border)  # 白い罫線(グリッド線)
        cell.fill = copy(source_cell.fill)  # 塗りつぶし(背景色)
        cell.number_format = copy(source_cell.number_format)  # カンマ区切り
        cell.protection = copy(source_cell.protection)  # シート保護やセルのロック
        cell.alignment = copy(source_cell.alignment)  # 配置(右揃え など)
```

値を挿入した後、書式を左隣のセル(前月時点の最終列の値)からコピーしています。

次に`expand_table_range`は以下です。

```python:function_app.py
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import TableColumn

def expand_table_range(ws: Worksheet) -> None:
    """
    サマリシート内の各テーブル範囲(ref)を拡張し、不足している列定義(TableColumn)を追加する。
    TableColum追加まで行わないと「ファイルが破損しています」というエラーになる。
    """

    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)

        # 1. 範囲（ref）の更新
        new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col + 1)}{max_row}"
        table.ref = new_ref

        # 2. オートフィルタの範囲更新（設定されている場合）
        if table.autoFilter:
            table.autoFilter.ref = new_ref

        # 3. 列定義（TableColumn）の追加
        # idはユニークである必要がある
        current_id = len(table.tableColumns) + 1

        # ヘッダー行(min_row)の最終列(今回挿入した新しい列)のセルから値を取得して、列名にする。つまり最新の 年/月/日 になる
        # テーブルの列名は必須であり、かつ重複してはいけない
        # nameは日付型などではなく文字列型が必須のため str() で変換する
        header_val = ws.cell(row=min_row, column=max_col + 1).value
        str_header_val = str(header_val)
        
        # 定義を作成して追加
        new_col = TableColumn(id=current_id, name=str_header_val)
        table.tableColumns.append(new_col)
```

`ws.tables.values()`の返り値に各`table`がすべて格納されているので、その中をループしています。

`table.ref`という部分を更新したり、`table.tableColumns`を追加したり、ということをしています。

## 10. グラフ範囲拡張の準備

```python:function_app.py
def main_process():

    ...

    # ---グラフのデータ範囲を1列分拡張(6シート)---

    # 編集が終わったbookをoutput_streamに保存
    logging.info('編集が終わったbookをoutput_streamに保存中...')
    book.save(output_stream)

    # シート名(文字列)をキーとして、拡張前の列名と拡張後の列名が入ったタプルを格納する辞書を作成
    replacements = create_replacements_dict(book)

    ...
```

bookが完成したため、`output_stream`に保存しています。
また、サマリが1列増えたことに伴ってグラフのデータ範囲も1列拡張したい訳ですが、その下準備として`create_replacements_dict`を実行しています。

`create_replacements_dict`は以下です。

```python:function_app.py
def create_replacements_dict(book: Workbook) -> dict[str, tuple[str, str]]:
    """replacements辞書を作成し返却"""

    summary_sheet_names = [
        '【サマリ】1(企業)', '【サマリ】1(社員)',
        '【サマリ】2(企業)', '【サマリ】2(社員)',
        '【サマリ】3(企業)', '【サマリ】3(社員)'
    ]
    replacements = {}

    for summary_sheet_name in summary_sheet_names:
        ws = book[summary_sheet_name]

        current_max_col = ws.max_column     # 現在の最終列（拡張後の列）
        prev_max_col = current_max_col - 1  # 拡張前の列（1つ左）
        
        old_letter = get_column_letter(prev_max_col)
        new_letter = get_column_letter(current_max_col)
        
        replacements[summary_sheet_name] = (old_letter, new_letter)
        logging.info(f"グラフのデータ範囲の最終列の置換ルール登録: {summary_sheet_name}シートの{old_letter}までを{new_letter}までに拡張")
    
    return replacements
```

これにより、どのシートの何列から何列まで拡張するか、という情報だけ事前に辞書として保存しています。

## 11. グラフ範囲拡張

```python:function_app.py
def main_process():

    ...

    # 関数を呼んでストリームの中身を書き換える
    logging.info("グラフ範囲のXML直接置換を実行中...")
    output_stream = patch_xlsx_charts(output_stream, replacements)

    ...
```

`patch_xlsx_charts`実行により実際にグラフの範囲を1列分拡張しています。

この関数が今回で一番複雑で面白いです。（多分ダーティハックですが）

これまでと同様openpyxlで実現したかったのですがなぜかうまくいかなかったので、`.xlsx`の中身である`.xml`を直接操作することにしました。

というかそもそもExcelの中身がxmlである(厳密には「`.xlsx`は`.zip`であり、`.zip`の中に複数`.xml`等がある」)ことを初めて知ったので、凄く勉強になりました。
オフィス製品である`.xlsx`, `.docx`, `.pptx`の最後のxはxmlのxだったんですね。

**OOXML**(Office Open XML)という規格で標準化されているようです。
https://e-words.jp/w/Office_Open_XML.html

で、excelがxmlであったことを知らないので当然、xmlを直接いじったこともなく、どういう構造/中身になっているのかもpythonからどう操作するのかも知らなかったのですが、それが理解できたのが良かったです。

ということで解説していきます。

```python:function_app.py
import zipfile
import re

def patch_xlsx_charts(
    input_stream: io.BytesIO,
    replacements: list[tuple[str, str]]
) -> io.BytesIO:
    """
    保存した後のxlsxファイル(zip)のバイナリ(ストリーム)を受け取り、内部のチャートXMLを直接書き換えて
    グラフのデータ範囲を拡張(1列分増やす)した新しいストリームを返す。
    サマリのテーブル拡張のようにopenpyxlでやろうとしたがグラフを認識できなかったためこの方法で行う。
    """
    
    # 読み取り位置を先頭に戻す
    input_stream.seek(0)
    
    # 出力用の新しいストリーム
    output_stream = io.BytesIO()
    
    # zipとして開き、ファイルをコピーしながら必要なら書き換え
    # 型1. ZipFile: Zipファイル全体
    # 型2. ZipInfo: ZipFileの中に入っている個々のファイルのメタ情報(filenameなど)
    with zipfile.ZipFile(input_stream, 'r') as zin:
        with zipfile.ZipFile(output_stream, 'w') as zout:
            # itemはzipfileの中の各xml
            for item in zin.infolist():
                data = zin.read(item.filename)

                # チャート定義ファイルの場合のみ置換処理を行う
                if item.filename.startswith('xl/charts/chart') and item.filename.endswith('.xml'):
                    # XMLはバイト列なので文字列にデコード
                    xml_str = data.decode('utf-8')

                    for summary_sheet_name in replacements:
                        # 見つけた.xmlがどのサマリシートを参照しているのか、特定するまでループ
                        # 特定したらそのサマリの拡張前・拡張後の列名のペアを取り出し、それを.xmlに適用(1列拡張)する
                        if summary_sheet_name in xml_str:
                            old_col, new_col = replacements[summary_sheet_name]
                            # 正規表現: コロン(:) + $ + 旧列文字 + $ + 数字
                            # 例: E列をF列にする場合、 :$E$5 -> :$F$5 に置換する
                            # これにより範囲の「終了位置」だけが伸びる
                            pattern = f"(:\\$){old_col}(\\$\\d+)"
                            repl = f"\\g<1>{new_col}\\g<2>"
                            
                            # re.sub(正規表現, 正規表現にマッチした部分の置換後の文字列, 置換対象の文字列)
                            xml_str = re.sub(pattern, repl, xml_str)
                            break
                    
                    # 書き換えたデータをUTF-8バイト列に戻す
                    data = xml_str.encode('utf-8')
                
                # 新しいzipに書き込み
                zout.writestr(item, data)
    
    # ポインタを先頭に戻して返す
    output_stream.seek(0)

    return output_stream
```

まず入力と最終出力(returnするもの)は、どちらもお馴染みのバイナリストリーム(`input_stream` / `output_stream`)にしました。

ただ、その中間表現として`zip`として扱う必要があります。

xlsxの中身であるxmlをいじりたいので、xmlを含んでいるzipを認識する必要があり、zipとして読んでいます。

先程はバイナリのままではなく`Workbook`にしないと操作ができないので`Workbook`にして、操作が終わったらバイナリに戻して、ということをしていましたがそういうことを今回もする必要があります。

↓ 前者を後者にしないといけないというイメージです
```
AzureのBLOB -> Workbook
```

```
AzureのBLOB -> Binary(BytesIO) -> Workbook
```

:::message
余談ですが、こういう変換は面倒と言えば面倒ではあるもののデータ形式を変えるので当然のことだし、そもそも`zipfile.ZipFile(BytesIO)`とやるだけでバイナリをzipとして読み取れたり、`StorageStreamDownloader[bytes].readinto(BytesIO)`とやるだけでazure blobをBytesIOに流し込めたり、やはりでかいエコシステムには便利なメソッドが当然のように用意されてるのでそこまで面倒だと悲観するような話ではないかもしれないと自省しました、魔法に感謝。
:::

各xmlファイル(のメタ情報)をループで回り、パスが`xl/charts/chart`で始まり`.xml`で終わる場合はグラフを表すxmlなので、処理を行います。

ちなみにそれを確認するには、適当な.xlsxファイルの拡張子を`.zip`にrenameしてその中を覗いてみるとわかります。

グラフが`xl/charts/chart1.xml`, `chart2.xml`, `chart3.xml` ...という名前で格納される構造になっていることが確認できます。

グラフの範囲を拡張するためにxmlの中身を文字列(str)として取り出したいです。
そこでややこしいのですが、わざわざバイナリをzipにしたにも関わらず`zin.read(filename)`をしてまたバイナリに戻し、それを`bytes.decode('utf-8')`とやることでutf-8として文字列にデコードしています。

それが`xml_str`変数です。

で、先ほど下準備で作成した`replacements`を使います。

replacementsの該当シートをキーとして保存してある「古い列(`E`など)」と「拡張後の新しい列(`F`など)」を`old_col`, `new_col`として取得します。

そして、例えばEをFまで拡張したい場合は`:$E$5`という文字列を`:$F$5`という文字列に置換すれば良いことになるので、これを正規表現のグルーピングで実現しています。

（正規表現のグルーピングはpythonの`re`モジュールの標準機能です）
https://docs.python.org/ja/3.13/howto/regex.html#grouping

置換が完了したxml文字列を`str.encode('utf-8')`でバイナリに戻して、それを`ZipFile.writestr(ZipInfo, binary)`でZipFileに書き込むことで、それと同期されている最終出力の`output_stream`も変更することができます。

ループが終わったらカーソルを先頭に戻してからその`output_stream`を返却して終了です。

---
ちなみに、実は最初matplotlibで0から図を作ろうとしていましたが、途中から既存のexcelの図を操作できたらそっちの方が良いということに気づきました。スタイルなどがすべて引き継げるのが良いです。

また、グラフのタイトルがExcelファイルの見た目上は見えているのに、xmlの中身を見てみるとタイトルが存在しないことになっているという事象がありました。
ここから考えられることは、Excelファイルはxmlと見た目が完全に一致している訳ではなく、一時キャッシュのようなものを恐らく持っており、それが見た目に寄与することがあるということです。

## 12. アップロードとリネーム

これでラストです。

```python:function_app.py
def main_process():

    ...

    # 新しいファイル名でアップロード(実行時の年月日を使用)
    today_str = datetime.now(ZoneInfo('Asia/Tokyo')).strftime('%Y%m%d')
    new_excel_blob_name = f'{EXCEL_FILE_NAME_PREFIX}{today_str}.xlsx'

    logging.info('get_blob_client(new_excel_blob_name)開始...')
    new_blob_client = container_client.get_blob_client(new_excel_blob_name)

    # timeout: 処理全体のタイムアウト秒数
    # max_concurrency: 並列アップロード数（デフォルトは1。増やすと速くなるが、不安定な回線では1か2が良い）
    logging.info('upload_blob()開始...')
    new_blob_client.upload_blob(output_stream, overwrite=True, timeout=600, max_concurrency=2)

    logging.info(f'新規ファイルをアップロードしました: {new_excel_blob_name}')

    # 古いファイルの名前に _old をつける(Azure Blobにはリネームのコマンドがないため、コピーして削除する)
    old_renamed_blob_name = latest_excel_blob_name.replace('.xlsx', '_old.xlsx')

    logging.info('get_blob_client(old_renamed_blob_name)開始...')
    old_blob_client = container_client.get_blob_client(old_renamed_blob_name)

    # 先月時点のblobのコピーとして _old というsuffixをつけたblobをコピーによりコンテナ上で作成
    logging.info('start_copy_from_url()開始...')
    old_blob_client.start_copy_from_url(blob_client.url)

    # 元のファイルを削除
    logging.info('delete_blob()開始...')
    blob_client.delete_blob()

    logging.info(f'古いファイルをリネームしました: {old_renamed_blob_name}')
```

要は、完成したバイナリストリームをblobとしてuploadし、古いファイルは`_old`付きのファイル名にrename(厳密にはコピー&削除)しているだけです。

詳細としては以下の流れで処理をしています。

- 集計月の新しい日付をつけたexcelファイル名を`new_excel_blob_name`として作成
- そのファイル名のファイルの接続操作用のクライアントを`new_blob_client`として作成
- 完成している`output_stream`をクライアントから先ほどのファイル名のblobとしてアップロードする
- 「集計表がおかしい」「〇月時点の集計表をもう一度見たい」などの要望に対応するために、過去の集計表はファイル名の末尾に`_old`をつけて一応スナップショットとして保存しておくことにしている
  - しかし前月時点のファイルの末尾に`_old`というsuffixをつけたいがazure上で直接renameする機能がない
  - そのためファイル名に`_old`をつけたファイルとして前月時点のファイルを中身だけコピーし、そのあと前月時点のファイルを削除することで実質renameを実現

:::message alert
なお`start_copy_from_url`関数についてなのですが、この関数は非同期処理で実際にコピーする前にAzureのキューにコピー命令を入れるだけでそれが終わるとすぐ次の行に行ってしまう可能性があり、そうなると先にdeleteが行われてcopyしようとしたらcopy元のファイルが見つからないということになる可能性もあるか...？とふと思いました。今のところエラーにはならず正常に実行できているので変更の予定はないですが、もし変えるなら以下のようにするのが良いかもしれません。（動作は未検証です）
:::

:::details copyが終了してからdeleteする安全な実装
```python
import time
 
# 1. コピー命令を出す（非同期で開始される）
logging.info('start_copy_from_url()開始...')
old_blob_client.start_copy_from_url(blob_client.url)
 
# 2. コピーが完了するまで監視する（ポーリング処理）
start_time = time.time()
while True:
    # 宛先（新しいファイル）のプロパティを取得してステータスを確認
    props = old_blob_client.get_blob_properties()
    copy_status = props.copy.status
    if copy_status == 'success':
        # コピー成功のためループを抜ける
        logging.info('コピー完了を確認しました。')
        break
    elif copy_status == 'pending':
        # 処理中のため少し待ってから再確認
        logging.info(f'コピー処理中... (status: {copy_status})')
        time.sleep(2)  # 2秒待機
    else:
        # 失敗 (failed, aborted など)
        raise Exception(f"コピーに失敗しました。Status: {copy_status}")
 
    # 無限ループ防止のタイムアウト設定
    if time.time() - start_time > 60:  # 60秒以上かかったらエラーにする
        raise TimeoutError("コピー処理がタイムアウトしました。")
 
# 3. 確実にコピーが終わったことを確認してから、元のファイルを削除
logging.info('delete_blob()開始...')
blob_client.delete_blob()
```
:::

# デプロイ

作成したPythonをAzure Functionsで自動定期実行させるために、デプロイをします。

なお、今回はAzure Functionsのflex consumptionプラン(flex 従量課金プラン)を使用します。

## 1. ログイン

```bash
az login --tenant xxxxx
```

テナントIDを指定してログインします。
2回目以降は`az login`だけでもいけました。

`az login`だけ実行するとテナント選択画面が出てくるのでそこから選択してもできることがありますが、それができず権限エラーのようになることもあったので、最初はテナントIDまで指定するのが確実だと思います。

テナントIDは、Azure Portalに入ってMicrosoft Entra IDと入力して出てきたものを選択すれば確認できます。

## 2. Zipファイル作成

今回作成したpythonの1ファイルだけではなく、計4ファイル作成し、それらを.zipに圧縮してそのzipをデプロイする必要があります。

1. function_app.py
   1. 作成したpythonファイル。このファイル名にしてください
2. host.json
3. local.settings.json
4. requirements.txt

https://learn.microsoft.com/ja-jp/azure/azure-functions/functions-reference-python?tabs=get-started%2Casgi%2Capplication-level&pivots=python-mode-decorators#folder-structure

python以外の各ファイルの中身は最小限の構成ということで、以下にします。

```json:host.json
{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "excludedTypes": "Request"
      }
    }
  },
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle",
    "version": "[4.*, 5.0.0)"
  }
}
```

```json:local.settings.json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "<language worker>",
    "AzureWebJobsStorage": "<connection-string>",
    "MyBindingConnection": "<binding-connection-string>",
    "AzureWebJobs.HttpExample.Disabled": "true"
  },
  "Host": {
    "LocalHttpPort": 7071,
    "CORS": "*",
    "CORSCredentials": false
  },
  "ConnectionStrings": {
    "SQLConnectionString": "<sqlclient-connection-string>"
  }
}
```

```txt:requirements.txt
requests==2.32.5
pyodbc==5.3.0
pandas==2.3.3
azure-storage-blob==12.27.1
openpyxl==3.1.5
azure-functions==1.24.0
```

そしてそれら4ファイルをエクスプローラーで同時選択し、右クリックからZipファイルに圧縮します。

こうしてできた1つのZipファイルをデプロイすることになります。

## 3. Zipファイルデプロイ

```bash
az functionapp deployment source config-zip --src <zipファイル名> --name <Azure Functionsアプリ名> --resource-group <リソースグループ名> --build-remote true
```

コマンドにもあるように、flex consumptionプランだと**remote build**という便利な機能があるため手元でpythonをビルドする必要がなく、またrequirements.txtさえ用意しておけばリモートでそれらの依存ライブラリを自動でインストールしてくれるため楽にデプロイができて良かったです。

これで以下のようにログが出たら成功です。

```log
Getting scm site credentials for zip deployment
Starting zip deployment. This operation can take a while to complete ...
Deployment endpoint responded with status code 202 for deployment id "xxxxx"
Waiting for sync triggers...
Checking the health of the function app
"Deployment was successful."
```

正常にデプロイできていればAzure Functionsのページの「名前」欄に新しくアプリが追加されているはずです。
![](https://storage.googleapis.com/zenn-user-upload/190fe7c3a3eb-20260106.png)

あとはコードが正しければpython内のTimer Trigger(v2)記法で指定した時間に定期実行されます。

またAzure Functions該当アプリのページから「テスト/実行」タブを開き、「実行」を押すと手動でも実行できます。
![](https://storage.googleapis.com/zenn-user-upload/eb3decb9711e-20260106.png)
