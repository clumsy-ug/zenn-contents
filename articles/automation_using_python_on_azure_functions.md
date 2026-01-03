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

今後は月に1度そのコンテナを見に行って、作成されているExcel(実体はBLOB)ファイルをダウンロードするだけで集計表が手に入る状態になりました。やったね。

GitHubにPythonファイルやrequirements.txtなどをまとめたリポジトリを上げています。

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
    - 余談ですが、このVLOOKUPの第四引数を省略してあいまい検索がかかってしまい、意図しない結果が入っていたことに1カ月後の集計時に気づくという[失態](https://x.com/clumsy_ug/status/2003330818111156456?s=46)を犯しました（泣）

8. 以上で完成した各一覧表シートを基に、列ごとの合計値や平均値をとった結果を記載するサマリシートを複数作成
    - 厳密には、すでに作成されている各サマリシートに1列追加して、その月の分として新しい値を入力していく
    - Excelの`SUM`, `AVERAGE`関数を使用

9.  各サマリシートを参照している各グラフシートの、データ範囲を1列分拡張する
    - 毎月サマリシートに1列追加されていくため

10. 完成

[^1]: [pandas](https://pandas.pydata.org/)の[`read_sql()`](https://pandas.pydata.org/docs/reference/api/pandas.read_sql.html)メソッドは、第一引数のSQLテキストを読み取り、実行までしてくれる便利なものですが、読み取れるSQLステートメントの数が最大1つまでという制約があります。例えば以下のコードはエラーになります。`USE`, `SELECT`という2つのステートメントがあるからですね。
    ```sql
    USE [Sample-Database];
    SELECT * FROM [Sample-Table];
    ```

---

これを部分的にでも可能な限り自動化したいと思い少しずつ試してみたら、最終的に完全自動化をすることができ3時間が0になったので、その方法を説明します。

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

これをAzure Functionsで提供されているTimer Triggerという機能の最新である[バージョン2(v2)の記法](https://learn.microsoft.com/ja-jp/azure/azure-functions/functions-bindings-timer?tabs=python-v2%2Cisolated-process%2Cnodejs-v4&pivots=programming-language-python#example)を使って、自動で定期実行します。今回の例では毎月20日の午前4時にしています。

（Timer Triggerという機能がAzure Functionsにあるということは知っていましたが、GUI上でポチポチ設定するのではなく決まった書式でコードに落とし込むことで利用できる機能であるということを初めて知りました）

公式の例に則り、以下のように書くことで定期実行されるようにします。

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

## 2. 実データを取得→DFに変換

```python:function_app.py
import requests
import pandas as pd
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
    # まず.sqlファイルは無視してエクセルだけlistに格納していく
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

## 7. ああ

# 下書きメモ

- flex consumptionだとremote buildという便利な機能があり、手元でpythonをビルドする必要がなく、またrequirements.txtさえ用意しておけばリモートでそれを自動でインストールしてくれるっぽく、凄く楽にデプロイ成功して良かった。

- 最初matplotlibで図を作ろうとしていたが、既存のexcelの図をopenpyxlで取得していじるところまでやっちゃえば良いことに気づいて楽にグラフを拡張できてよかった。スタイル引き継げるのが良い

- 最初はローカルでこういう風に実行して試してた。毎回sql走って実行を5~10分待つのが面倒だったが、本番に近い状態で正常に実行されるか常に確認しておきたかったので我慢した

- 余談だがexcelなどのmsアプリのファイルは内部でxmlになっていること、だからこそ .xlsx / .docx / .pptx のxはxmlのxとして使われていること、を知った。で、グラフのタイトルが消えてしまうときとかにxmlの中身を見に行ってxmlが存在しないからxmlとは違って内部キャッシュを恐らくexcelはもっていてそこが消えてしまったんだろうとかそういうアタリをつけられて、楽しかった。実際に1文字消す、戻す、とやってxmlにそれが認識されていることを確認し、再度実行したらちゃんと直ったので、良かった。
 
 - 集計表がおかしいとか、ある時点での集計表をもう一度見たいとか、集計表を見る営業サイドだけでなく開発サイド(主に私)も過去の集計表が見たいということで、月ごとの集計表はファイル名の末尾に`_old`をつけて一応スナップショットとして保存しておくことにしている

- 単語のリンクをハイパーリンクでのせるべきところはのせとく
