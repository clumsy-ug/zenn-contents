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
    - Excelの[`SUM`](https://support.microsoft.com/ja-jp/office/sum-%E9%96%A2%E6%95%B0%E3%82%92%E4%BD%BF%E3%81%A3%E3%81%A6%E7%AF%84%E5%9B%B2%E5%86%85%E3%81%AE%E6%95%B0%E5%80%A4%E3%82%92%E5%90%88%E8%A8%88%E3%81%99%E3%82%8B-323569b2-0d2b-4e7b-b2f8-b433f9f0ac96), [`AVERAGE`](https://support.microsoft.com/ja-jp/office/average-%E9%96%A2%E6%95%B0-047bac88-d466-426c-a32b-8f33eb960cf6)関数を使用

9.  各サマリシートを参照している各グラフシートの、データ範囲を1列分拡張する
    - ひと月ごとに1列サマリシートの列が追加されていくため

10. 完成

[^1]: [pandas](https://pandas.pydata.org/)の[`read_sql()`](https://pandas.pydata.org/docs/reference/api/pandas.read_sql.html)メソッドは、第一引数のSQLテキストを読み取り、実行までしてくれる便利なものですが、読み取れるSQLステートメントの数が最大1つまでという制約があります。例えば以下のコードはエラーになります。`USE`, `SELECT`という2つのステートメントがあるからですね。
    ```sql
    USE [Sample-Database];
    SELECT * FROM [Sample-Table];
    ```

---

これを部分的にでも可能な限り自動化したいと思い少しずつ試してみたら、最終的に完全自動化をすることができ、3時間が0になった(やったね)ので、その方法やプログラムを共有します。


# 下書きメモ

- Azure Functionsはあくまで実行環境で、それを**定期**実行させたいならTimer Trigger機能を使う必要がある。
で、これはazure functionsに内包されている機能でgui上でポチポチ設定するだけでやれるんだろうなとか思っていたが、そうではなく完全にコードベースで、timer trigger機能を使用するためのpythonの書き方、というのがバージョンごとに分かれて存在しており、それに従ってコードを書く必要がある。まぁほぼコピペでいけて簡単なので問題はない。

- timer trigger v2のcron式、utcだから気をつけて。自分はdatetime.now()をdatetime.now(ZoneInfo('Asia/Tokyo'))に変更することで解決した。

- flex consumptionだとremote buildという便利な機能があり、手元でpythonをビルドする必要がなく、またrequirements.txtさえ用意しておけばリモートでそれを自動でインストールしてくれるっぽく、凄く楽にデプロイ成功して良かった。

- 最初matplotlibで図を作ろうとしていたが、既存のexcelの図をopenpyxlで取得していじるところまでやっちゃえば良いことに気づいて楽にグラフを拡張できてよかった。スタイル引き継げるのが良い

- 最初はローカルでこういう風に実行して試してた。毎回sql走って実行を5~10分待つのが面倒だったが、本番に近い状態で正常に実行されるか常に確認しておきたかったので我慢した

- 最後の方は急いでいてリファクタしてないのだが、ご愛嬌

- 余談だがexcelなどのmsアプリのファイルは内部でxmlになっていること、だからこそ .xlsx / .docx / .pptx のxはxmlのxとして使われていること、を知った。で、グラフのタイトルが消えてしまうときとかにxmlの中身を見に行ってxmlが存在しないからxmlとは違って内部キャッシュを恐らくexcelはもっていてそこが消えてしまったんだろうとかそういうアタリをつけられて、楽しかった。実際に1文字消す、戻す、とやってxmlにそれが認識されていることを確認し、再度実行したらちゃんと直ったので、良かった。
 
 - 集計表がおかしいとか、ある時点での集計表をもう一度見たいとか、集計表を見る営業サイドだけでなく開発サイド(主に私)も過去の集計表が見たいということで、月ごとの集計表はファイル名の末尾に`_old`をつけて一応スナップショットとして保存しておくことにしている

- 単語のリンクをハイパーリンクでのせるべきところはのせとく
