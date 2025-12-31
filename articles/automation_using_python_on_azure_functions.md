---
title: "（執筆中）Azure Functions + Python(openpyxl)で毎月のExcel集計を自動化した"
emoji: "🎍"
type: "tech"
topics: ["azurefunctions", "python", "openpyxl", "excel"]
published: true
---

# 下書きメモ

- Azure Functionsはあくまで実行環境で、それを**定期**実行させたいならTimer Trigger機能を使う必要がある。
で、これはazure functionsに内包されている機能でgui上でポチポチ設定するだけでやれるんだろうなとか思っていたが、そうではなく完全にコードベースで、timer trigger機能を使用するためのpythonの書き方、というのがバージョンごとに分かれて存在しており、それに従ってコードを書く必要がある。まぁほぼコピペでいけて簡単なので問題はない。

- timer trigger v2のcron式、utcだから気をつけて。自分はdatetime.now()をdatetime.now(ZoneInfo('Asia/Tokyo'))に変更することで解決した。

- flex consumptionだとremote buildという便利な機能があり、手元でpythonをビルドする必要がなく、またrequirements.txtさえ用意しておけばリモートでそれを自動でインストールしてくれるっぽく、凄く楽にデプロイ成功して良かった。

- 最初matplotlibで図を作ろうとしていたが、既存のexcelの図をopenpyxlで取得していじるところまでやっちゃえば良いことに気づいて楽にグラフを拡張できてよかった。スタイル引き継げるのが良い

- 最初はローカルでこういう風に実行して試してた。毎回sql走って実行を5~10分待つのが面倒だったが、本番に近い状態で正常に実行されるか常に確認しておきたかったので我慢した

- 最後の方は急いでいてリファクタしてないのだが、ご愛嬌

- 余談だがexcelなどのmsアプリのファイルは内部でxmlになっていること、だからこそ .xlsx / .docx / .pptx のxはxmlのxとして使われていること、を知った。で、グラフのタイトルが消えてしまうときとかにxmlの中身を見に行ってxmlが存在しないからxmlとは違って内部キャッシュを恐らくexcelはもっていてそこが消えてしまったんだろうとかそういうアタリをつけられて、楽しかった。実際に1文字消す、戻す、とやってxmlにそれが認識されていることを確認し、再度実行したらちゃんと直ったので、良かった。
 