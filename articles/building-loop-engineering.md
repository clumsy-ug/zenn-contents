---
title: "Backlogを起点としたループエンジニアリングの構築"
emoji: "🔁"
type: "tech"
topics: ["loopengineering", "backlog", "teams"]
published: false
---

最近社内で Backlog -> Claude Code -> Teams　の流れを繰り返すループエンジニアリングを構築することができたので、その体験記です。

以下のイメージです。

# 課題

何に困っていたか

# やったこと

最初はSkillで1分おきにpollingするようなSkillを作ろうと思ってたんですが、冷静に考えたら単なるcronジョブじゃないかときづき、powershellスクリプトにしました。
skillだとclaudeが動くので、1分おきにトークンを消費するというのがアホらしく、考えてみたらシェルで良いじゃんということに気づけたのはよかったです。

# 他の選択肢

# 結果（定性）

# 結果（定量）

# 今後

残る課題は〜〜。
個人の興味としてはグラフエンジニアリングを組んでみたい。
