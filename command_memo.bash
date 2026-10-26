# articles

## articles/what-is-slug.md を作成
npx zenn new:article --slug what-is-slug

# scraps/

## Scrap作成（1.mdを作成し、そこに最初のコメントを書いておく）
npx zenn scrap create --title "eMudhraの件を雰囲気だけ理解" --file "1.md"

## url.txtを作っておいてそこに作成されたScrapのurlを貼り付けておく

## タグをつける
npx zenn scrap update "$(cat url.txt)" --topics "eMudhra,証明書"
## どんどんコメント追加していく
npx zenn scrap post "$(cat url.txt)" --file "2.md"

## コメントを更新
### まずそのコメント部分のslugを取得
npx zenn scrap get "$(cat url.txt)"

### 更新
npx zenn scrap update-comment "$(cat url.txt)" <先ほどのslug> --file "1.md"  # 1.mdを更新

