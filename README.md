# マリオカート 8DX ラウンジ戦績記録 VBA

## 対応バージョン

マリオカート 8DX コース追加パス**第 6 弾**に対応済み。

## 概要

マリオカート 8DX のラウンジ等での戦績を記録し、結果を分析できるツール。  
VBA でデータを記録し、ピボットテーブルでそれを分析する。  
[以前作成したもの](https://github.com/usui324/MK8DX_track_db)の機能追加版。

### 画面構成

1. データ登録シート
<img height="300px" src="https://github.com/usui324/MK8DX_track_db_02/assets/54677286/1db331ef-1861-46b9-b6cd-246c36654594"/>

2. グラフシート（ピボットテーブル）
<img height="300px" src="https://github.com/usui324/MK8DX_track_db_02/assets/54677286/40144af5-8257-4244-b3ff-312570d01435"/>

3. データシート
<img src="https://github.com/usui324/MK8DX_track_db_02/assets/54677286/f5412e92-f44a-4716-a1ad-88503aadb1c7"/>

### 収集データ一覧

1. コース名
2. スタート位置
3. 結果順位
4. 備考
5. 模擬 Tier
6. 形式
7. 日付

### 他の機能

1. アイテムテーブルの表示（2024/01/28 追加)
コース名を選択したときにアイテムテーブルの画像が表示される機能。  
画像は[Mario Kart Blog様](http://japan-mk.blog.jp/mk8dx.info-4)で配布されているものを使用。

2. 多言語対応（仮）
一部の表示が英語（English）に対応。  
設定（Settings）シートから「言語 / Language」を変更することで英語での表示が可能。

3. 規定レース数の設定
ピボットテーブルで各コースのデータを平均得点順、平均順位順にソートするとき、「最低何レース走ったコースを表示対象にするか」を設定できる。  
設定（Settings）シート「規定レース数」にて設定できる。

