#### こんなのを自動で作ってくれます！

<img width="439" alt="image" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/59597a90-e379-4ccb-a3db-556d06317318">
<img width="180" alt="image" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/3e4bbeea-9d6f-4cf6-9736-a04499c5f948">

# [Shonai_GAS](https://github.com/k4fn/Shonai_GAS)とShonai_GAS_v2の違い
同じ日に同じ部屋を、時間を開けて取ったときの挙動が違います。  
⬇Shonai_GAS  
<img width="485" alt="cccccccccc" src="https://github.com/user-attachments/assets/669f0b07-605b-4760-b3b0-a3545cd94e37">  
  
⬇Shonai_GAS_v2  
<img width="556" alt="aaaaaaaaaaaaaaaaaaaaaaaaaaaa" src="https://github.com/user-attachments/assets/5d052c0d-4337-4e94-acfb-6b58fd0644e4">

# 使い方
> [!CAUTION]
> 自己責任でお願いします。提出前に間違いが無いか自分で確認してください。**特に年の変わる1月の学内集会願の曜日**

## Google Apps Scriptの作成
[Google Apps Script](https://script.google.com/home/start)(以下GAS)を開き、新しいプロジェクトを作成する。[コード.gs](/コード.gs)のコードをコピーし、GASのコード.gsにコピーする。

## テンプレファイルの作成
[テンプレ.xlsx](/テンプレ.xlsx)をダウンロードして、[Google Drive](https://drive.google.com/)上にアップロードする。Google Driveにアップロードしたテンプレ.xlsxを開き、Googleスプレッドシートとして保存する。

<img width="500" alt="image" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/bd1b4202-95c2-4874-9fff-8653b1f5f1b6">


## 初期設定
- 折衝表のスプレッドシートのURLを、4行目のsesshoURLに入力する。
- 折衝表のスプレッドシートに記入している自分のサークル名を、7行目のcircleに入力する。
- 先ほど作成したGoogleスプレッドシートのテンプレのURLを、13行目のtmpURLに入力する。
- 作成するファイルを保存したいフォルダのidを16行目のidに入力する。ルートフォルダに保存したい場合は、''の中を空にする。
  
  　画像のxxxの部分
  
  <img width="500" alt="image" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/639ed0e0-f0cb-4579-a501-d830bba2e0e0">

- 学内集会願の人数のところに記入したい人数を入力する。空欄でも可。
- 19行目の所で、使わない部屋を消すと処理が早く終わります。
- 他にも変えたい所あったら変える。

>[!TIP]
>2回目以降は、折衝表のスプレッドシートのURLを変更するだけでOK

## GASの実行
1. 初期設定が終わったら、プロジェクトを保存し、実行する。
   
     <img width="500" alt="image" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/43e84635-1cc1-4b16-8897-6c6f9470fa07">

2. 権限を確認をクリックする。
3. 自分のGoogleアカウントを選択する。
4. 左下の詳細をクリックし、無題のプロジェクト（安全ではないページ）に移動をクリックする。
5. 許可をクリックする。
6. するとGASが実行される。
>[!TIP]
>2回目以降は、2~5は省略できます。
## コピペの方法
1. 実行ログに実行完了と出力されたら、実行ログに出力されているURLを開く。（さっき指定した保存したいフォルダからも開けます。）
2. URLを開いたら、xlsx形式でダウンロードする。
   
   <img width="500" alt="image" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/95897698-efa8-4199-ba24-dedc8995ac60">

> [!CAUTION]
> **xlsx形式からではなく、スプレッドシートからコピペすると書式が崩れます。**

3. ダウンロードしたxlsxファイルを開き、コピペすれば完了！

## その他
何か動かないとかあったら、Issuesで教えてほしいです
