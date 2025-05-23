# Shonai_GASの使い方ガイド

## ✅ 自動生成されるファイル例

以下のような書類を自動で作成できます：

<img width="439" alt="example1" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/59597a90-e379-4ccb-a3db-556d06317318">
<img width="180" alt="example2" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/3e4bbeea-9d6f-4cf6-9736-a04499c5f948">

---

## 🔄 Shonai_GASとShonai_GAS_v2の違い

同じ日・同じ部屋で時間を変えて予約した場合の動作が異なります。

**Shonai_GAS の挙動：**  
<img width="485" alt="Shonai_GAS_behavior" src="https://github.com/user-attachments/assets/669f0b07-605b-4760-b3b0-a3545cd94e37">

**Shonai_GAS_v2 の挙動：**  
<img width="556" alt="Shonai_GAS_v2_behavior" src="https://github.com/user-attachments/assets/5d052c0d-4337-4e94-acfb-6b58fd0644e4">

---

## ⚠ 注意事項

- **自己責任で使用してください。**
- 提出前に内容に間違いがないか必ず確認してください。
- **特に1月の学内集会願（曜日が変わることあり）には注意！**

---

## 🔧 セットアップ手順

### 1. Google Apps Script の作成

1. [Google Apps Script](https://script.google.com/home/start) を開く。
2. 新しいプロジェクトを作成。
3. [コード.gs](/コード.gs) の内容をコピーして貼り付け。

---

### 2. テンプレートファイルの用意

1. [テンプレ.xlsx](/テンプレ.xlsx) をダウンロード。
2. [Google Drive](https://drive.google.com/) にアップロード。
3. スプレッドシートとして開き、「Googleスプレッドシート」として保存。

<img width="500" alt="テンプレ作成" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/bd1b4202-95c2-4874-9fff-8653b1f5f1b6">

---

### 3. 初期設定（コード.gs）

以下の情報をスクリプト内に入力します：

- **4行目 `sesshoURL`**：折衝表のスプレッドシートURL  
- **7行目 `circle`**：自分のサークル名  
- **13行目 `tmpURL`**：Googleスプレッドシートとして保存したテンプレのURL  
- **16行目 `id`**：出力先フォルダのID（ルートに保存する場合は `''` に）

<img width="500" alt="初期設定" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/639ed0e0-f0cb-4579-a501-d830bba2e0e0">

- 学内集会願の「人数」も入力可（空欄でもOK）
- **19行目：不要な部屋の削除で処理高速化可**
- その他、必要に応じてカスタマイズ可能。

> 💡 **2回目以降は、URLの差し替えだけでOK！**

---

## ▶ 実行方法

1. プロジェクトを保存し、「実行」ボタンをクリック。

<img width="500" alt="run" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/43e84635-1cc1-4b16-8897-6c6f9470fa07">

2. 「権限を確認」クリック  
3. Googleアカウントを選択  
4. 「詳細」 → 「無題のプロジェクトに移動（安全でないページ）」を選択  
5. 「許可」ボタンを押す  
6. GASが実行されます

> 💡 **2回目以降はこの手順（2〜5）はスキップ可能です**

---

## 📋 ファイルの取得＆コピペ方法

1. 実行ログに「実行完了」と出たら、ログ内のURLを開く  
　または保存フォルダから開く  
2. **xlsx形式でダウンロード**

<img width="500" alt="download" src="https://github.com/k4fn/Shonai_GAS/assets/121530063/95897698-efa8-4199-ba24-dedc8995ac60">

> ⚠ **スプレッドシートから直接コピペすると書式が崩れます！**

3. ダウンロードした `.xlsx` ファイルを開いてコピペすれば完了！

---
