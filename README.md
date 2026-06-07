# Word Style Unifier

Word(.docx)内の文字種に応じてフォントを統一し、結果をZIPでダウンロードするWebサービスです。

## 起動方法

### ローカル(Python)
1. 依存をインストール
   - `pip install -r requirements.txt`
2. サーバー起動
   - `uvicorn app.main:app --reload --host 0.0.0.0 --port 8000`
3. ブラウザ
   - `http://localhost:8000`

### Docker
- `docker compose up --build`

## 実装済み仕様（フェーズ1）
- 単一ファイル/フォルダ(再帰)アップロード
- .docxのみ変換対象（それ以外はスキップ）
- 半角英数字記号(ASCII)はフォントA
- それ以外はフォントB
- 半角カタカナは全角カタカナへ変換
- 本文とテーブルを中心に、ヘッダー/フッターも処理
- 同期処理（目安タイムアウト5分）
- 常にZIPでダウンロード
- 成功分のみZIPに格納
- 失敗・スキップ対象をZIPルートの `faield.txt` に出力
- 一時ファイルは処理完了後に削除（ZIPはメモリ保持）

## 制限
- 1ファイルあたり上限10MB
