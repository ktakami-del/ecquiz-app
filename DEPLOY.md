# デプロイ手順書（Flask + Excel → Render）

経済思想史クイズアプリを本番公開するための手順と運用メモ。

---

## 構成の概要

```
ブラウザ ──HTTPS──▶ Render (gunicorn + Flask) ──読み込み──▶ questions.xlsx
```

- アプリ本体: `app.py`（Flask）
- 画面: `Templates/`（Jinja2テンプレート）
- 問題データ: `questions.xlsx`（Excelで管理。空白行は無視される）
- 本番サーバ: `gunicorn`
- ホスティング: Render（Blueprint = `render.yaml` で自動構築）

---

## フェーズ1：コードを本番仕様に修正（`app.py`）

| 変更 | 内容 | 理由 |
|---|---|---|
| `debug` の環境変数化 | `app.run(debug=True)` → `FLASK_DEBUG` で制御（本番は無効） | デバッグ画面からのコード実行脆弱性を防ぐ |
| `SECRET_KEY` の環境変数化 | 環境変数から読み、未設定時は警告 | セッション（得点・進捗）の署名に必要 |

ローカルでのデバッグ起動:
```bash
# PowerShell
$env:FLASK_DEBUG=1; python app.py
# 通常起動（デバッグ無効）
python app.py
```

---

## フェーズ2：デプロイ用ファイル

| ファイル | 役割 |
|---|---|
| `render.yaml` | Render Blueprint。起動コマンド・`SECRET_KEY`自動生成・Pythonバージョン固定を定義 |
| `Procfile` | 起動コマンド `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2` |
| `.gitignore` | `.venv`・キャッシュ・Excelロックファイル（`~$*.xlsx`）を除外 |
| `requirements.txt` | 本番サーバ `gunicorn` を追加 |

---

## フェーズ3：動作検証

- 実データで問題が全件読み込めるか確認（空白行スキップは正常動作）
- 「トップ → クイズ → 採点 → 解説 → 結果」の一連の流れを確認
- `debug=False` になっていることを確認

---

## フェーズ4：Gitでバージョン管理

```bash
git init -b main
git add .
git commit -m "初回コミット"
```

- `.venv` はコミットしない（本番機では作り直す）
- `questions.xlsx` はコミットする（デプロイ時に必要）

---

## フェーズ5：GitHubへ公開

```bash
git remote add origin https://github.com/<ユーザー名>/ecquiz-app.git
git push -u origin main
```

GitHub CLI（`gh`）が使える場合はリポジトリ作成も含めて実行できる:
```bash
gh repo create ecquiz-app --public   # または --private
git push -u origin main
```

---

## フェーズ6：Renderで公開

1. https://render.com に GitHub アカウントでサインイン
2. **New +** → **Blueprint**
3. リポジトリ `ecquiz-app` を選択
4. Blueprint Name: 任意（例 `ecquiz-app`） / Branch: **`main`**
5. **Apply** → ビルド → 起動 → URL 発行（`https://ecquiz-app-xxxx.onrender.com`）

`SECRET_KEY` は `render.yaml` の `generateValue: true` により Render が自動生成するため、手動設定は不要。

---

## つまずいた点と対処（今回の記録）

### エラー: `Blueprint file render.yaml not found on main branch`
- **原因**: 非公開(private)リポジトリの中身を Render が読めなかった。
- **対処**:
  1. リポジトリを公開(public)に変更（`gh repo edit <repo> --visibility public --accept-visibility-change-consequences`）、
     または GitHub の Settings → Applications → Render → Configure で対象リポジトリへのアクセスを許可（非公開のまま可）。
  2. 失敗した Blueprint は **削除して作り直す**（Render が失敗状態をキャッシュするため、Retry だけでは直らないことがある）。

> 補足: リポジトリの公開/非公開は「問題データ(`questions.xlsx`)の答えを世界中から読めるか」を左右する。
> アプリの利用者を限定したいだけなら、リポジトリ公開ではなく **アプリ側に簡易パスワード認証** を足すのが本来の方法。

---

## 今後の運用：問題の更新

```bash
# questions.xlsx を編集・保存したあと
git add questions.xlsx
git commit -m "問題を追加"
git push          # → Render が自動で再デプロイ（数分で反映）
```

---

## 補足メモ

- **無料プランの挙動**: 15分アクセスがないとスリープし、次回アクセスの初回のみ起動に30〜50秒かかる。常時即応が必要なら有料プランでスリープ無効化。
- **得点は利用者ごと**: ブラウザのセッション(Cookie)で管理するため、複数人が同時に使っても点数は混ざらない。
- **Excelの列構成**: `section_id / section_title / question / choice1〜4 / answer(1〜4) / explanation`。
- **本番機でのvenv**: コピーせず必ず新規作成（`python -m venv .venv` → `pip install -r requirements.txt`）。venvをコピーするとランチャー(`pip.exe`等)のパスが壊れる。
