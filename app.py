import os
import random

from openpyxl import load_workbook
from flask import (
    Flask, render_template, request, redirect, url_for, abort, session
)

# Templates フォルダ（大文字）を明示的に指定（Windows/他OS どちらでも動くように）
app = Flask(__name__, template_folder="Templates")

# セッション（出題順・得点の保持）の署名に使う鍵。
# 本番では必ず環境変数 SECRET_KEY を設定すること（未設定なら開発用の値を使う）。
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-in-production")
if not app.debug and app.secret_key == "dev-secret-key-change-in-production":
    app.logger.warning(
        "SECRET_KEY が未設定です。本番では環境変数 SECRET_KEY を必ず設定してください。"
    )

# ---- クイズデータ ----------------------------------------------------------
# 問題と解答は questions.xlsx（Excel）で管理する（4択式）。
#
# 1行 = 1問。列（1行目はヘッダー、2行目以降がデータ）：
#   section_id     : URL に使う英数字のID（同じ単元の行には同じIDを書く）
#   section_title  : トップ画面に表示する単元名
#   question       : 問題文
#   choice1〜4      : 選択肢（最低2つ。空欄の選択肢は無視される）
#   answer         : 正解の選択肢番号（1〜4）
#   explanation    : 解説（任意。結果画面に表示。空欄可）
# 単元の表示順・問題の出題順は、行の並び順のとおり。
CHOICE_COLUMNS = ["choice1", "choice2", "choice3", "choice4"]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
QUESTIONS_FILE = os.path.join(BASE_DIR, "questions.xlsx")


def load_sections():
    """questions.xlsx を読み込んでセクションのリストを返す"""
    wb = load_workbook(QUESTIONS_FILE, read_only=True, data_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        raise ValueError("questions.xlsx が空です。")

    # ヘッダー名 → 列番号（列の順番が違っても動くように名前で対応づける）
    index = {}
    for i, name in enumerate(header):
        if name is not None:
            index[str(name).strip()] = i
    required = ["section_id", "section_title", "question",
                "choice1", "choice2", "answer"]
    missing = [c for c in required if c not in index]
    if missing:
        raise ValueError(
            "questions.xlsx のヘッダーに次の列が必要です: " + ", ".join(missing)
        )

    def cell(row, name):
        """存在しない列（explanation 未定義など）にも安全にアクセスする"""
        i = index.get(name)
        return row[i] if i is not None and i < len(row) else None

    sections = []
    by_id = {}
    for line_no, row in enumerate(rows, start=2):  # 2行目からがデータ
        # 完全に空の行はスキップ
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue

        sec_id = str(cell(row, "section_id") or "").strip()
        sec_title = str(cell(row, "section_title") or "").strip()
        question = cell(row, "question")
        answer_raw = cell(row, "answer")

        if not sec_id or question is None or answer_raw is None:
            continue  # 必須項目が欠けた行はスキップ

        # 選択肢（空欄は除外）。番号と対応づけるため (番号, テキスト) で持つ
        numbered_choices = []
        for n, col in enumerate(CHOICE_COLUMNS, start=1):
            val = cell(row, col)
            if val is not None and str(val).strip() != "":
                numbered_choices.append((n, str(val).strip()))
        if len(numbered_choices) < 2:
            raise ValueError(f"{line_no}行目: 選択肢が2つ未満です。")

        # 正解番号（1〜4）→ choices の何番目か
        try:
            answer_no = int(float(str(answer_raw).strip()))
        except ValueError:
            raise ValueError(
                f"{line_no}行目: answer は選択肢番号(1〜4)で指定してください。"
            )
        valid_numbers = [n for n, _ in numbered_choices]
        if answer_no not in valid_numbers:
            raise ValueError(
                f"{line_no}行目: answer={answer_no} に対応する選択肢がありません。"
            )

        choices = [text for _, text in numbered_choices]
        answer_index = valid_numbers.index(answer_no)  # 0始まりの位置
        explanation = cell(row, "explanation")
        explanation = str(explanation).strip() if explanation is not None else ""

        if sec_id not in by_id:
            section = {"id": sec_id, "title": sec_title, "questions": []}
            by_id[sec_id] = section
            sections.append(section)
        by_id[sec_id]["questions"].append({
            "question": str(question).strip(),
            "choices": choices,
            "answer_index": answer_index,
            "explanation": explanation,
        })

    wb.close()

    if not sections:
        raise ValueError("questions.xlsx に有効な問題がありません。")
    return sections


# Excel を読み直すためのキャッシュ。
# アクセスのたびに Excel の更新日時(mtime)を確認し、前回から変わっていれば
# 読み直す。変わっていなければキャッシュを使う（膨大な問題でも毎回パースしない）。
_cache = {"mtime": None, "sections": None, "by_id": {}}


def get_sections():
    """最新のセクション一覧を返す（Excel が更新されていれば読み直す）"""
    try:
        mtime = os.path.getmtime(QUESTIONS_FILE)
    except OSError:
        # ファイルが見つからない等。直前の内容があればそれを使う。
        if _cache["sections"] is not None:
            return _cache["sections"]
        raise

    if mtime != _cache["mtime"]:
        try:
            sections = load_sections()
        except Exception as e:
            # Excel を編集中で読めない・書式ミス等。
            # 直前に読めた内容があれば、それで動かし続ける。
            if _cache["sections"] is not None:
                app.logger.warning("questions.xlsx を読み直せませんでした: %s", e)
                return _cache["sections"]
            raise
        _cache["mtime"] = mtime
        _cache["sections"] = sections
        _cache["by_id"] = {s["id"]: s for s in sections}

    return _cache["sections"]


def get_section(section_id):
    """セクションを取得（存在しなければ 404）"""
    get_sections()  # 最新の状態に更新
    section = _cache["by_id"].get(section_id)
    if section is None:
        abort(404)
    return section


def current_quiz(section_id):
    """このセクションで進行中のセッションを返す（無ければ None）"""
    quiz = session.get("quiz")
    if not quiz or quiz.get("section_id") != section_id:
        return None
    return quiz


@app.route("/")
def index():
    """トップ画面：セクション（単元）を選ぶ"""
    sections = [
        {"id": s["id"], "title": s["title"], "count": len(s["questions"])}
        for s in get_sections()
    ]
    return render_template("index.html", sections=sections)


@app.route("/start/<section_id>")
def start(section_id):
    """セクション開始：出題順をシャッフルし、得点をリセットする"""
    section = get_section(section_id)
    order = list(range(len(section["questions"])))
    random.shuffle(order)  # 出題順をシャッフル
    session["quiz"] = {
        "section_id": section_id,
        "order": order,
        "pos": 0,      # 現在の出題位置（0始まり）
        "score": 0,    # 正解数
    }
    session.pop("last", None)
    return redirect(url_for("quiz", section_id=section_id))


@app.route("/quiz/<section_id>")
def quiz(section_id):
    """クイズ画面（問題文＋4つの選択肢）"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    questions = section["questions"]
    order = quiz_state["order"]
    pos = quiz_state["pos"]

    # 全問終了していれば結果画面へ
    if pos >= len(order):
        return redirect(url_for("result", section_id=section_id))

    qidx = order[pos]
    # Excel が編集されて問題数が変わった場合などの保険
    if qidx >= len(questions):
        return redirect(url_for("start", section_id=section_id))

    q = questions[qidx]
    # 選択肢を (元の番号, テキスト) にして表示順だけシャッフルする。
    # 送信値には「元の番号」を使うので、並びを変えても採点は正しい。
    choices = list(enumerate(q["choices"], start=1))
    random.shuffle(choices)

    return render_template(
        "quiz.html",
        section_id=section_id,
        section_title=section["title"],
        question=q["question"],
        choices=choices,
        number=pos + 1,
        total=len(order),
    )


@app.route("/answer/<section_id>", methods=["POST"])
def answer(section_id):
    """採点してセッションを進め、フィードバック画面へリダイレクト（PRG）"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    questions = section["questions"]
    order = quiz_state["order"]
    pos = quiz_state["pos"]
    if pos >= len(order):
        return redirect(url_for("result", section_id=section_id))

    qidx = order[pos]
    if qidx >= len(questions):
        return redirect(url_for("start", section_id=section_id))
    q = questions[qidx]

    # 送信された選択肢番号（1始まり）を 0始まりの位置に変換
    try:
        chosen_index = int(request.form.get("choice", "")) - 1
    except ValueError:
        chosen_index = -1

    if not (0 <= chosen_index < len(q["choices"])):
        # 未選択・不正な値ならクイズ画面に戻す
        return redirect(url_for("quiz", section_id=section_id))

    is_correct = chosen_index == q["answer_index"]
    if is_correct:
        quiz_state["score"] += 1

    # フィードバック表示に必要な情報を保存（リロードしても再採点されないように）
    session["last"] = {
        "section_id": section_id,
        "is_correct": is_correct,
        "user_answer": q["choices"][chosen_index],
        "correct_answer": q["choices"][q["answer_index"]],
        "explanation": q["explanation"],
    }
    quiz_state["pos"] = pos + 1  # 次の問題へ進める
    session.modified = True

    return redirect(url_for("feedback", section_id=section_id))


@app.route("/feedback/<section_id>")
def feedback(section_id):
    """直前の解答に対する正誤＋解説を表示（リロード安全）"""
    section = get_section(section_id)
    last = session.get("last")
    quiz_state = current_quiz(section_id)
    if last is None or last.get("section_id") != section_id or quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    has_next = quiz_state["pos"] < len(quiz_state["order"])
    template = "correct.html" if last["is_correct"] else "incorrect.html"
    return render_template(
        template,
        section_id=section_id,
        section_title=section["title"],
        user_answer=last["user_answer"],
        correct_answer=last["correct_answer"],
        explanation=last["explanation"],
        has_next=has_next,
    )


@app.route("/result/<section_id>")
def result(section_id):
    """セクション終了：得点を表示する"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    score = quiz_state["score"]
    total = len(quiz_state["order"])
    percent = round(score / total * 100) if total else 0
    return render_template(
        "result.html",
        section_id=section_id,
        section_title=section["title"],
        score=score,
        total=total,
        percent=percent,
    )


if __name__ == "__main__":
    # ローカル開発用の簡易サーバ。debug は環境変数で明示的に有効化する。
    #   通常起動      : python app.py
    #   デバッグ起動  : FLASK_DEBUG=1 python app.py  （PowerShell: $env:FLASK_DEBUG=1）
    # 本番では gunicorn 等の WSGI サーバから app を起動する（このブロックは使わない）。
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug)
