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
# 問題と解答は questions.xlsx（Excel）で管理する（一問一答式）。
#
# 1行 = 1問。列（1行目はヘッダー、2行目以降がデータ）：
#   section_id     : URL に使う英数字のID（同じ単元の行には同じIDを書く）
#   section_title  : トップ画面に表示する単元名
#   question       : 問題文
#   answer         : 模範解答（語句など）。表示するだけで自動採点はしない
#   explanation    : 解説（任意。解答表示画面に出る。空欄可）
# 単元の表示順は、行の並び順のとおり。問題の出題順はシャッフルされる。
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
QUESTIONS_FILE = os.path.join(BASE_DIR, "questions.xlsx")

# 出題数の選択肢。単元の問題数を超えるものは表示せず、末尾に「全問」を足す。
COUNT_CHOICES = [10, 30, 50, 100]


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
    required = ["section_id", "section_title", "question", "answer"]
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
    for row in rows:  # 2行目からがデータ
        # 完全に空の行はスキップ
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue

        sec_id = str(cell(row, "section_id") or "").strip()
        sec_title = str(cell(row, "section_title") or "").strip()
        question = cell(row, "question")
        answer = cell(row, "answer")

        # 必須項目（単元ID・問題文・模範解答）が欠けた行はスキップ
        if not sec_id or question is None or answer is None:
            continue
        question = str(question).strip()
        answer = str(answer).strip()
        if not question or not answer:
            continue

        explanation = cell(row, "explanation")
        explanation = str(explanation).strip() if explanation is not None else ""

        if sec_id not in by_id:
            section = {"id": sec_id, "title": sec_title, "questions": []}
            by_id[sec_id] = section
            sections.append(section)
        by_id[sec_id]["questions"].append({
            "question": question,
            "answer": answer,
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


def current_question(section, quiz_state):
    """いま出題中の問題を返す（全問終了・Excel 変更時は None）"""
    questions = section["questions"]
    pos = quiz_state["pos"]
    if pos >= len(quiz_state["order"]):
        return None
    qidx = quiz_state["order"][pos]
    if qidx >= len(questions):  # Excel が編集されて問題数が減った場合の保険
        return None
    return questions[qidx]


@app.route("/")
def index():
    """トップ画面：セクション（単元）を選ぶ"""
    sections = [
        {"id": s["id"], "title": s["title"], "count": len(s["questions"])}
        for s in get_sections()
    ]
    return render_template("index.html", sections=sections)


@app.route("/setup/<section_id>")
def setup(section_id):
    """出題数を選ぶ画面（10問・30問・…・全問）"""
    section = get_section(section_id)
    total = len(section["questions"])
    # 用意した選択肢のうち、その単元の問題数を超えないものだけ出す。
    # 最後は必ず「全問」（総数が選択肢と重なる場合は重複させない）。
    counts = [c for c in COUNT_CHOICES if c < total] + [total]
    return render_template(
        "setup.html",
        section_id=section_id,
        section_title=section["title"],
        total=total,
        counts=counts,
    )


@app.route("/start/<section_id>")
def start(section_id):
    """セクション開始：出題順をシャッフルし、選ばれた問題数だけ出題する"""
    section = get_section(section_id)
    total = len(section["questions"])

    # 出題数（?count=10 など）。未指定・不正な値なら全問。
    try:
        count = int(request.args.get("count", total))
    except ValueError:
        count = total
    count = max(1, min(count, total))  # 1〜総数の範囲に収める

    order = list(range(total))
    random.shuffle(order)  # 出題順をシャッフル
    order = order[:count]  # 先頭 count 問だけを今回の出題にする
    return begin_quiz(section_id, order)


@app.route("/retry/<section_id>")
def retry(section_id):
    """間違えた問題だけをもう一度出題する（結果画面から）"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("setup", section_id=section_id))

    # 直前の回で「不正解」にした問題だけを出題対象にする。
    # Excel が編集されて問題が減った場合に備え、範囲外の番号は捨てる。
    wrong = [i for i in quiz_state.get("wrong", []) if i < len(section["questions"])]
    if not wrong:
        return redirect(url_for("result", section_id=section_id))

    random.shuffle(wrong)
    return begin_quiz(section_id, wrong)


def begin_quiz(section_id, order):
    """出題リストを受け取り、セッションを初期化してクイズ画面へ送る"""
    session["quiz"] = {
        "section_id": section_id,
        "order": order,
        "pos": 0,      # 現在の出題位置（0始まり）
        "score": 0,    # 自己採点で「正解」にした数
        "wrong": [],   # 「不正解」にした問題の番号（やり直し用）
    }
    session.pop("last", None)
    return redirect(url_for("quiz", section_id=section_id))


@app.route("/quiz/<section_id>")
def quiz(section_id):
    """出題画面（問題文＋解答の記入欄）"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    if quiz_state["pos"] >= len(quiz_state["order"]):
        return redirect(url_for("result", section_id=section_id))

    q = current_question(section, quiz_state)
    if q is None:
        return redirect(url_for("start", section_id=section_id))

    return render_template(
        "quiz.html",
        section_id=section_id,
        section_title=section["title"],
        question=q["question"],
        number=quiz_state["pos"] + 1,
        total=len(quiz_state["order"]),
    )


@app.route("/answer/<section_id>", methods=["POST"])
def answer(section_id):
    """解答を受け取り、答え合わせ画面へリダイレクト（PRG）。ここでは採点しない"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    q = current_question(section, quiz_state)
    if q is None:
        return redirect(url_for("result", section_id=section_id))

    user_answer = (request.form.get("answer") or "").strip()

    # 答え合わせ画面に必要な情報を保存。
    # pos も一緒に持ち、二重採点（戻る／リロード）を防ぐ。
    session["last"] = {
        "section_id": section_id,
        "pos": quiz_state["pos"],
        "user_answer": user_answer,
        "correct_answer": q["answer"],
        "explanation": q["explanation"],
    }
    session.modified = True

    return redirect(url_for("reveal", section_id=section_id))


@app.route("/reveal/<section_id>")
def reveal(section_id):
    """答え合わせ画面：模範解答と解説を表示し、自分で〇×を押してもらう"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    last = session.get("last")
    if (quiz_state is None or last is None
            or last.get("section_id") != section_id
            or last.get("pos") != quiz_state["pos"]):
        return redirect(url_for("quiz", section_id=section_id))

    return render_template(
        "reveal.html",
        section_id=section_id,
        section_title=section["title"],
        user_answer=last["user_answer"],
        correct_answer=last["correct_answer"],
        explanation=last["explanation"],
        number=quiz_state["pos"] + 1,
        total=len(quiz_state["order"]),
    )


@app.route("/grade/<section_id>", methods=["POST"])
def grade(section_id):
    """自己採点（〇 or ×）を受け取り、次の問題へ進む"""
    get_section(section_id)
    quiz_state = current_quiz(section_id)
    last = session.get("last")
    if (quiz_state is None or last is None
            or last.get("section_id") != section_id
            or last.get("pos") != quiz_state["pos"]):
        # 押し直し・リロードなど。二重に加点せずやり直させる。
        return redirect(url_for("quiz", section_id=section_id))

    if request.form.get("judge") == "correct":
        quiz_state["score"] += 1
    else:
        # 間違えた問題は控えておき、結果画面から復習・やり直しできるようにする
        quiz_state.setdefault("wrong", []).append(quiz_state["order"][quiz_state["pos"]])

    quiz_state["pos"] += 1  # 次の問題へ進める
    session.pop("last", None)
    session.modified = True

    if quiz_state["pos"] >= len(quiz_state["order"]):
        return redirect(url_for("result", section_id=section_id))
    return redirect(url_for("quiz", section_id=section_id))


@app.route("/result/<section_id>")
def result(section_id):
    """セクション終了：自己採点の結果を表示する"""
    section = get_section(section_id)
    quiz_state = current_quiz(section_id)
    if quiz_state is None:
        return redirect(url_for("start", section_id=section_id))

    score = quiz_state["score"]
    total = len(quiz_state["order"])
    percent = round(score / total * 100) if total else 0

    # 間違えた問題の一覧（復習用）。Excel が編集された場合に備えて範囲外は除く。
    questions = section["questions"]
    wrong = [questions[i] for i in quiz_state.get("wrong", []) if i < len(questions)]

    return render_template(
        "result.html",
        section_id=section_id,
        section_title=section["title"],
        score=score,
        total=total,
        percent=percent,
        wrong=wrong,
    )


if __name__ == "__main__":
    # ローカル開発用の簡易サーバ。debug は環境変数で明示的に有効化する。
    #   通常起動      : python app.py
    #   デバッグ起動  : FLASK_DEBUG=1 python app.py  （PowerShell: $env:FLASK_DEBUG=1）
    # 本番では gunicorn 等の WSGI サーバから app を起動する（このブロックは使わない）。
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug)
