#!/usr/bin/env python3
"""
Flash Card Game with an Excel backend.

Everything lives in one workbook (default: flashcards.xlsx):
  * "Cards"  - your deck. Edit it in Excel any time: add rows, fix typos, etc.
              Per-card stats (Times Seen / Times Correct) update after each game.
  * "Scores" - one row per finished game (powers the leaderboard).

Requires:  pip install openpyxl
Run:       python flashcards.py [path/to/deck.xlsx]
"""

import difflib
import random
import re
import sys
import time
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
DEFAULT_FILE = "flashcards.xlsx"
CARD_HEADERS = ["ID", "Category", "Question", "Answer",
                "Times Seen", "Times Correct", "Accuracy",
                "Wrong Choice 1", "Wrong Choice 2", "Wrong Choice 3"]
CATA_HEADERS = ["ID", "Category", "Question",
                "Correct Answers (separate with ;)",
                "Wrong Answers (separate with ;)",
                "Times Seen", "Times Correct", "Accuracy"]
SCORE_HEADERS = ["Date", "Player", "Category", "Questions", "Correct",
                 "Time (s)", "Best Streak", "Accuracy", "Mode", "Result"]

MAX_SCORES = 20          # leaderboard keeps the 20 most recent games
FUZZY_THRESHOLD = 0.85    # typo tolerance (1.0 = exact match only)

# (category, question, correct answer, [wrong choices for Multiple choice])
SAMPLE_DECK = [
    ("Geography", "Capital of Australia?", "Canberra", ["Sydney", "Melbourne", "Perth"]),
    ("Geography", "Largest ocean on Earth?", "Pacific", ["Atlantic", "Indian", "Arctic"]),
    ("Geography", "Capital of Canada?", "Ottawa", ["Toronto", "Vancouver", "Montreal"]),
    ("Geography", "Longest river in South America?", "Amazon", ["Orinoco", "Parana", "Magdalena"]),
    ("Science", "Chemical symbol for gold?", "Au", ["Ag", "Fe", "Pb"]),
    ("Science", "Hardest natural substance?", "Diamond", ["Quartz", "Graphite", "Topaz"]),
    ("Science", "Planet known as the Red Planet?", "Mars", ["Venus", "Jupiter", "Mercury"]),
    ("Science", "Powerhouse of the cell?", "Mitochondria", ["Nucleus", "Ribosome", "Golgi apparatus"]),
    ("Math", "Square root of 144?", "12", ["14", "11", "16"]),
    ("Math", "What is 7 x 8?", "56", ["54", "63", "49"]),
    ("Math", "Value of pi to two decimals?", "3.14", ["3.12", "3.41", "3.16"]),
    ("History", "Year World War II ended?", "1945", ["1944", "1946", "1939"]),
    ("History", "First President of the United States?", "George Washington",
     ["John Adams", "Thomas Jefferson", "Abraham Lincoln"]),
    ("History", "Ancient wonder located in Giza?", "Great Pyramid",
     ["Hanging Gardens", "Colossus of Rhodes", "Lighthouse of Alexandria"]),
    ("Languages", "How do you say 'thank you' in Spanish?", "Gracias",
     ["Por favor", "Hola", "Adios"]),
    ("Languages", "How do you say 'hello' in Japanese?", "Konnichiwa",
     ["Sayonara", "Arigato", "Ohayo"]),
]

# (category, question, [all correct answers], [wrong answers])
SAMPLE_CATA = [
    ("Geography", "Which of these are oceans?", ["Pacific", "Atlantic", "Indian"],
     ["Mediterranean", "Caribbean", "Baltic"]),
    ("Science", "Which of these are planets?", ["Mars", "Venus", "Jupiter"],
     ["Moon", "Sun", "Halley's Comet"]),
    ("Science", "Which of these are noble gases?", ["Helium", "Neon", "Argon"],
     ["Oxygen", "Nitrogen", "Hydrogen"]),
    ("Math", "Which of these are prime numbers?", ["2", "7", "11"],
     ["9", "15", "21"]),
    ("History", "Which of these were US presidents?", ["Abraham Lincoln", "Thomas Jefferson"],
     ["Benjamin Franklin", "Alexander Hamilton", "Napoleon Bonaparte"]),
    ("Languages", "Which of these are Romance languages?", ["Spanish", "French", "Italian"],
     ["German", "Japanese", "Russian"]),
]

DEFAULT_MESSAGES = {
    "pass": ["Excellent work, {name}! You passed!",
             "Great job, {name}! Keep it up!",
             "Nailed it, {name}!",
             "Impressive, {name}! You really know your stuff.",
             "Well done, {name}! Passing score reached."],
    "fail": ["Good effort, {name}! A little more practice and you'll get there.",
             "Don't give up, {name}! Every try makes you stronger.",
             "Almost there, {name}! Review the tricky ones and try again.",
             "Keep going, {name} - you're learning with every card!",
             "Mistakes help us learn. You've got this, {name}!"],
    "correct": ["Correct!", "Nice one!", "Exactly right!", "You got it!",
                "Spot on!", "Well done!"],
    "wrong": ["Not quite", "Almost!", "Oops, not this time", "Good try!",
              "Not this one", "Close, but no"],
    "skip": ["Skipped", "No worries - here's the answer",
             "Let's look at this one"],
}
DEFAULT_PASSING = 70
MESSAGE_COLUMNS = {"pass": "A", "fail": "B", "correct": "C", "wrong": "D",
                   "skip": "E"}

HELP_MESSAGE_LINES = [
    "",
    "MESSAGES sheet - the texts shown during the game.",
    "   - Pass Messages / Fail Messages: shown on the results screen, one picked at random.",
    "     A player passes when their score (% correct) reaches Passing Score (%) in the Setting / Value table.",
    "   - Correct / Wrong / Skipped Messages: shown after every answer, one picked at random each time.",
    "   - Use {name} for the player's name and {percent} for their score, e.g.  Great job, {name}!",
    "   - Add as many rows as you like. Blank cells are ignored; an empty column falls back to built-in texts.",
]

HELP_LINES = [
    "HOW TO EDIT THIS WORKBOOK",
    "",
    "CARDS sheet - one question per row (used by Identification, Multiple choice and True or False).",
    "   - Category, Question and Answer are required. Answer is the correct answer.",
    "   - Wrong Choice 1-3 are the wrong options that appear in Multiple choice. They are also used as the",
    "     false answers in True or False. Fill in only the ones you want (even just one).",
    "   - If all three Wrong Choice cells are blank, the game borrows other cards' answers automatically.",
    "",
    "CHECKALL sheet - one question per row (used by Check all that apply).",
    "   - Correct Answers: every right option, separated by a semicolon, e.g.  Pacific; Atlantic; Indian",
    "   - Wrong Answers: the options that should NOT be ticked, separated by a semicolon.",
    "   - Up to 6 options are shown per question. If you list more, the game picks 6 at random each time.",
    "",
    "SCORES sheet - one row per finished game. Only the 20 most recent games are kept; older ones are erased automatically.",
    "Times Seen, Times Correct and Accuracy are filled in by the game - don't edit them.",
] + HELP_MESSAGE_LINES + [
    "",
    "Close this file in Excel before playing, otherwise the game can't save your scores.",
]


def app_dir():
    """Folder of the .exe when frozen by PyInstaller, else of this script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


# --------------------------------------------------------------------------
# Excel storage layer
# --------------------------------------------------------------------------
class Store:
    """Reads/writes the workbook. The only place that touches Excel."""

    def __init__(self, path):
        self.path = Path(path)
        if not self.path.exists():
            self._create_sample()
        self.wb = load_workbook(self.path)
        for name in ("Cards", "Scores"):
            if name not in self.wb.sheetnames:
                raise SystemExit(f"Workbook is missing the '{name}' sheet.")
        self._ensure_schema()

    # ---- creation / upgrade ----------------------------------------------
    @staticmethod
    def _make_cata_sheet(wb):
        ws = wb.create_sheet("CheckAll", 1)
        ws.append(CATA_HEADERS)
        Store._style(ws, [6, 16, 46, 44, 44, 12, 14, 12])
        return ws

    @staticmethod
    def _make_messages_sheet(wb):
        ws = wb.create_sheet("Messages", 2)
        ws.append(["Pass Messages", "Fail Messages", "Correct Messages",
                   "Wrong Messages", "Skipped Messages", None, "Setting",
                   "Value"])
        Store._style(ws, [46, 52, 26, 30, 34, 3, 20, 10])
        ws["F1"].fill = PatternFill(fill_type=None)       # spacer column
        n = max(len(v) for v in DEFAULT_MESSAGES.values())
        for i in range(n):
            row = [DEFAULT_MESSAGES[k][i] if i < len(DEFAULT_MESSAGES[k])
                   else None for k in ("pass", "fail", "correct", "wrong",
                                       "skip")]
            ws.append(row + [None] + (["Passing Score (%)", DEFAULT_PASSING]
                                      if i == 0 else [None, None]))
        return ws

    @staticmethod
    def _make_help_sheet(wb):
        ws = wb.create_sheet("How to edit")
        for line in HELP_LINES:
            ws.append([line])
        ws.column_dimensions["A"].width = 110
        ws["A1"].font = Font(name="Arial", bold=True, size=13, color="1F4E78")
        ws.sheet_view.showGridLines = False

    def _ensure_schema(self):
        """Upgrade older workbooks in place (adds columns/sheets, keeps data)."""
        wb, changed = self.wb, False
        cards = wb["Cards"]
        for i in range(8, 11):
            c = cards.cell(row=1, column=i)
            if c.value is None:
                c.value = CARD_HEADERS[i - 1]
                c._style = copy(cards["G1"]._style)
                cards.column_dimensions[get_column_letter(i)].width = 22
                changed = True
        if "CheckAll" not in wb.sheetnames:
            self._make_cata_sheet(wb)
            changed = True
        if "Messages" not in wb.sheetnames:
            self._make_messages_sheet(wb)
            changed = True
        if "How to edit" not in wb.sheetnames:
            self._make_help_sheet(wb)
            changed = True
        else:
            hs = wb["How to edit"]
            if not any("MESSAGES sheet" in str(r[0].value or "")
                       for r in hs.iter_rows()):
                for line in HELP_MESSAGE_LINES:
                    hs.append([line])
                changed = True
        sc = wb["Scores"]
        if sc["F1"].value == "Points":          # the points system was removed
            sc["F1"] = "Time (s)"
            for r in range(2, sc.max_row + 1):
                sc[f"F{r}"] = None
            changed = True
        for col, name in (("I", "Mode"), ("J", "Result")):
            if sc[f"{col}1"].value is None:
                sc[f"{col}1"] = name
                sc[f"{col}1"]._style = copy(sc["H1"]._style)
                sc.column_dimensions[col].width = 30 if col == "I" else 12
                changed = True
        if changed:
            try:
                wb.save(self.path)
            except PermissionError:
                pass            # saved with the next game instead

    def _create_sample(self):
        wb = Workbook()
        cards = wb.active
        cards.title = "Cards"
        cards.append(CARD_HEADERS)
        for i, (cat, q, a, wrong) in enumerate(SAMPLE_DECK, start=1):
            cards.append([i, cat, q, a, 0, 0, None] + list(wrong))
        cata = self._make_cata_sheet(wb)
        for i, (cat, q, right, wrong) in enumerate(SAMPLE_CATA, start=1):
            cata.append([i, cat, q, "; ".join(right), "; ".join(wrong), 0, 0])
        self._make_messages_sheet(wb)
        scores = wb.create_sheet("Scores")
        scores.append(SCORE_HEADERS)
        self._style(cards, [6, 16, 48, 28, 12, 14, 12, 22, 22, 22])
        self._style(scores, [20, 16, 16, 11, 9, 9, 13, 11, 30, 12])
        self._make_help_sheet(wb)
        self._write_formulas(wb)
        wb.save(self.path)

    @staticmethod
    def _style(ws, widths):
        head_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in ws[1]:
            cell.font = Font(name="Arial", bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center")
        for col, w in zip("ABCDEFGHIJ", widths):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"

    @staticmethod
    def _write_formulas(wb):
        """Accuracy columns are live Excel formulas, not stored numbers."""
        cards, scores = wb["Cards"], wb["Scores"]
        for r in range(2, cards.max_row + 1):
            cards[f"G{r}"] = f'=IF(E{r}=0,"",F{r}/E{r})'
            cards[f"G{r}"].number_format = "0%"
        if "CheckAll" in wb.sheetnames:
            cata = wb["CheckAll"]
            for r in range(2, cata.max_row + 1):
                cata[f"H{r}"] = f'=IF(F{r}=0,"",G{r}/F{r})'
                cata[f"H{r}"].number_format = "0%"
        for r in range(2, scores.max_row + 1):
            scores[f"H{r}"] = f'=IF(D{r}=0,"",E{r}/D{r})'
            scores[f"H{r}"].number_format = "0%"

    # ---- cards -----------------------------------------------------------
    @staticmethod
    def _split(value):
        """'a; b; c' (or one per line) -> ['a', 'b', 'c']"""
        return [p.strip() for p in re.split(r"[;\n]", str(value or ""))
                if p.strip()]

    def get_cards(self):
        """Every non-empty row of the Cards sheet (Identification / MC / T-F)."""
        out, ws = [], self.wb["Cards"]
        for r in range(2, ws.max_row + 1):
            q, a = ws[f"C{r}"].value, ws[f"D{r}"].value
            if q in (None, "") or a in (None, ""):
                continue
            out.append({
                "type": "card",
                "row": r,
                "category": str(ws[f"B{r}"].value or "General"),
                "question": str(q),
                "answer": str(a),
                "wrong": [str(ws[f"{col}{r}"].value).strip()
                          for col in "HIJ"
                          if ws[f"{col}{r}"].value not in (None, "")
                          and str(ws[f"{col}{r}"].value).strip()],
                "seen": int(ws[f"E{r}"].value or 0),
                "correct": int(ws[f"F{r}"].value or 0),
            })
        return out

    def get_cata(self):
        """Every usable row of the CheckAll sheet (Check all that apply)."""
        out, ws = [], self.wb["CheckAll"]
        for r in range(2, ws.max_row + 1):
            q = ws[f"C{r}"].value
            right, wrong = self._split(ws[f"D{r}"].value), self._split(ws[f"E{r}"].value)
            if q in (None, "") or not right or len(right) + len(wrong) < 2:
                continue
            out.append({
                "type": "cata",
                "row": r,
                "category": str(ws[f"B{r}"].value or "General"),
                "question": str(q),
                "right": right,
                "wrong": wrong,
                "seen": int(ws[f"F{r}"].value or 0),
                "correct": int(ws[f"G{r}"].value or 0),
            })
        return out

    def get_messages(self):
        """Texts + passing score from the Messages sheet (defaults if blank)."""
        out = {k: list(v) for k, v in DEFAULT_MESSAGES.items()}
        out["passing"] = float(DEFAULT_PASSING)
        if "Messages" not in self.wb.sheetnames:
            return out
        ws = self.wb["Messages"]
        for key, col in MESSAGE_COLUMNS.items():
            vals = [str(ws[f"{col}{r}"].value).strip()
                    for r in range(2, ws.max_row + 1)
                    if ws[f"{col}{r}"].value not in (None, "")
                    and str(ws[f"{col}{r}"].value).strip()]
            if vals:
                out[key] = vals
        for r in range(2, ws.max_row + 1):
            if "passing" in str(ws[f"G{r}"].value or "").lower():
                raw = ws[f"H{r}"].value
                try:
                    x = float(str(raw).strip().rstrip("%"))
                    out["passing"] = max(0.0, min(100.0,
                                                  x * 100 if 0 < x <= 1 else x))
                except (TypeError, ValueError):
                    pass
                break
        return out

    def record_result(self, row, was_correct, kind="card"):
        if kind == "cata":
            ws, seen, ok = self.wb["CheckAll"], "F", "G"
        else:
            ws, seen, ok = self.wb["Cards"], "E", "F"
        ws[f"{seen}{row}"] = int(ws[f"{seen}{row}"].value or 0) + 1
        if was_correct:
            ws[f"{ok}{row}"] = int(ws[f"{ok}{row}"].value or 0) + 1

    def add_card(self, category, question, answer, wrong=()):
        ws = self.wb["Cards"]
        r = ws.max_row + 1
        ids = [c.value for c in ws["A"][1:] if isinstance(c.value, int)]
        wrong = list(wrong)[:3] + [None] * (3 - len(list(wrong)[:3]))
        ws.append([max(ids, default=0) + 1, category, question, answer, 0, 0,
                   None] + wrong)
        ws[f"G{r}"] = f'=IF(E{r}=0,"",F{r}/E{r})'
        ws[f"G{r}"].number_format = "0%"

    # ---- scores ----------------------------------------------------------
    def add_score(self, player, category, questions, correct, seconds, streak,
                  mode="Identification", result=""):
        ws = self.wb["Scores"]
        for col, name in (("I", "Mode"), ("J", "Result")):
            if ws[f"{col}1"].value is None:    # older workbooks: add the column
                ws[f"{col}1"] = name
                ws[f"{col}1"]._style = copy(ws["H1"]._style)
                ws.column_dimensions[col].width = 30 if col == "I" else 12
        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M"), player, category,
                   questions, correct, round(seconds, 1), streak])
        r = ws.max_row
        ws[f"I{r}"] = mode
        ws[f"J{r}"] = result
        while ws.max_row - 1 > MAX_SCORES:      # erase the oldest record(s)
            ws.delete_rows(2)
        for r in range(2, ws.max_row + 1):      # formulas must follow their row
            ws[f"H{r}"] = f'=IF(D{r}=0,"",E{r}/D{r})'
            ws[f"H{r}"].number_format = "0%"

    def clear_scores(self):
        ws = self.wb["Scores"]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

    def leaderboard(self, limit=MAX_SCORES):
        """Best first: highest % correct, then most correct, then fastest.
        Rows: (player, correct, questions, seconds, mode, category, date, result)"""
        rows, ws = [], self.wb["Scores"]
        for r in range(2, ws.max_row + 1):
            if ws[f"B{r}"].value is None:
                continue
            t = ws[f"F{r}"].value
            rows.append((ws[f"B{r}"].value, int(ws[f"E{r}"].value or 0),
                         int(ws[f"D{r}"].value or 0),
                         float(t) if isinstance(t, (int, float)) else None,
                         ws[f"I{r}"].value or "Identification",
                         ws[f"C{r}"].value, ws[f"A{r}"].value,
                         ws[f"J{r}"].value or ""))
        rows.sort(key=lambda x: (-(x[1] / x[2] if x[2] else 0), -x[1],
                                 x[3] if x[3] is not None else 1e9))
        return rows[:limit]

    def save(self):
        while True:
            try:
                self.wb.save(self.path)
                return
            except PermissionError:
                input(f"\n! Can't write '{self.path}'. Close it in Excel, "
                      "then press Enter to retry...")


# --------------------------------------------------------------------------
# Game logic
# --------------------------------------------------------------------------
def normalize(text):
    text = re.sub(r"[^\w\s.]", "", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def is_correct(given, expected):
    g, e = normalize(given), normalize(expected)
    if not g:
        return False
    if g == e:
        return True
    return difflib.SequenceMatcher(None, g, e).ratio() >= FUZZY_THRESHOLD


def make_hint(answer):
    """First letter of each word, rest masked: 'George Washington' -> 'G_____ W________'"""
    return " ".join(w[0] + "_" * (len(w) - 1) for w in answer.split())


def pick_questions(cards, n):
    """Weighted draw: cards you miss often or haven't seen show up more."""
    pool = cards[:]
    weights = []
    for c in pool:
        acc = c["correct"] / c["seen"] if c["seen"] else 0.5
        weights.append(1.5 - acc)  # 0.5 .. 1.5
    chosen = []
    n = min(n, len(pool))
    for _ in range(n):
        idx = random.choices(range(len(pool)), weights=weights)[0]
        chosen.append(pool.pop(idx))
        weights.pop(idx)
    return chosen


def play(store, player):
    cards = store.get_cards()
    if not cards:
        print("The deck is empty. Add cards first.")
        return

    cats = sorted({c["category"] for c in cards})
    print("\nCategories: " + ", ".join(cats))
    choice = input("Category (Enter = all): ").strip()
    if choice:
        match = [c for c in cats if c.lower() == choice.lower()]
        if not match:
            print("No such category.")
            return
        cards = [c for c in cards if c["category"] == match[0]]
        cat_label = match[0]
    else:
        cat_label = "All"

    try:
        n = int(input(f"How many cards? (max {len(cards)}, default 10): ") or 10)
    except ValueError:
        n = 10
    deck = pick_questions(cards, max(1, n))

    correct = streak = best_streak = 0
    print("\nType your answer. Commands: 'h' = hint, 's' = skip, "
          "'q' = quit round.\n")

    played = 0
    started = time.time()
    for i, card in enumerate(deck, start=1):
        print(f"--- Card {i}/{len(deck)}  [{card['category']}]  "
              f"Correct: {correct}/{len(deck)} ---")
        print(f"Q: {card['question']}")
        hinted = False
        while True:
            ans = input("A: ").strip()
            if ans.lower() == "h" and not hinted:
                print("   Hint:", make_hint(card["answer"]))
                hinted = True
                continue
            break

        if ans.lower() == "q":
            break
        played += 1

        ok = ans.lower() != "s" and is_correct(ans, card["answer"])
        if ok:
            streak += 1
            best_streak = max(best_streak, streak)
            correct += 1
            print("   Correct!\n")
        else:
            print(f"   Answer: {card['answer']}")
            if ans.lower() != "s" and input(
                    "   Count as correct anyway? (y/N): ").lower() == "y":
                ok = True
                streak += 1
                best_streak = max(best_streak, streak)
                correct += 1
                print("   OK, counted as correct\n")
            else:
                streak = 0
                print("   Streak reset.\n")
        store.record_result(card["row"], ok)

    if played == 0:
        print("No cards answered - nothing saved.")
        return

    total = time.time() - started
    pct = round(100 * correct / played)
    msgs = store.get_messages()
    passed = pct >= msgs["passing"]
    text = random.choice(msgs["pass" if passed else "fail"])
    print("=" * 44)
    print(" " + text.replace("{name}", player).replace("{percent}", str(pct)))
    print(f" Score: {correct}/{played} correct ({pct}%)   "
          f"Time: {int(total // 60)}:{int(total % 60):02d}")
    print(f" Best streak: {best_streak}")
    print("=" * 44)
    store.add_score(player, cat_label, played, correct, total, best_streak,
                    result="Passed" if passed else "Failed")
    store.save()
    print("Results saved to Excel.")


def show_leaderboard(store):
    rows = store.leaderboard()
    if not rows:
        print("\nNo scores yet. Play a game!")
        return
    print("\n  #  Player           Score    Time   Category    Date")
    print("  " + "-" * 62)
    for i, (p, c, q, t, mode, cat, d, _) in enumerate(rows, start=1):
        tm = "-" if t is None else f"{int(t // 60)}:{int(t % 60):02d}"
        print(f" {i:>2}  {str(p)[:15]:<15} {c:>3}/{q:<3}  {tm:>5}   "
              f"{str(cat)[:10]:<10}  {d}")


def show_stats(store):
    cards = [c for c in store.get_cards() if c["seen"]]
    if not cards:
        print("\nNo card stats yet.")
        return
    cards.sort(key=lambda c: c["correct"] / c["seen"])
    print("\nHardest cards (lowest accuracy):")
    for c in cards[:5]:
        print(f"  {c['correct']}/{c['seen']}  {c['question']}  ->  {c['answer']}")


def add_card_flow(store):
    cat = input("Category: ").strip() or "General"
    q = input("Question: ").strip()
    a = input("Answer: ").strip()
    if not q or not a:
        print("Question and answer are required.")
        return
    store.add_card(cat, q, a)
    store.save()
    print("Card added.")


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else str(app_dir() / DEFAULT_FILE)
    new_file = not Path(path).exists()
    store = Store(path)
    if new_file:
        print(f"Created sample deck: {path}")

    print("\n=== FLASH CARD GAME ===")
    player = input("Your name: ").strip() or "Player"

    while True:
        print(f"\n[{player}]  1) Play   2) Leaderboard   3) Add card   "
              "4) Hardest cards   5) Quit")
        cmd = input("> ").strip()
        if cmd == "1":
            play(store, player)
        elif cmd == "2":
            show_leaderboard(store)
        elif cmd == "3":
            add_card_flow(store)
        elif cmd == "4":
            show_stats(store)
        elif cmd == "5":
            print("Bye!")
            break
        else:
            print("Pick 1-5.")


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\nBye!")
