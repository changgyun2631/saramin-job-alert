"""사람인 회계 공고를 모아 내 이력 기준으로 점수를 매기고 텔레그램으로 보낸다."""
import html
import json
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font

KST = timezone(timedelta(hours=9))
API = "https://oapi.saramin.co.kr/job-search"
STATE = Path("jobs.json")
XLSX = Path("jobs.xlsx")

REGIONS = ("106000", "110000")  # 부산 전체, 경남 전체 → CITIES로 좁힘
CITIES = ("부산", "창원", "마산", "진해", "김해", "양산")
OFFICE_NAMES = ("세무사", "세무회계", "세무법인", "회계사무소", "세무사무소")  # 회사명에 있으면 제외
EXCLUDE_WORDS = ("세무사무소", "회계사무소", "세무법인", "기장대행", "기장 대행", "파견")
RULES = [  # (찾을 단어, 점수, 근거 표시)
    (("원가",), 3, "원가"),
    (("연결",), 3, "연결"),
    (("결산",), 2, "결산"),
    (("제조",), 2, "제조업"),
    (("조선", "선박", "기자재", "기계", "금속", "주조", "주물"), 2, "조선·기계"),
    (("외부감사", "회계감사"), 1, "외부감사"),
    (("부가세", "부가가치세", "급여"), 1, "부가세·급여"),
    (("erp", "더존", "sap", "이카운트", "ecount"), 1, "ERP"),
    (("과장", "차장", "팀장"), 2, "직급"),
    (("경리",), -2, "경리"),
]
LISTED_WORDS = ("상장", "코스닥", "코스피", "ipo")
HEADERS = ["상태", "점수", "점수 근거", "회사명", "공고 제목", "지역", "경력", "연봉", "마감일", "최초 발견일", "링크"]


def fetch(key, **extra):
    """회계·세무·재무 / 정규직 / 헤드헌팅·파견 제외 공고를 지역별로 전부 가져온다."""
    jobs = {}
    for loc in REGIONS:
        start = 0
        while True:
            r = requests.get(API, headers={"Accept": "application/json"}, timeout=30, params={
                "access-key": key, "loc_cd": loc, "job_mid_cd": "3", "job_type": "1",
                "sr": "directhire", "count": 110, "start": start, **extra})
            r.raise_for_status()
            data = r.json()["jobs"]
            for j in data.get("job", []):
                jobs[str(j["id"])] = j
            start += 1
            if start * 110 >= int(data["total"]):
                break
    return jobs


def text_of(j):
    p = j["position"]
    return " ".join([p["title"], j.get("keyword", ""), p.get("industry", {}).get("name", ""),
                     p.get("job-code", {}).get("name", "")]).lower()


def in_area(j):
    return any(c in j["position"]["location"]["name"] for c in CITIES)


def excluded(j):
    company = j["company"]["detail"]["name"]
    return any(w in company for w in OFFICE_NAMES) or any(w in text_of(j) for w in EXCLUDE_WORDS)


def score(j, listed):
    text = text_of(j)
    pts, why = 0, []
    for words, p, label in RULES:
        if any(w in text for w in words):
            pts += p
            why.append(f"{label}{p:+d}")
    if str(j["id"]) in listed or any(w in text for w in LISTED_WORDS):
        pts += 2
        why.append("상장+2")
    exp = j["position"]["experience-level"]
    lo, hi = int(exp.get("min") or 0), int(exp.get("max") or 0)
    if lo >= 10:
        pts += 2
        why.append("경력10년↑+2")
    elif int(exp.get("code") or 0) == 1 or 0 < hi <= 3:
        pts -= 2
        why.append("저연차-2")
    return pts, why


def make_row(j, pts, why):
    p = j["position"]
    close = j.get("close-type", {})
    ts = int(j.get("expiration-timestamp") or 0)
    on_date = str(close.get("code")) == "1" and ts
    return {
        "score": pts, "why": ", ".join(why), "company": j["company"]["detail"]["name"],
        "title": p["title"], "location": html.unescape(p["location"]["name"]),
        "experience": p["experience-level"].get("name", ""), "salary": j.get("salary", {}).get("name", ""),
        "deadline": datetime.fromtimestamp(ts, KST).strftime("%Y-%m-%d") if on_date else close.get("name", ""),
        "deadline_ts": ts if on_date else 0, "url": j["url"],
    }


def write_xlsx(state):
    wb = Workbook()
    ws = wb.active
    ws.title = "공고"
    ws.append(HEADERS)
    for r in sorted(state.values(), key=lambda r: (r["status"] != "진행중", -r["score"])):
        ws.append([r["status"], r["score"], r["why"], r["company"], r["title"], r["location"],
                   r["experience"], r["salary"], r["deadline"], r["first_seen"], r["url"]])
        ws.cell(ws.max_row, 11).hyperlink = r["url"]
    for col, width in zip("ABCDEFGHIJK", (7, 6, 30, 22, 50, 22, 14, 18, 12, 12, 40)):
        ws.column_dimensions[col].width = width
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(XLSX)


def line(r):
    e = html.escape
    return f'[{r["score"]}] {e(r["company"])} | <a href="{e(r["url"])}">{e(r["title"][:40])}</a> | ~{e(r["deadline"])}'


def message(new, soon, now):
    lines = [f"<b>사람인 회계 공고 {now:%m/%d}</b>", f"신규 {len(new)}건 · 마감임박 {len(soon)}건"]
    if not new and not soon:
        lines.append("오늘 신규 없음")
    for title, items in (("신규 상위 10", new), ("마감 D-3", soon)):
        if items:
            lines += ["", f"<b>{title}</b>"] + [line(r) for r in sorted(items, key=lambda r: -r["score"])[:10]]
    return "\n".join(lines)


def send(text):
    token, chat = os.environ.get("TELEGRAM_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if not token:  # 로컬 점검용: 보내지 않고 출력만
        print(text)
        return
    base = f"https://api.telegram.org/bot{token}"
    requests.post(f"{base}/sendMessage", timeout=30, json={
        "chat_id": chat, "text": text, "parse_mode": "HTML", "disable_web_page_preview": True}).raise_for_status()
    with XLSX.open("rb") as f:
        requests.post(f"{base}/sendDocument", timeout=60, data={"chat_id": chat},
                      files={"document": (XLSX.name, f)}).raise_for_status()


def main():
    key = os.environ["SARAMIN_KEY"]
    now = datetime.now(KST)
    state = json.loads(STATE.read_text("utf-8")) if STATE.exists() else {}
    jobs = fetch(key)
    listed = set(fetch(key, stock="kosdaq")) | set(fetch(key, stock="kospi"))
    new, soon, alive = [], [], set()
    for jid, j in jobs.items():
        if not in_area(j) or excluded(j):
            continue
        alive.add(jid)
        rec = state.get(jid)
        is_new = rec is None
        rec = (rec or {"first_seen": now.strftime("%Y-%m-%d"), "reminded": False})
        rec.update(make_row(j, *score(j, listed)), status="진행중")
        state[jid] = rec
        near = rec["deadline_ts"] and 0 <= rec["deadline_ts"] - now.timestamp() <= 3 * 86400
        if is_new:
            new.append(rec)
            rec["reminded"] = bool(near)  # 신규로 이미 알렸으니 D-3 재알림 생략
        elif near and not rec["reminded"]:
            soon.append(rec)
            rec["reminded"] = True
    for jid, rec in state.items():
        if jid not in alive:
            rec["status"] = "마감"
    STATE.write_text(json.dumps(state, ensure_ascii=False, indent=1), "utf-8")
    write_xlsx(state)
    send(message(new, soon, now))


if __name__ == "__main__":
    main()
