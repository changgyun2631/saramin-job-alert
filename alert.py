"""사람인 회계 공고를 모아 내 이력 기준으로 점수를 매기고 텔레그램으로 보낸다."""
import html
import json
import os
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

KST = timezone(timedelta(hours=9))
API = "https://oapi.saramin.co.kr/job-search"
WEB = "https://www.saramin.co.kr/zf_user/search/recruit"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"}
VIEW = "https://www.saramin.co.kr/zf_user/jobs/relay/view"
DETAIL = "https://www.saramin.co.kr/zf_user/jobs/relay/view-detail"
DETAIL_MAX = 60  # 한 번에 본문을 읽을 공고 수 (신규 먼저, 남으면 기존 공고 순차 보충)
WEB_LOCS ="106000,110040,110100,110150,110053,110055,110057,110059,110130"  # 부산, 김해, 양산, 창원(구 포함)
STATE = Path("jobs.json")
XLSX = Path("jobs.xlsx")

REGIONS = ("106000", "110000")  # 부산 전체, 경남 전체 → CITIES로 좁힘
CITIES = ("부산", "창원", "마산", "진해", "김해", "양산")
OFFICE_NAMES = ("세무사", "세무회계", "세무법인", "회계법인", "회계사무소", "세무사무소")  # 회사명에 있으면 제외
EXCLUDE_WORDS = ("세무사무소", "회계사무소", "세무법인", "기장대행", "기장 대행", "파견")
RULES = [  # (찾을 단어, 점수, 근거 표시)
    (("원가",), 3, "원가"),
    (("연결",), 3, "연결"),
    (("결산",), 2, "결산"),
    (("제조",), 2, "제조업"),
    (("조선", "선박", "기자재", "기계", "금속", "주조", "주물"), 2, "조선·기계"),
    (("외부감사", "회계감사"), 1, "외부감사"),
    (("부가세", "부가가치세", "급여업무", "급여 업무", "급여관리", "급여 관리", "급여계산", "급여 계산", "4대보험"), 1, "부가세·급여"),
    (("erp", "더존", "sap", "이카운트", "ecount"), 1, "ERP"),
    (("과장", "차장", "팀장"), 2, "직급"),
    (("경리",), -2, "경리"),
]
LISTED_WORDS = ("상장", "코스닥", "코스피", "ipo")
BODY_WORDS = [w for words, p, _ in RULES if p > 0 for w in words] + list(LISTED_WORDS)  # 본문에선 가산 단어만 본다
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


def fetch_web():
    """API 키가 없을 때: 사람인 검색 화면에서 같은 조건(회계·세무·재무, 정규직) 공고를 읽는다."""
    # ponytail: 화면 구조가 바뀌면 깨짐 → API 키 받으면 SARAMIN_KEY 등록으로 자동 전환
    jobs = {}
    for page in range(1, 21):
        r = requests.get(WEB, headers=UA, timeout=30, params={
            "searchType": "search", "searchword": "", "loc_cd": WEB_LOCS, "cat_mcls": "3",
            "job_type": "1", "recruitPageCount": "100", "recruitPage": page})
        r.raise_for_status()
        items = BeautifulSoup(r.text, "html.parser").select("div.item_recruit")
        for it in items:
            j = parse_item(it)
            jobs[j["id"]] = j
        if len(items) < 100:
            break
        time.sleep(2)
    if not jobs:
        raise RuntimeError("사람인 검색 결과 0건: 화면 구조가 바뀌었거나 접속이 막혔을 수 있음")
    return jobs


def parse_item(it, today=None):
    """검색 결과 한 칸을 API 응답과 같은 모양으로 바꾼다."""
    today = today or datetime.now(KST)
    rid = it["value"]
    cond = [s.get_text(" ", strip=True) for s in it.select("div.job_condition > span")] + ["", ""]
    loc, exp = cond[0], cond[1]
    nums = [int(n) for n in re.findall(r"\d+", exp)]
    lo = nums[0] if nums else 0
    hi = nums[1] if "~" in exp and len(nums) > 1 else 0
    date = it.select_one("div.job_date .date").get_text(strip=True)
    m = re.search(r"(\d+)/(\d+)", date)
    due = None
    if m:
        due = today.replace(month=int(m[1]), day=int(m[2]), hour=23, minute=59)
        if due < today - timedelta(days=1):  # 12월에 본 1월 마감 → 내년
            due = due.replace(year=due.year + 1)
    elif "오늘" in date or "내일" in date:
        due = today.replace(hour=23, minute=59) + timedelta(days="내일" in date)
    sector = it.select_one("div.job_sector")
    return {
        "id": rid, "url": f"https://www.saramin.co.kr/zf_user/jobs/relay/view?rec_idx={rid}",
        "company": {"detail": {"name": it.select_one(".corp_name").get_text(strip=True)}},
        "position": {"title": it.select_one(".job_tit a")["title"], "location": {"name": loc},
                     "experience-level": {"code": 1 if exp == "신입" else 2, "min": lo, "max": hi, "name": exp}},
        "keyword": ", ".join(a.get_text(strip=True) for a in sector.select("a")) if sector else "",
        "expiration-timestamp": int(due.timestamp()) if due else 0,
        "close-type": {"code": "1" if due else "2", "name": date},
    }


def fetch_detail(jid):
    """공고 본문에서 가산 단어를, 공고 페이지 요약에서 연봉을 뽑는다. 본문 전체는 저장하지 않는다."""
    body = requests.get(DETAIL, headers=UA, timeout=30, params={"rec_idx": jid, "rec_seq": 0})
    page = requests.get(VIEW, headers=UA, timeout=30, params={"rec_idx": jid})
    body.raise_for_status()
    text = BeautifulSoup(body.text, "html.parser").get_text(" ").lower()
    m = re.search(r'연봉:(.*?)(?:, [가-힣]+:|")|(면접 후 결정|회사내규에 따름)', page.text)
    return {"body_words": " ".join(w for w in BODY_WORDS if w in text), "page_salary": (m[1] or m[2]).strip() if m else ""}


def text_of(j):
    p = j["position"]
    return " ".join([p["title"], j.get("keyword", ""), p.get("industry", {}).get("name", ""),
                     p.get("job-code", {}).get("name", "")]).lower()


def in_area(j):
    return any(c in j["position"]["location"]["name"] for c in CITIES)


def excluded(j):
    company = j["company"]["detail"]["name"]
    return any(w in company for w in OFFICE_NAMES) or any(w in text_of(j) for w in EXCLUDE_WORDS)


def score(j, listed, body_words=""):
    text = text_of(j) + " " + body_words
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
    elif int(exp.get("code") or 0) == 1 or 0 < hi <= 3 or 0 < lo <= 2:  # 신입, ~3년, 1·2년↑
        pts -= 2
        why.append("저연차-2")
    return pts, why


def make_row(j, pts, why, page_salary=""):
    p = j["position"]
    close = j.get("close-type", {})
    ts = int(j.get("expiration-timestamp") or 0)
    on_date = str(close.get("code")) == "1" and ts
    return {
        "score": pts, "why": ", ".join(why), "company": j["company"]["detail"]["name"],
        "title": p["title"], "location": html.unescape(p["location"]["name"]),
        "experience": p["experience-level"].get("name", ""), "salary": j.get("salary", {}).get("name", "") or page_salary,
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


def send(text, attach=True):
    token, chat = (os.environ.get(k, "").strip() for k in ("TELEGRAM_TOKEN", "TELEGRAM_CHAT_ID"))  # 붙여넣기 공백·줄바꿈 제거
    if not token:  # 로컬 점검용: 보내지 않고 출력만
        print(text)
        return
    base = f"https://api.telegram.org/bot{token}"
    requests.post(f"{base}/sendMessage", timeout=30, json={
        "chat_id": chat, "text": text, "parse_mode": "HTML", "disable_web_page_preview": True}).raise_for_status()
    if not attach:
        return
    with XLSX.open("rb") as f:
        requests.post(f"{base}/sendDocument", timeout=60, data={"chat_id": chat},
                      files={"document": (XLSX.name, f)}).raise_for_status()


def main():
    if os.environ.get("CI") and not os.environ.get("TELEGRAM_TOKEN"):  # 못 보낸 공고를 '본 공고'로 저장하지 않게
        raise SystemExit("TELEGRAM_TOKEN Secret이 등록되지 않았습니다")
    key = os.environ.get("SARAMIN_KEY")
    now = datetime.now(KST)
    state = json.loads(STATE.read_text("utf-8")) if STATE.exists() else {}
    if key:
        jobs = fetch(key)
        listed = set(fetch(key, stock="kosdaq")) | set(fetch(key, stock="kospi"))
    else:  # 검색 화면에는 업종·상장 정보가 없어 단어로만 판단
        jobs, listed = fetch_web(), set()
    seen = set(state)
    alive = {jid: j for jid, j in jobs.items() if in_area(j) and not excluded(j)}
    todo = sorted((jid for jid in alive if "body_words" not in state.get(jid, {})), key=lambda jid: jid in seen)
    for jid in todo[:DETAIL_MAX]:
        try:
            state.setdefault(jid, {}).update(fetch_detail(jid))
        except requests.RequestException:  # 다음 실행에서 다시 시도
            pass
        time.sleep(1)
    new, soon = [], []
    for jid, j in alive.items():
        rec = state.setdefault(jid, {})
        is_new = jid not in seen
        if is_new:
            rec.update(first_seen=now.strftime("%Y-%m-%d"), reminded=False)
        pts, why = score(j, listed, rec.get("body_words", ""))
        rec.update(make_row(j, pts, why, rec.get("page_salary", "")), status="진행중")
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
    try:
        main()
    except Exception as e:
        try:
            send(f"⚠️ 사람인 알림 실패\n{html.escape(f'{type(e).__name__}: {e}')[:500]}", attach=False)
        except Exception:
            pass  # 텔레그램 자체가 문제면 GitHub 실패 메일로 알 수 있음
        raise
