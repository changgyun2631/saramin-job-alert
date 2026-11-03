"""API 키 없이 가짜 공고로 필터·점수·재알림·엑셀을 점검한다. 실행: python test_alert.py"""
import os
import tempfile
import time

import alert


def job(jid, title, company="(주)테스트", loc="부산 &gt; 사상구", kw="", lo=0, hi=0, code=2, days=30):
    return {
        "id": jid, "url": f"https://www.saramin.co.kr/zf_user/jobs/relay/view?rec_idx={jid}",
        "company": {"detail": {"name": company}},
        "position": {"title": title, "location": {"name": loc}, "industry": {"name": "금속·주조"},
                     "job-code": {"name": "회계"},
                     "experience-level": {"code": code, "min": lo, "max": hi, "name": f"경력 {lo}년↑"}},
        "keyword": kw, "salary": {"name": ""},  # 웹 수집처럼 연봉 없음 → 상세 페이지 값 사용
        "expiration-timestamp": str(int(time.time()) + days * 86400), "close-type": {"code": "1", "name": "접수마감일"},
    }


JOBS = {
    "1": job(1, "원가회계·연결결산 과장", kw="결산,ERP", lo=10),
    "2": job(2, "경리 사무원", hi=3),
    "3": job(3, "회계 담당", company="OO세무회계"),       # 회사명 제외
    "4": job(4, "회계 파견직 모집"),                       # 파견 제외
    "5": job(5, "재무회계 차장", loc="경남 &gt; 거제시"),  # 지역 제외
    "6": job(6, "회계 대리", loc="경남 &gt; 김해시", days=2),
    "7": job(7, "회계 담당", lo=2),                         # 경력2년↑ 저연차
    "8": job(8, "회계 감사 보조", company="회계법인공감"),  # 회계법인 제외
}


ITEM = """<div class="item_recruit" value="99"><h2 class="job_tit"><a title="원가회계 과장"></a></h2>
<div class="job_date"><span class="date">~ 01/08(목)</span></div>
<div class="job_condition"><span><a>경남</a> <a>김해시</a></span> <span>경력 3~5년</span> <span>정규직</span></div>
<div class="job_sector"><a>원가회계</a>, <a>결산</a></div><strong class="corp_name"><a> (주)가나 </a></strong></div>"""


def demo():
    it = alert.BeautifulSoup(ITEM, "html.parser").div
    j = alert.parse_item(it, today=alert.datetime(2026, 12, 20, tzinfo=alert.KST))
    assert j["company"]["detail"]["name"] == "(주)가나" and j["position"]["location"]["name"] == "경남 김해시", j
    assert (j["position"]["experience-level"]["min"], j["position"]["experience-level"]["max"]) == (3, 5)
    assert alert.make_row(j, 0, [])["deadline"] == "2027-01-08"  # 해 넘김

    os.chdir(tempfile.mkdtemp())
    os.environ["SARAMIN_KEY"] = "fake"
    os.environ.pop("TELEGRAM_TOKEN", None)
    alert.fetch = lambda key, **extra: {} if extra else {k: v for k, v in JOBS.items() if k != "6"}

    alert.fetch_detail = lambda jid: {"body_words": "원가" if jid == "6" else "", "page_salary": "연봉 5,000만원"}
    alert.main()
    state = alert.json.loads(alert.STATE.read_text("utf-8"))
    assert set(state) == {"1", "2", "7"}, state.keys()
    assert state["7"]["score"] == 2 - 2 and state["7"]["salary"] == "연봉 5,000만원", state["7"]
    assert state["1"]["score"] == 3 + 3 + 2 + 2 + 1 + 2 + 2, state["1"]  # 원가 연결 결산 조선·기계 ERP 직급 경력10년↑
    assert state["2"]["score"] == 2 - 2 - 2, state["2"]                  # 조선·기계(업종) 경리 저연차

    # 둘째 날: 6번 신규(D-2 → 신규로만 알림), 2번 사라짐 → 마감
    alert.fetch = lambda key, **extra: {} if extra else {k: v for k, v in JOBS.items() if k != "2"}
    sent = []
    alert.send = sent.append
    alert.main()
    state = alert.json.loads(alert.STATE.read_text("utf-8"))
    assert state["2"]["status"] == "마감" and state["6"]["status"] == "진행중"
    assert state["6"]["score"] == 2 + 3, state["6"]  # 조선·기계 + 본문 원가
    assert "신규 1건 · 마감임박 0건" in sent[0], sent[0]

    # 셋째 날: 1번 마감이 D-2로 당겨짐 → D-3 재알림 1회, 넷째 날엔 다시 안 옴
    JOBS["1"]["expiration-timestamp"] = str(int(time.time()) + 2 * 86400)
    alert.main()
    alert.main()
    assert "마감임박 1건" in sent[1] and "오늘 신규 없음" in sent[2], sent[1:]
    assert alert.XLSX.exists() and "원가회계·연결결산 과장" in alert.HTML.read_text("utf-8")
    print("OK", os.getcwd())


if __name__ == "__main__":
    demo()
