# 사람인 회계 공고 알림

평일 오전 8시, 사람인에서 부산·창원·양산·김해 회계 정규직 공고를 모아 점수순으로 텔레그램(`@cg_job_bot`)에 보냅니다.
메시지에는 신규 상위 10건, 마감 D-3 재알림, 리포트 링크가 들어가고, 누적 엑셀(`jobs.xlsx`)이 함께 첨부됩니다.
리포트: https://changgyun2631.github.io/saramin-job-alert/jobs.html (GitHub Pages, 열람 기록은 브라우저별 저장)

## 설정 순서

API 키가 없으면 사람인 검색 화면에서 같은 조건의 공고를 읽어옵니다. 그래서 텔레그램 Secrets 두 개만 등록하면 바로 동작합니다.

1. **사람인 API 키 신청** (선택. 키를 등록하면 업종과 상장사 정보까지 반영됩니다): https://oapi.saramin.co.kr
   - API 사용할 곳 URL: `https://github.com/changgyun2631/saramin-job-alert`
   - 용도: 개인 구직용 채용공고 조회
2. **GitHub 비공개 저장소** `saramin-job-alert`에 이 폴더를 올립니다.
3. 저장소 **Settings → Secrets and variables → Actions → New repository secret**에서 세 개를 등록합니다.
   | 이름 | 값 |
   |---|---|
   | `SARAMIN_KEY` | 사람인 access-key |
   | `TELEGRAM_TOKEN` | BotFather 새 토큰 |
   | `TELEGRAM_CHAT_ID` | 봇 getUpdates로 확인한 채팅 ID |
   - (선택) 네이버 메일로도 받으려면 `NAVER_USER`(네이버 아이디, @naver.com 제외)와 `NAVER_PASSWORD`(앱 비밀번호)를 추가합니다. 네이버 메일 환경설정 → POP3/SMTP 설정에서 **IMAP/SMTP 사용**을 켜야 합니다.
4. 저장소 **Actions → daily → Run workflow**로 첫 실행을 합니다. 텔레그램 메시지가 오면 완료입니다.

## 조정

- 점수 가중치: `alert.py`의 `RULES`
- 제외 단어: `OFFICE_NAMES`(회사명), `EXCLUDE_WORDS`(제목·키워드)
- 지역: `CITIES`
- 점검: `python test_alert.py` (API 키 없이 가짜 공고로 확인)

## 참고

- API는 공고 **본문을 주지 않습니다.** 그래서 점수는 제목·키워드·업종·직무명만 보고 매깁니다.
- "제조업" 가산점은 업종명에 "제조"가 들어갈 때만 붙습니다.
- 상장사 여부는 API의 코스피·코스닥 필터로 따로 조회해서 판단합니다.
- 엑셀은 매일 새로 만들어집니다. 엑셀에 직접 적은 메모는 다음 날 사라지니, 메모는 엑셀을 복사해서 따로 관리하세요.
- 실행이 실패하면 GitHub가 가입 메일로 알려줍니다.
