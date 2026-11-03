# 와규 재고 체크

팀원이 Teams 링크로 와규 개수를 입력하면 구글 시트에 서버 시각(브리즈번)과 대략 kg이 쌓이고, 매니저가 있어야 할 개수를 넣으면 차이와 CCTV 확인 구간이 나오는 Google Apps Script 웹 앱이에요. 지금은 와규 한 품목만 세는 화면이에요.

| 파일 | 붙여넣을 곳 |
| --- | --- |
| `apps-script/Code.gs` | Apps Script의 `Code.gs` |
| `apps-script/Index.html` | Apps Script에 새로 만든 HTML 파일 `Index` |

설치 순서, 폰 화면 미리보기, 복사 버튼은 `index.html`(설치 안내 페이지)에 있어요. 웹 주소(`…/meat-stock/`)로 열어야 미리보기와 복사가 동작해요.

짧게: 새 구글 시트 → 확장 프로그램 → Apps Script → 두 파일 붙여넣기 → `setup` 한 번 실행 → 배포 → 웹 앱(실행: 나, 액세스: 모든 사용자) → Settings 탭에서 Manager PIN 바꾸기.

Settings 탭

| 줄 | 기본값 | 뜻 |
| --- | --- | --- |
| Manager PIN | 1234 | 매니저 화면 PIN. 10번 틀리면 15분 잠김 |
| Avg kg per unit | 0.1425 | 와규 한 개 평균 무게(kg) |
| Price per kg | 비어 있음 | 넣으면 차이가 $로도 나와요 |
| Page title | Wagyu Stock Check | 화면 맨 위 제목 |
