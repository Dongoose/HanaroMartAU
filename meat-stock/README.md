# 정육 재고 체크

팀원이 Teams 링크로 품목마다 매장 내 보유 갯수와 시스템 상의 재고를 입력하면, 구글 시트가 서버 시각(브리즈번)과 함께 기록하고 둘을 비교해요. 한 개쯤 차이 나면 빨간불이 들어오고, 품목별 직전 카운트와 CCTV 확인 구간도 붙어요. Google Apps Script 웹 앱이에요.

지금 세는 품목 (시트의 Products 탭에서 바꿔요)

| 품목 | 매장에서 | 시스템 | 빨간불 |
| --- | --- | --- | --- |
| 와규 (Wagyu) | 갯수 | kg (갯수 × 0.1455 kg으로 비교) | ±0.10 kg 이상 |
| 냉장 삼겹살 (Pork belly) | 갯수 | 갯수 | 1개 이상 |
| 와규 국거리 (Wagyu soup cuts) | 현재 매장 재고 없음 · 나중에 확인 (Check later) | kg | ±0.10 kg 이상 |

| 파일 | 붙여넣을 곳 |
| --- | --- |
| `apps-script/Code.gs` | Apps Script의 `Code.gs` |
| `apps-script/Index.html` | Apps Script에 새로 만든 HTML 파일 `Index` |

설치 순서, 폰 화면 미리보기, 복사 버튼은 `index.html`(설치 안내 페이지)에 있어요. 웹 주소(`…/meat-stock/`)로 열어야 미리보기와 복사가 동작해요.

짧게: 새 구글 시트 → 확장 프로그램 → Apps Script → 두 파일 붙여넣기 → `setup` 한 번 실행 → 배포 → 웹 앱(실행: 나, 액세스: 모든 사용자) → Settings 탭에서 Manager PIN 바꾸기.

코드를 바꿀 때는 두 파일을 다시 붙여넣고 **배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포**. 링크는 그대로예요.

화면
- 첫 화면: 품목 카드마다 ① 매장 내 보유 갯수 입력, ② 시스템 상의 재고 입력(품목 단위). 센 품목만 넣고 한 번에 저장해요. "Check later" 품목은 접혀 있고 "지금 입력하기"로 열어요. 맞는지 여부는 팀원 화면에 안 보여줘요.
- 맨 위 초록 바: ‹ › 이전·다음 화면, 로고·제목이나 집 버튼을 누르면 첫 화면.
- 매니저 화면: 품목별 최근 상태, 품목 칩으로 골라 보기, 한 번에 저장한 기록은 한 카드로 묶여요. 맨 아래 **Reset all data**는 "Are you sure…?"에서 한 번 더 확인해야 지워져요. 실수로 지웠다면 시트의 파일 → 버전 기록에서 되돌릴 수 있어요.

시트
- Log 탭: 품목마다 한 줄. 숫자 옆에 단위(kg, pcs)가 붙어 보이고, 차이가 그 줄의 `Red when off by` 이상이면(부족·초과 모두) 차이 칸이 빨갛게 표시돼요.
- Products 탭: Product · 한글 이름 · System unit(kg / pcs) · Avg kg per piece · Red when off by · Price per system unit · Status(Active / Check later / Off) · Note for staff. 줄을 더하면 품목이 늘고, Off로 바꾸면 화면에서 빠져요. 재배포는 필요 없어요.
- Settings 탭: Manager PIN(기본 1234, 10번 틀리면 15분 잠김), Page title(기본 Meat Stock Check).

와규만 세던 버전에서 바꾸면, 새 버전으로 배포한 뒤 페이지를 처음 열 때 Products 탭이 생기고(와규 무게·빨간불 기준·가격은 Settings에서 옮겨져요) 기존 Log를 그대로 이어서 써요. 매니저가 재고를 넣던 첫 버전의 Log는 `Log (old)`로 보관돼요.
