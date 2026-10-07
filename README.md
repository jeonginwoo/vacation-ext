# 휴가 캘린더 (ProWorkflow) 크롬 확장

ProWorkflow Power Apps가 불러오는 휴가 데이터를 가로채서 월간 캘린더로 보여줍니다.
새 권한이나 앱 등록 없이, 내 브라우저에 이미 내려온 데이터만 사용합니다.

## 설치

1. 이 폴더를 원하는 위치에 둡니다 (예: `C:\Projects\playground\vacation-ext`).
2. 크롬에서 `chrome://extensions` → 오른쪽 위 **개발자 모드** 켜기
3. **압축해제된 확장 프로그램을 로드합니다** → 이 폴더 선택
4. 퍼즐 아이콘에서 "휴가 캘린더"를 고정해두면 편합니다.

## 사용

1. ProWorkflow Power Apps → 휴가현황 → **휴가자조회**를 엽니다.
2. 확장 아이콘에 숫자 배지가 뜨면 수집된 것입니다.
3. 확장 아이콘을 누르면 캘린더가 새 탭으로 열립니다.

데이터는 Power Apps를 열 때마다 갱신되고, 캘린더 탭이 열려 있으면 자동으로 다시 그려집니다.

## 동작 방식

- `hook.js` — Power Apps iframe(`runtime-app.powerplatform.com`) 안에서 fetch/XHR 응답 중
  SharePoint 커넥터의 리스트 조회 결과만 읽습니다. 휴가(`UseDate`/`UseYear` 포함)와
  공휴일(`HolidayDate` 포함) 리스트의 필요한 필드만 추리고, 다른 리스트(직원 정보 등)는 버립니다.
- `background.js` — 받은 항목을 `리스트ID:ID` 기준으로 합쳐 `chrome.storage.local`에 저장합니다.
  취소 여부(`CancelYN`)처럼 다른 요청에서 오는 필드도 같은 항목에 합쳐집니다.
- `calendar.html/js/css` — 저장된 데이터를 캘린더로 그립니다. 취소된 휴가는 제외합니다.

## 한계

- Power Apps가 실제로 가져온 데이터만 보입니다. 앱이 리스트를 500건 단위로 조회하므로
  리스트가 그보다 크면 일부 기간이 비어 보일 수 있습니다.
- Power Apps 내부 구조가 바뀌면 수집이 멈출 수 있습니다. 배지가 안 뜨면 확인이 필요합니다.
- 수집 데이터는 이 PC의 크롬에만 저장되며, 캘린더의 "수집 데이터 지우기"로 지울 수 있습니다.
