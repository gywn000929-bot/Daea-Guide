# 대아 업무용 VBA · 리본 변경 및 검증 기록

2026-09-29 / 배포 코드 2026-09-23 기준

## 확인한 원본

- 저장소: `gywn000929-bot/Daea-Guide`의 작업 시점 최신 파일
- `vba-reinstall-practice.html`: `0d05b63b47c4f55e6a70d42e527c7e28fb041f50`
- `downloads/daea-macros-24-20260923.zip`: `6dba39aa6708b1672990f38017007e295fd41ef3`
- `downloads/daea-ribbon-24.exportedUI`: `0f4792eb133dad1c03ad35de42caa828928eb559`
- 설명 대조: `vba-special-classroom.html`
- 정보 구조 참고: `parts-catalog.html#start`

원본 ZIP의 XLAM에서 `xl/vbaProject.bin`을 추출해 VBA 소스와 연결 진입점을 확인했습니다. 통합본의 코드와 파일명은 변경하지 않았습니다. HTML에서 직접 내려받는 XLAM은 원본 ZIP의 파일과 바이트 단위로 같습니다. 학습용 V6·V10으로 대체하지 않았습니다.

새 안내 주소는 `vba-work-guide.html`입니다. 기존 `vba-reinstall-practice.html`은 새 안내로 이동하며, 같은 사이트에 저장된 기존 19단계 확인 기록과 메모를 이어받습니다.

## 리본 설계

기존: `매크로` 그룹에 주 버튼 24개 + `보조 도구`에 버튼 1개.
개선: `구매물류` 탭 아래 다음 6개 그룹, 합계 25개.

|순서|그룹|개수|배치 이유|
|---|---|---:|---|
|1|품목·재고|4|품목 수량과 거래처·품명·재고 확인을 먼저 찾도록 배치|
|2|AIR|5|요청표 정리 → 잔량 → 전달표 → 동시 생성 대안 → 인쇄 서식|
|3|선박·집계|4|선박 정리와 그에 필요한 출고·미납·부분합 집계|
|4|연호|6|발주 필요량 → ERP 수량 대조 → 입고표·거래명세·재고표|
|5|KET·파렛트|5|재고 갱신 → 위치 → 박스 목록 → 합산 → 라벨|
|6|설정|1|일상 처리와 구분하여 시트 조회 버튼 추가|

실제 사용 횟수 기록은 제공되지 않았습니다. 위 순서는 업무 공통 작업을 앞에 둔 설계이며, 사용 빈도를 측정했다는 뜻은 아닙니다. 그룹 안 기능은 모두 연속 실행하는 순서가 아닙니다. 예를 들어 AIR 요청표 정리와 AIR·전달표 동시 생성은 입력·중복 처리 방식이 다르므로 목적에 맞는 기능을 선택합니다.

표시 이름, 실행 프로시저, 파일 위치를 구분합니다. `onAction`은 `설치한 XLAM 전체 경로!원래 실행 프로시저` 구조입니다. 표시 이름은 바꾸되 `!` 뒤의 실행 프로시저 25개는 대소문자와 모듈 한정 이름을 포함해 그대로 보존했습니다. 원본의 `imageMso`도 유지했습니다.

## 개인 메뉴를 보존하는 적용 방법

1. 현재 PC의 Excel → 파일 → 옵션 → 리본 사용자 지정 → 가져오기/내보내기 → 모든 사용자 지정 내보내기를 선택합니다.
2. 업무 파일과 기존 추가 기능도 별도 백업합니다. 교체 대상으로 확인한 것만 해제·정리합니다.
3. 새 `매크로 신상.xlam`을 고정 위치에 두고 Excel 추가 기능으로 등록합니다.
4. HTML의 설치 파일·리본 생성에서 방금 내보낸 설정과 XLAM의 실제 전체 경로를 입력합니다.
5. 변경 내용 확인에서 교체 대상, 보존 대상, 빠른 실행 경로 변경을 읽고 확인합니다.
6. `구매물류_리본_이PC.exportedUI`를 내려받아 Excel에서 가져옵니다.
7. 구매물류 6개 그룹과 원래의 다른 메뉴·빠른 실행 버튼을 확인합니다. Excel 재시작 후 작업본에서 필요한 기능을 실행하고 결과를 대조합니다.

Excel 가져오기는 기존 리본과 빠른 실행 도구 모음 사용자 지정을 전체 교체합니다. 생성기는 **현재 PC에서 내보낸 설정에 새 탭을 합치는 방식**입니다. 내보내기 이후에 설정을 수정했다면 새로 내보낸 파일로 다시 생성합니다. 다른 사람의 백업을 넣으면 그 사람의 메뉴를 가져오게 됩니다.

생성기는 원본 파일명을 `매크로 신상.xlam`으로 확인하고 알려진 25개 실행 프로시저와 정확히 대조한 리본 버튼만 교체합니다. 해당 버튼만 들어 있던 빈 그룹·탭은 정리합니다. 다른 파일의 동명 매크로, 모르는 명령, 다른 탭·그룹과 기본 제공 컨트롤은 보존합니다. 빠른 실행 도구 모음은 위치·이름·구성을 유지하며, 선택한 경우 알려진 대아 버튼의 경로만 갱신합니다. 파일명과 프로시저가 같더라도 실제 교체 대상인지 변경 목록을 사람이 확인해야 합니다.

XML 구조, 25개 고유 실행 연결, 사용자 경로 형식을 검사합니다. 이 검사는 PC의 XLAM 파일 존재·등록·신뢰 상태나 실제 실행 성공을 보장하지 않습니다. 브라우저 안에서 파일을 처리하며 서버로 올리지 않습니다.

`daea-ribbon-layout-20260929.exportedUI`는 **검토용 기준 구성**입니다. 기존 공개 파일의 `C:\Users\manju\AppData\Roaming\Microsoft\AddIns\매크로 신상.xlam` 경로와 원본 개인 설정을 보존한 상태입니다. 다른 PC에는 이 기준 파일을 그대로 가져오지 않고 HTML 생성기로 현재 PC용 파일을 만듭니다.

직접 구성하려면 새 탭 → 새 그룹 → 명령 선택: 매크로 → 정확한 실행 대상 추가 → 표시 이름 바꾸기 순서입니다. XLAM 매크로가 목록에 모두 나타나지 않는 경우에는 코드를 다른 파일에 복사하지 말고 준비된 메뉴 병합 방식을 사용합니다. 자세한 클릭 순서는 HTML의 리본 적용·직접 구성에 구분해 두었습니다.

## 복구

현재 업무 저장 → 새 추가 기능 해제 → 모든 Excel 종료 → 이전 XLAM을 기록한 위치로 복원 → 이전 추가 기능만 등록 → 교체 전 `.exportedUI` 가져오기 → 작업본으로 결과 확인.

메뉴 복구와 VBA 파일 복구는 별개입니다. 업무 파일에서 제거한 전용 모듈은 이후 데이터 변경을 확인한 뒤 파일 백업 또는 내보낸 `.bas`로 복구합니다. PERSONAL.XLSB·전체 사용자 지정·무관한 추가 기능을 일괄 삭제하지 않습니다.

## 설치 전에 특히 확인할 조건

- AIR·전달표 동시 생성은 FAC+Item cd 중복의 첫 행만 남기며 수량을 합산하지 않습니다. 입력 검사보다 앞서 기존 결과 시트를 삭제하는 기능도 있으므로 작업본에서 확인합니다.
- 발주 필요량 계산은 입고 내역 작성이 아닙니다. 선박/Air의 전용 열 구조가 일반 AIR 표와 다릅니다.
- 박스·파렛트 목록은 `%USERPROFILE%\OneDrive - Dae-A Electronics (Thailand) Company Limited\바탕 화면\한국단자 MOQ(20250522).xlsx`를 고정 참조합니다. 해당 파일 존재를 먼저 검사하므로 다른 위치의 파일을 열어두기만 해서는 해결되지 않습니다. PC별 실제 자료 위치를 확인해야 하며 이번 리본 변경으로 VBA 내부 참조까지 바뀌지는 않습니다.
- 두 시트 파렛트 합산은 `0714`, `0715`라는 고정 시트를 읽습니다. 임의의 두 시트 선택 기능이 아닙니다.
- 파렛트 라벨 작성은 같은 업무 파일의 MOQ 시트 또는 외부 파일 선택을 이용할 수 있습니다. 기본 머리글 F5/KET를 확인합니다. 작성자 Hyoju는 기존 코드 상수이므로 실제 출력 전 표기를 확인합니다. 자동 인쇄 명령은 없습니다.

## 버튼 전체 기능표

원본 위치 공통: `downloads/daea-macros-24-20260923.zip` → `매크로 신상.xlam` → `xl/vbaProject.bin`. 아래 코드 위치는 추출한 VBA 모듈의 줄 번호입니다. PC 설치 경로와 구분합니다.

|현재 표시 이름|실제 onAction 매크로|확인된 기능|제안 표시 이름|새 그룹|실행 전 필요한 파일·시트|결과|확인 상태|
|---|---|---|---|---|---|---|---|
|품목조회_현재시트_자동채우기|Module10.CF9_Fill_Active_Sheet|활성 시트의 품목·공장을 기준으로 BALANCE / STOCK / 입고 열을 채웁니다.|품목 수량 채우기|품목·재고|ITEM CD와 결과 열이 있는 작업 시트. 대상과 다른 원본 통합문서를 같은 Excel에서 열기. STOCK 파일명에는 STO/재고/출고, 입고 파일명에는 INBOUND/입고가 포함되어야 합니다.|대상 시트 결과 열 갱신. 완료 창에 읽은 원본 이름 표시.|코드 확인 · Excel 실행 미검증|
|거래처_조회|거래처_조회기|MAKER와 품명 예외표로 거래처명을 찾습니다.|거래처명 찾기|품목·재고|활성 시트의 MAKER·품명 열 알파벳 입력. 거래처코드 시트 A:B=거래처·MAKER, E:F=예외 거래처·품명.|맨 오른쪽에 거래처명(매칭) 열 추가.|코드 확인 · Excel 실행 미검증|
|마스터_품명_변환|자유자재_마스터_변환기|마스터에서 연호명·대아명을 찾아 함께 표시합니다.|연호·대아 품명 대조|품목·재고|활성 시트의 품명 열 알파벳 입력. 마스터 시트 A=연호명, B=대아명.|오른쪽에 마스터: 연호명 / 마스터: 대아명 열 추가.|코드 확인 · Excel 실행 미검증|
|재고_전주대비_비교|UpdateInventory_LayoutFixed|이전 목록을 유지하면서 이후 수량을 옆 열에 표시하고 신규 품목을 추가합니다.|주간 재고 비교|품목·재고|이전·이후 시트. A열=품목명, B열=수량.|업데이트결과 시트. B=지난주, C=이번주, 신규 품목 파란색.|코드 확인 · Excel 실행 미검증|
|에어_통합정리|modUnifiedMenu.에어_통합정리|열을 표준 순서로 재배치하고 ADD·최종 요청수량·공장별 조회·부족분 필터를 적용합니다.|AIR 요청표 정리|AIR|활성 시트 1행 FAC / ITEM CD / PURCHASE REQUEST 필수. 같은 파일의 공장별 Bal·Sto 및 JST 조회 시트는 선택적 원본.|현재 시트를 재작성. BALANCE 1=Balance+Stock, B-R=Balance 1−요청.|코드 확인 · Excel 실행 미검증|
|에어_재고차감_잔여요청|Air_Up_생성|요청에서 기존 STOCK을 뺀 양수 잔량만 남기고 최신 Sto up 수량을 표시합니다.|재고 차감 잔량표|AIR|Air 시트(없으면 활성 시트). FAC / Item cd / Purchase request / STOCK. F1·F2·F5 Sto up 시트는 코드열+2열 수량.|Air up 재생성. Purchase request→Remain. 충족된 행 제외.|코드 확인 · Excel 실행 미검증|
|에어_반장님_출력표|Air_To_반장님|URGENT·KET·에어장납 목록에 해당하는 행을 추려 전달표를 만듭니다.|반장님 전달표|AIR|Air 고정 열: A FAC, B Item cd, C Maker, D Part Name, E 요청, G Stock, J Pallet, L Short Del., M NEED ETD. 에어장납 B3부터 코드.|반장님 시트 재생성, Maker·품명 정렬 및 출력 서식.|코드 확인 · Excel 실행 미검증|
|에어_원클릭_정리|에어_원클릭_정리|Air 데이터를 정리한 시트와 반장님 전달표를 함께 만듭니다.|AIR·전달표 동시 생성|AIR|Air 시트의 FAC / Item cd / Purchase request. 선택적으로 같은 파일의 Bal·Sto·에어장납.|AIR_정리와 반장님 시트 재생성.|코드 확인 · Excel 실행 미검증|
|에어_흑백인쇄_서식|Air_Print_Grayscale|활성 AIR 표를 흑백 인쇄에 맞게 서식·열 너비·인쇄 영역으로 정리합니다.|흑백 인쇄 서식|AIR|상단 10행 안에 Maker 헤더가 있는 활성 작업 시트.|현재 시트에 A4 가로, 너비 1페이지, 높이 자동, Maker 구분선 적용.|코드 확인 · Excel 실행 미검증|
|선박_품목코드_정리|선박정리거래처코드|주별 수량 합계와 공장별 Balance·Stock·거래처를 정리합니다.|선박 수량·거래처 정리|선박·집계|활성 시트 1~3행 SUM 헤더, E열~SUM 앞까지 수량. A 공장·B 품목코드·C Maker·D 품명 구조, 같은 파일의 조회 시트와 거래처코드.|현재 시트의 SUM과 뒤쪽 Balance / Stock / B-R / 거래처 열 갱신.|코드 확인 · Excel 실행 미검증|
|출고_공장별_합계|StockYES|고객주문번호의 앞 2글자 공장과 품목코드·품명 기준으로 출고수량을 합산합니다.|출고수량 합계|선박·집계|활성 ERP 출고표 C=품명, I=출고수량, T=고객주문번호, U=거래처 품목코드.|출고수량합계_시간 시트 생성.|코드 확인 · Excel 실행 미검증|
|잔량_공장별_합계|Balance_onetimeatall|수주처·품목코드·품명별 미납수량을 합산합니다.|미납수량 합계|선박·집계|활성 ERP 수주표 B=수주처, E=거래처 품목코드, G=품명, N=미납수량.|미납수량합계_시간 시트 생성.|코드 확인 · Excel 실행 미검증|
|부분합_만들기|modUnifiedMenu.부분합_만들기|한 기준 또는 두 기준으로 선택한 표의 수량·금액을 합산합니다.|선택 범위 부분합|선박·집계|헤더 포함 연속 범위를 선택. 기준 열·합산 열은 선택 범위 안의 번호로 입력.|새 결과 시트에 그룹별 합계 및 선택한 추가 열 표시.|코드 확인 · Excel 실행 미검증|
|연호_입고계산|modUnifiedMenu.연호_입고계산|선박·AIR 수요, 본사대기, 연호 ERP 대기량을 반영해 발주 필요량을 계산합니다.|발주 필요량 계산|연호|연호 ERP(A 품명/B 수량/C 공장), 마스터(A 연호명/B 대아명), 선박(A 공장/B 코드/C 품명/D 수요/E 본사대기/F Balance), Air(A 공장/B 코드/C 품명/D 수요/E Balance). 별도 재고 선택 시 A 공장/B 코드/C 품명/D 수량.|최종발주 재작성. 선박·AIR·필요합·본사대기·연호대기·발주필요량·발란스.|코드 확인 · Excel 실행 미검증|
|연호_입고수량_대조|Run_Final_연호입고점검0000|두 ERP 자료를 공장·품명 기준으로 합산한 뒤 수량 차이를 계산합니다.|연호·대아 수량 대조|연호|연호 ERP A 품명/B 수량/C 공장. 대아 ERP A 공장/B 품명/C 수량. 마스터 A 연호명/B 대아명.|최종수량대조_결과 재작성. 차이=연호−대아.|코드 확인 · Excel 실행 미검증|
|연호_거래명세_입고기입|FillYeonhoReceiving_AllInOne_Fixed|거래명세 품명·공장별 수량을 합산하고 연호매칭으로 입고표를 만듭니다.|거래명세 → 입고표|연호|이름에 거래명세서가 포함된 첫 시트. 연호매칭 A 대아명/B 거래처 품목코드/C 고객 품목코드. 명세서 B 품명·C 수량·G 비고 구조.|연호입고 시트 재작성. ITEM CD·ITEM NM·입고수량 등 표시.|코드 확인 · Excel 실행 미검증|
|연호_거래명세_정리|연호거래명세|반복 머리글·불필요 행을 제외하고 품명과 비고를 정돈합니다.|거래명세 원형 정리|연호|활성 거래명세서. 19행 머리글, 20행부터 데이터. B 품명·G 비고.|연호_정제_시분 새 시트. 원래 열 구조 유지.|코드 확인 · Excel 실행 미검증|
|연호_거래명세_3열출력|연호거래명세_3열출력|공장 비고와 품명 기준으로 수량을 합쳐 간결한 표를 만듭니다.|거래명세 3열 요약|연호|활성 거래명세서 20행부터 데이터. B 품명/C 수량/G 비고. G열 앞 2글자를 공장 구분으로 사용합니다.|연호_정제_시분 시트: 비고 / 품명 및 규격 / 수량.|코드 확인 · Excel 실행 미검증|
|연호_반장님_재고목록|CreatePerfectFinalInventoryList|거래명세서의 품명별 수량을 합산해 재고 확인용 인쇄표를 만듭니다.|반장님 재고표|연호|거래명세서 (19) 시트 우선, 없으면 활성 시트. A 번호/B 품명/C 수량.|재고 리스트 시트 재생성. 작성일자와 인쇄용 빈 기록 칸 포함.|코드 확인 · Excel 실행 미검증|
|KET_대기재고_업데이트|KET이후값_업데이트_정렬|이후 수량으로 갱신하고 신규 품목을 넣으며 이후에 없는 품목을 제외합니다.|대기재고 갱신|KET·파렛트|이전·이후 시트 A 품목명/B 수량.|업데이트결과 재생성. 갱신 수량은 B열, 신규 품목 파란색.|코드 확인 · Excel 실행 미검증|
|파렛트_위치_정리|스마트_파렛트_정밀정리000|출고수량과 Sum 사이에서 값이 있는 파렛트 열 제목을 모읍니다.|품목별 파렛트 위치|KET·파렛트|활성 시트 1행 출고수량 / Sum. A열 품목명, 그 사이 파렛트 수량.|Sum 뒤에 확인 품목 / 해당 파렛트 두 열 작성.|코드 확인 · Excel 실행 미검증|
|파렛트_박스목록_생성|파렛트_박스정리_생성|MOQ의 박스당 수량과 파렛트별 수량·위치를 결합합니다.|박스·파렛트 목록|KET·파렛트|활성 시트 A 품명, C열부터 파렛트, 소문자 sum 앞까지. MOQ 파일의 ITEM NM / BOX 헤더.|파렛트정리 시트 재생성. 품명·박스당·총수량·위치/수량.|코드 확인 · Excel 실행 미검증|
|KET_파렛트_품번합산|파렛트_전체_품번합산|두 시트의 같은 품명·파렛트 수량을 합칩니다.|두 시트 파렛트 합산|KET·파렛트|코드에 고정된 0714 및 0715 시트. A 품목명, 파렛트 번호 헤더.|통합 시트 재생성. 품목별 출고합계·파렛트 수량·위치.|코드 확인 · Excel 실행 미검증|
|KET_파렛트_라벨인쇄|LabelPa.MakePalletLabels|파렛트별 품목·박스 수량을 담은 인쇄용 라벨 시트를 만듭니다.|파렛트 라벨 작성|KET·파렛트|정리 시트의 품목명과 숫자 파렛트 헤더. MOQ 시트 우선, 없으면 외부 MOQ 검색 및 파일 선택.|Pallet Labels 재생성. 머리글 입력 창 표시.|코드 확인 · Excel 실행 미검증|
|품목조회_실행버튼_만들기|Module10.CF9_Install_One_Click_Button|현재 시트에 품목 수량 조회용 도형 버튼을 추가합니다.|시트 조회 버튼 추가|설정|버튼을 놓을 업무 시트를 선택.|btnItemAutoFill 도형을 만들고 ITEM AUTO FILL 표시.|코드 확인 · Excel 실행 미검증|

### 코드 근거와 기능별 주의

- **품목 수량 채우기** — `Module10.bas:6`: 대상 파일 내부 시트는 원본으로 읽지 않습니다. 여러 Balance 원본을 읽을 수 있어 다른 기간 파일은 닫습니다. 미조회 Balance는 #N/A, 재고·입고는 0이므로 원본 누락과 구별합니다.
- **거래처명 찾기** — `FIND_COMPANY.bas:3`: 결과 공란과 예외 품목을 확인합니다. 반복 실행하면 결과 열이 추가됩니다.
- **연호·대아 품명 대조** — `MASTER_CHANGER.bas:3`: 원래 품명 셀을 일괄 치환하는 기능은 아닙니다. 미등록 품명은 매칭 결과를 확인합니다.
- **주간 재고 비교** — `BEFORE_AFTER.bas:2`: 업데이트결과를 재생성합니다. 이후에 없는 이전 품목도 남습니다. KET 대기재고 갱신과 처리 방식이 다릅니다.
- **AIR 요청표 정리** — `modUnifiedMenu.bas:3`: air_add → Module6 호출. 원본 열 중 지정되지 않은 열은 사라질 수 있습니다. 같은 공장·품번 행을 중복 제거하지 않습니다. B-R<0 필터가 켜집니다.
- **재고 차감 잔량표** — `Module5.bas:10`: 기존 Air up은 입력 헤더 검사 전에 삭제됩니다. Sto up 미조회는 0으로 표시합니다. 필터링된 행과 총 잔량을 대조합니다.
- **반장님 전달표** — `AirPrint.bas:8`: 출력표 생성이며 자동 인쇄가 아닙니다. Air 열 위치가 다르면 먼저 확인합니다. 출력된 수량 0 행 여부도 확인합니다.
- **AIR·전달표 동시 생성** — `Module3.bas:11`: FAC+Item cd가 같은 행은 첫 행만 유지하며 수량을 합산하지 않습니다. 기존 두 결과를 입력 검사 전에 삭제합니다. 복수 요청을 보존해야 하면 AIR 요청표 정리를 검토합니다.
- **흑백 인쇄 서식** — `Module9.bas:5`: 실제 인쇄 명령은 없습니다. Maker 변경 위치에 선을 넣으며 Maker순 재정렬 기능은 아닙니다. Ctrl+P로 확인합니다.
- **선박 수량·거래처 정리** — `Ship.bas:3`: SUM=0 행을 삭제합니다. SUM 오른쪽 5열을 결과로 사용하므로 기존 데이터 위치를 확인합니다.
- **출고수량 합계** — `STOCK.bas:2`: 정해진 열 위치를 읽습니다. 숫자가 아닌 출고수량은 0으로 취급하므로 원본 숫자 형식을 확인합니다.
- **미납수량 합계** — `BALANCE_ATONE.bas:2`: 1010→F1, 1060→F2, 1150→F5, 1050→삼도. 숫자가 아닌 수량은 0으로 처리합니다.
- **선택 범위 부분합** — `modUnifiedMenu.bas:12`: modSubtotal 호출. 추가 정보는 각 그룹 첫 행 값을 사용합니다. 추가 정보가 모두 같은지 별도로 확인합니다.
- **발주 필요량 계산** — `modUnifiedMenu.bas:6`: Module4 호출. 예=재고 시트 사용, 아니오=선박 E열 본사대기. 이 Air는 일반 AIR 요청표와 열 구조가 다릅니다. 선택한 모드와 원본 합계를 확인합니다.
- **연호·대아 수량 대조** — `Yeonhocheck.bas:3`: 문자형 1,000 등에 Val을 사용하므로 수량을 실제 숫자로 준비합니다. 차이가 0이어도 공장·품명 매칭을 확인합니다.
- **거래명세 → 입고표** — `YEONHO_INCOME_EXCEL.bas:2`: 연호매칭은 XLAM 안을 먼저 찾고 없으면 업무 파일을 찾습니다. 현재 배포본은 업무 파일에 연호매칭을 준비하는 방식입니다.
- **거래명세 원형 정리** — `YEONHO_TABLE.bas:2`: 수량 합산용 3열 요약과 다릅니다. 같은 분에 다시 실행하면 결과 시트 이름 충돌 가능성이 있습니다.
- **거래명세 3열 요약** — `YeonhoTablefine.bas:2`: 원형 정리와 결과 이름이 같아 같은 분에 실행하면 충돌할 수 있습니다. 원본 명세서에서 실행합니다.
- **반장님 재고표** — `YeonhoBanjang.bas:2`: 공장별로 나누는 집계가 아니라 품명별 합계입니다. 작성일자를 못 읽으면 오늘 날짜를 사용합니다. 자동 인쇄는 하지 않습니다.
- **대기재고 갱신** — `KET_BEFORE_AFTER.bas:2`: 주간 재고 비교와 달리 이후에 없는 품목은 삭제됩니다. 두 기능은 같은 결과 시트 이름을 사용합니다.
- **품목별 파렛트 위치** — `Pallet.bas:2`: 빈칸 여부로 판단하여 숫자 0도 위치로 포함할 수 있습니다. Sum 뒤 두 열의 기존 내용을 확인합니다.
- **박스·파렛트 목록** — `PalletOne.bas:2`: USERPROFILE 아래 지정된 회사 OneDrive 바탕 화면의 한국단자 MOQ(20250522).xlsx가 필요합니다. 파일 존재를 먼저 검사해 단순히 다른 위치에서 열어두는 것으로 해결되지 않습니다. 계장님 PC 경로 확인 후 사용합니다.
- **두 시트 파렛트 합산** — `KETpalletmerge.bas:2`: 임의 두 시트를 선택하는 기능이 아닙니다. 실제 대상이 0714·0715 구조인지 확인해야 합니다. 날짜만 맞추려고 업무 시트 이름을 임의로 바꾸지 않습니다.
- **파렛트 라벨 작성** — `LabelPa.bas:22`: 기본 머리글 F5 / KET는 입력 창에서 확인합니다. 작성자 Hyoju는 코드 상수라 계장님 PC에서도 그대로 표시됩니다. 실제 출력 전 확인하세요. 자동 인쇄는 하지 않습니다.
- **시트 조회 버튼 추가** — `Module10.bas:93`: 같은 이름의 도형은 교체합니다. 연결은 CF9_Fill_Active_Sheet로 지정되므로 구버전 중복 설치를 정리하고 버튼의 매크로 지정을 확인합니다. Excel 리본 탭 생성 기능은 아닙니다.

## 검증 결과

- Chromium 실제 브라우저 자동 점검 52개 통과. JavaScript 실행 오류 0개.
- 데스크톱 1366×900 / 1280×800 및 작은 화면 390×844에서 화면 확인. 한국어 렌더링은 검사 환경에 설치한 Noto Sans KR로 보완했습니다. 제공 HTML은 Windows 맑은 고딕/Apple SD Gothic Neo를 사용합니다.
- HTML5 파서 오류 0개. 중복 ID 없음. JavaScript를 끈 경우에도 19단계 본문 표시.
- 단계 이동, 완료 표시, 새로고침 후 메모 유지, 이전 저장 키의 기록 이관, JSON 내보내기·가져오기, 검색, 기능별 상세 보기, 모바일 메뉴 확인.
- 25개 실행 프로시저 집합 일치, 중복 없음, 6개 그룹별 개수 일치, 새 파일 재병합 시 중복 없음.
- 원본의 무관한 탭·QAT 보존, 다른 XLAM의 동명 매크로 보존, 대아 QAT 경로만 갱신/갱신 해제 검증. UTF-16 입력·한국어/특수문자 경로·UNC 경로·잘못된 경로/XML 차단 확인.
- 변경 확인 이후 파일·경로를 바꾸면 기존 다운로드 승인을 무효화합니다.
- XLAM 브라우저 다운로드를 원본과 SHA-256 대조했습니다.

**남은 확인:** 이 실행 환경에는 Windows Excel이 없습니다. Excel의 실제 가져오기, XLAM 로드, 25개 매크로의 업무 결과는 직접 실행하지 않았습니다. 계장님 PC에서 실제 경로·추가 기능 체크·기존 메뉴 유지·입력 자료 및 결과를 확인해야 합니다.

### 원본 파일 SHA-256

- `적용방법_및_변경사항.md`: `ef0bf3c4199264ba7c8f359bf5db8fa70bda7166f2a2626a03f4e344f64875d3`
- `매크로 신상.xlam`: `b99b73f67268aff289b970dc12f5a4f17e505b6539b0c8ced16797a0acccd478`
- `메뉴_주24개_보조1개.exportedUI`: `0a13f978a790cba7a99882cd0cb2789abe6a4fe2ce3576c51a707bdb64fe0fbf`

### 실행한 브라우저 점검

- old URL redirects to new guide
- valid head/body; no duplicate IDs
- initial step and 25 catalog rows
- step 1 navigation
- step 2 navigation
- step 3 navigation
- step 4 navigation
- step 5 navigation
- step 6 navigation
- step 7 navigation
- step 8 navigation
- step 9 navigation
- step 10 navigation
- step 11 navigation
- step 12 navigation
- step 13 navigation
- step 14 navigation
- step 15 navigation
- step 16 navigation
- step 17 navigation
- step 18 navigation
- step 19 navigation
- completion and notes persist
- navigation does not mark done
- search returns code-specific conditions
- modal functional detail
- catalog filter
- CSV contains all 25 mappings
- manual ribbon lesson pagination
- missing macro list troubleshooting
- XLAM download byte-identical to production ZIP
- no baseline rejected
- merge preview25 confirmed; gated download
- generated file removes original user path in actions
- changed path invalidates reviewed download
- 25 unique actions preserved verbatim
- six groups 4/5/4/6/5/1
- original25 removed once; regeneration idempotent
- Unicode and XML-special path encoded correctly
- unrelated tabs and QAT preserved
- same macro in another file preserved
- QAT updates only known button; optout supported
- invalid paths and XML rejected; UNC accepted
- UTF16 menu supported
- record export excludes uploaded menu contents
- record import restores step
- legacy progress migration
- mobile body fits viewport
- mobile menu opens
- mobile route closes menu
- no page JavaScript errors
- no-script installation instructions remain readable

## 공식 참고

- [Microsoft — Office 리본 사용자 지정](https://support.microsoft.com/ko-kr/office/foundations-experiences/customize-the-ribbon-in-office)
- [Microsoft — Excel 추가 기능 추가·제거](https://support.microsoft.com/ko-kr/excel/add-or-remove-add-ins-in-excel)
- [Microsoft — 단추에 매크로 지정](https://support.microsoft.com/ko-kr/excel/assign-a-macro-to-a-button)
