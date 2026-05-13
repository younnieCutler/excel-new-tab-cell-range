# CellFocus — Excel 집중 편집 Add-in 구현 계획

복잡한 Excel 파일에서 필요한 셀 범위만 골라 새 탭(패널 내 탭)으로 분리하고, 원본과 실시간 동기화하며, 리본/도구 모음 없이 **집중 편집**하는 Office.js Excel Add-in.

## 사용상황 재정의 — SharePoint xlsx + Windows Excel Desktop 공동편집

이 Add-in의 1차 사용 경로는 브라우저 Excel이 아니라, **SharePoint/OneDrive에 저장된 `.xlsx` 파일을 Windows Excel 데스크톱 앱에서 직접 열어 작업하는 상황**이다. 사용자는 SharePoint 웹 UI에서 파일을 찾더라도 실제 편집은 **"데스크톱 앱에서 열기"**로 진행한다.

중요한 구분:
- **하지 않는 것**: 브라우저에서 Excel Online을 열어 작업하는 흐름을 1차 목표로 삼지 않는다.
- **하는 것**: Windows Excel 데스크톱 앱에서 `.xlsx`를 열고, 그 Excel 창 안의 Office Add-in TaskPane으로 CellFocus를 사용한다.
- **기술적 현실**: Office Add-in UI는 데스크톱 Excel 안에서도 WebView 기반 TaskPane으로 렌더링된다. 하지만 사용자 경험은 브라우저 탭이 아니라 **Excel 데스크톱 창 내부 패널**이다.

### 실제 워크플로우

1. 사용자가 SharePoint 문서 라이브러리에서 **데스크톱 앱으로 열기**를 누르거나, OneDrive 동기화 폴더의 `.xlsx`를 더블클릭한다.
2. Windows Excel Desktop에서 AutoSave가 켜진 상태로 여러 사람이 같은 통합문서를 동시에 편집한다.
3. 사용자는 복잡한 시트 전체를 보지 않고, 특정 범위만 CellFocus 탭으로 분리해 집중 편집한다.
4. 다른 사용자가 원본 범위를 수정하면 CellFocus 탭도 stale 상태가 되지 않아야 한다.
5. 사용자가 CellFocus에서 값을 쓰는 순간, 원본 셀에 덮어쓰기 전에 최신 상태와 충돌 가능성을 확인해야 한다.

### 로컬 작업 요구사항

- 사용자는 `.xlsx`를 브라우저에서 편집하지 않고 **Windows Excel 앱으로 직접 연다**.
- CellFocus는 별도 웹사이트가 아니라 **열려 있는 Excel 데스크톱 창에 붙는 Add-in**이어야 한다.
- 개발 중에는 `npm start`로 로컬 개발 서버를 띄우고 `manifest.dev.xml`을 Windows Excel Desktop에 sideload한다.
- 실제 사용 중에는 SharePoint/OneDrive 파일이라도 workbook은 Excel Desktop에서 열리며, Add-in은 해당 workbook 세션의 선택 범위와 이벤트를 읽고 쓴다.
- 네트워크가 잠시 흔들려도 사용자가 입력 중인 draft를 잃지 않아야 한다. 원본 반영이 불확실하면 `Local draft` 또는 `Conflict` 상태로 남겨야 한다.

### 이 전제가 바꾸는 핵심 문제

Excel/SharePoint는 **통합문서 내용**을 동기화하지만, Add-in의 JavaScript 메모리와 탭 상태는 사용자별 로컬 인스턴스다. 즉, 공동편집 환경에서는 다음 문제가 실제로 발생할 수 있다.

| 위험 | 사용자에게 보이는 증상 | 설계 요구 |
|------|------------------------|-----------|
| Stale cache | 다른 사람이 바꾼 값이 CellFocus 탭에 늦게 반영됨 | `onChanged`/Binding 이벤트 기반 재조회 |
| Silent overwrite | A가 오래된 값을 보고 편집해 B의 최신 값을 덮음 | 쓰기 직전 셀 재조회 + 충돌 감지 |
| 원격 변경 중 로컬 편집 | 사용자가 입력 중인 셀이 원격 변경됨 | 편집 중 셀은 즉시 덮어쓰지 말고 conflict 표시 |
| Add-in 세션 재시작 | Excel을 닫았다 열면 이벤트 리스너와 메모리 상태 소실 | 탭 복원은 localStorage 가능, 리스너는 재등록 필요 |
| SharePoint sync 지연 | AutoSave/네트워크 상태에 따라 반영 타이밍이 흔들림 | 상태바에 `Synced / Refreshing / Conflict / Local draft` 표시 |

### 제품 원칙

- **Windows Desktop-first**: Windows Excel Desktop + SharePoint/OneDrive 저장 파일을 기준으로 설계한다.
- **Web은 보조 검증**: Excel Web 지원은 가능성 검증 수준으로 두고, 1차 성공 기준에 넣지 않는다.
- **낙관적 쓰기 + 충돌 감지**: 대부분의 편집은 즉시 쓰되, 쓰기 직전 원본 셀 값이 마지막으로 읽은 값과 다르면 사용자에게 선택권을 준다.
- **값 손실 방지 우선**: 집중 편집 UX보다 공동편집 중 데이터 손실 방지가 우선이다.
- **메모리는 진실이 아니다**: CellFocus 탭의 `cells` 캐시는 렌더링 최적화용일 뿐, 쓰기의 근거가 되는 단일 진실로 취급하지 않는다.

## 아키텍처 요약

```mermaid
graph LR
    subgraph Excel Host
        A[사용자 셀 선택] -->|우클릭 / Ctrl+Shift+F| B[CellFocus 트리거]
    end
    subgraph CellFocus Add-in TaskPane
        B --> C[Range 데이터 로드]
        C --> D[탭 생성 — 미니멀 그리드 UI]
        D -->|사용자 편집| E[변경 감지]
        E -->|Excel.run write| F[원본 셀 업데이트]
        F -->|onChanged event| D
    end
```

---

## User Review Required

> [!IMPORTANT]
> **단일 TaskPane 내 탭 구조 vs. Dialog 팝업 방식**
> Office.js 제약상 진정한 "새 브라우저 탭"은 불가능합니다. 두 가지 대안이 있습니다:
> 1. **TaskPane 내 탭 UI** (추천) — 오른쪽 패널에 탭을 만들어 여러 범위를 전환. 안정적이고 Excel API 접근이 자유롭다.
> 2. **displayDialogAsync 팝업** — 독립 윈도우로 열림. 좀 더 "분리된 느낌"이지만, 한 번에 1개 Dialog만 가능하고, Excel API 직접 호출 불가(메시지 패싱 필요).
>
> **→ 방식 1(TaskPane 탭 UI)을 기본으로 진행합니다. 변경 원하시면 알려주세요.**

> [!WARNING]
> **배포 방식**: 사용자는 브라우저가 아니라 Windows Excel Desktop에서 사용합니다. 개발/개인 사용은 `manifest.xml` 수동 Sideload, 조직 배포는 Microsoft 365 Admin 중앙 배포가 기준입니다. GitHub Pages는 Add-in 파일을 제공하는 정적 호스팅일 뿐, 사용자가 브라우저 Excel에서 작업한다는 뜻이 아닙니다.

---

## Open Questions

> [!IMPORTANT]
> 1. **Add-in 이름**: "CellFocus"로 진행해도 될까요? 다른 이름 선호 시 알려주세요.
> 2. **다국어 지원**: 한국어만? 영어도 포함? (UI 라벨, 메뉴 항목 등)
> 3. **셀 서식 동기화 수준**: 값만 동기화? 아니면 폰트, 배경색, 테두리 등 서식도 표시?
> 4. **수식 지원**: 수식이 있는 셀 편집 시, 수식 자체를 편집? 아니면 계산된 값만 표시?

---

## 기술 스택

| 구분 | 선택 | 이유 |
|------|------|------|
| **프레임워크** | Office.js (Excel JavaScript API) | 공식 Add-in 플랫폼 |
| **런타임** | Shared Runtime | 컨텍스트 메뉴 + TaskPane + 키보드 단축키 모두 동일 JS 컨텍스트 |
| **UI** | Vanilla HTML/CSS/JS | 의존성 최소화, 번들 크기 최적화 |
| **빌드** | Webpack (yo office 기본) | 공식 scaffold 도구 |
| **호스팅** | 고객사 테넌트 HTTPS 정적 호스팅 | Excel Desktop Add-in은 브라우저에서 실행되는 것이 아니라 Excel 창 안의 TaskPane에서 로드됨. 단, Office Add-in 리소스 자체는 HTTPS로 제공되어야 함 |
| **매니페스트** | XML (Add-in only manifest) | 호환성 최대화 |
| **1차 실행 환경** | Windows Excel Desktop + SharePoint/OneDrive `.xlsx` | 사용자가 실제로 작업하는 경로 |

---

## Proposed Changes

### 1. 프로젝트 초기화

#### [NEW] 프로젝트 스캐폴딩
- `yo office` (Yeoman Generator)로 Excel TaskPane 프로젝트 생성
- JavaScript + Shared Runtime 선택
- 프로젝트 루트: `c:\Users\金貞潤\Documents\excel-new-tab-specific-cell`

생성되는 기본 구조:
```
excel-new-tab-specific-cell/
├── manifest.xml          # Add-in 매니페스트
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html # TaskPane UI
│   │   ├── taskpane.css  # 스타일
│   │   └── taskpane.js   # 로직
│   └── commands/
│       └── commands.js   # 리본/컨텍스트 메뉴 명령
├── webpack.config.js
├── package.json
└── ...
```

---

### 2. 매니페스트 설정 (manifest.xml)

#### [MODIFY] manifest.xml
핵심 설정:
- **Shared Runtime** 활성화 → TaskPane + Commands 같은 JS 런타임 공유
- **ContextMenu** 확장점 → 셀 우클릭 시 "CellFocus로 열기" 메뉴 추가
- **Keyboard Shortcuts** → `Ctrl+Shift+F` 단축키 등록
- **Ribbon 최소화** → 리본에 작은 아이콘 버튼 1개만 추가

```xml
<!-- 컨텍스트 메뉴 확장점 -->
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Button" id="CellFocusContextBtn">
      <Label resid="CellFocus.ContextMenu.Label" />
      <Action xsi:type="ExecuteFunction">
        <FunctionName>openInCellFocus</FunctionName>
      </Action>
    </Control>
  </OfficeMenu>
</ExtensionPoint>

<!-- 키보드 단축키 -->
<!-- shortcuts.json 파일로 Ctrl+Shift+F 매핑 -->
```

---

### 3. 핵심 기능 구현

#### [NEW] src/taskpane/taskpane.js — 메인 로직

**3-1. 셀 범위 캡처 & 탭 생성**
```javascript
async function captureSelectedRange() {
    await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address, values, worksheet/name, rowCount, columnCount");
        await context.sync();
        
        // 탭 데이터 구조 생성
        const tabData = {
            id: generateId(),
            sheetName: range.worksheet.name,
            address: range.address,
            values: range.values,
            rowCount: range.rowCount,
            colCount: range.columnCount
        };
        
        addTab(tabData);
    });
}
```

**3-2. 실시간 동기화 (Excel → Add-in)**
```javascript
async function registerChangeListener(tabData) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(tabData.sheetName);
        sheet.onChanged.add(async (eventArgs) => {
            // 변경된 셀이 감시 범위에 포함되는지 확인
            if (isWithinRange(eventArgs.address, tabData.address)) {
                await refreshTabData(tabData.id);
            }
        });
        await context.sync();
    });
}
```

**3-3. 공동편집 내성 있는 쓰기 (Add-in → Excel)**

쓰기 직전에는 Add-in이 마지막으로 읽은 값(`baseValue`)과 현재 원본 셀 값을 비교한다. 값이 다르면 다른 공동편집자가 먼저 수정했을 가능성이 있으므로 즉시 덮어쓰지 않는다.

```javascript
async function writeBackToExcel(tabId, row, col, newValue) {
    const tab = tabs.get(tabId);
    await Excel.run(async (context) => {
        context.runtime.enableEvents = false; // 루프 방지
        const sheet = context.workbook.worksheets.getItem(tab.sheetName);
        const range = sheet.getRange(tab.address);
        const cell = range.getCell(row, col);

        cell.load("values, text");
        await context.sync();

        const currentValue = cell.values[0][0];
        const baseValue = tab.cells[row][col].value;

        if (currentValue !== baseValue) {
            markConflict(tabId, row, col, {
                baseValue,
                currentValue,
                pendingValue: newValue
            });
            context.runtime.enableEvents = true;
            return;
        }

        cell.values = [[newValue]];
        await context.sync();
        context.runtime.enableEvents = true;
    });
}
```

충돌 시 UI 선택지:
- **Reload**: 원본 최신 값으로 CellFocus 셀을 갱신하고 내 입력은 버린다.
- **Overwrite**: 현재 원본 값을 확인한 뒤 내 입력으로 덮어쓴다.
- **Keep draft**: 원본은 그대로 두고 CellFocus 셀에 draft 표시를 유지한다.

**3-4. 원격 변경 처리 정책**

`onChanged` 이벤트가 들어왔을 때 변경 소스가 원격 공동편집자인지 확인할 수 있으면 원격 변경으로 표시한다. 활성 편집 중인 셀은 즉시 DOM을 갈아엎지 않고, 셀 우상단에 `Updated remotely` 상태를 띄운 뒤 사용자가 편집을 끝낼 때 충돌 처리로 넘긴다.

```javascript
async function handleExcelChange(tabId, changedAddress, source) {
    const tab = tabs.get(tabId);
    if (!tab) return;

    if (isEditingIntersectingCell(tab, changedAddress)) {
        markRemoteChangedDuringEdit(tabId, changedAddress, source);
        return;
    }

    const latest = await captureRange(tab.sheetName, tab.address);
    updateTabWithLatest(tabId, latest);
}
```

---

#### [NEW] src/taskpane/taskpane.html — 집중 편집 UI

**UI 구성:**
```
┌─────────────────────────────────────┐
│ [Tab1: Sheet1!A1:D10] [Tab2] [+ ✕] │  ← 탭 바
├─────────────────────────────────────┤
│                                     │
│   ┌───┬───┬───┬───┐               │
│   │ A1│ B1│ C1│ D1│               │  ← 미니멀 그리드
│   ├───┼───┼───┼───┤               │     (editable cells)
│   │ A2│ B2│ C2│ D2│               │
│   ├───┼───┼───┼───┤               │
│   │ A3│ B3│ C3│ D3│               │
│   └───┴───┴───┴───┘               │
│                                     │
│ ● Synced  |  Sheet1!A1:D10        │  ← 상태 바
└─────────────────────────────────────┘
```

핵심 디자인 원칙:
- **제로 크롬**: 리본, 도구 모음 없음. 셀 그리드 + 탭 바만
- **다크 모드 기본**: 눈의 피로 감소, 집중 환경
- **미니멀 그리드**: contenteditable div 기반 또는 `<input>` 기반 셀 그리드
- **동기화 인디케이터**: 실시간 연결 상태 표시 (`Synced`, `Refreshing`, `Conflict`, `Local draft`)
- **충돌 표시**: 원격 변경과 내 draft가 충돌한 셀은 테두리/배지로 표시하고 덮어쓰기 전 확인

---

#### [NEW] src/taskpane/taskpane.css — 프리미엄 미니멀 스타일

```css
/* Design Tokens */
:root {
    --bg-primary: #1a1a2e;
    --bg-secondary: #16213e;
    --bg-cell: #0f3460;
    --text-primary: #e8e8e8;
    --text-secondary: #a0a0b0;
    --accent: #00d2ff;
    --accent-glow: rgba(0, 210, 255, 0.15);
    --border: rgba(255, 255, 255, 0.06);
    --sync-green: #00e676;
    --danger: #ff5252;
    --radius: 8px;
    --transition: 200ms ease;
}
```

특징:
- 글래스모피즘 탭 바
- 셀 호버/포커스 시 부드러운 글로우 효과
- 편집 중 셀 하이라이트 애니메이션
- 동기화 상태 펄스 애니메이션

---

#### [NEW] src/commands/commands.js — 명령 처리

```javascript
// 컨텍스트 메뉴 & 단축키에서 호출되는 함수
function openInCellFocus(event) {
    // TaskPane 표시 + 선택된 범위 캡처
    Office.addin.showAsTaskpane();
    captureSelectedRange();
    event.completed();
}

// 단축키 액션 연결
Office.actions.associate("openInCellFocus", openInCellFocus);
```

---

#### [NEW] src/shortcuts.json — 키보드 단축키 정의

```json
{
    "actions": [
        {
            "id": "openInCellFocus",
            "type": "ExecuteFunction",
            "name": "Open in CellFocus"
        }
    ],
    "shortcuts": [
        {
            "action": "openInCellFocus",
            "key": {
                "default": "Ctrl+Shift+F"
            }
        }
    ]
}
```

---

### 4. 탭 관리 시스템

#### [NEW] src/taskpane/modules/tabManager.js

```javascript
class TabManager {
    constructor() {
        this.tabs = new Map();      // tabId → tabData
        this.activeTabId = null;
        this.listeners = new Map(); // tabId → eventHandler
    }
    
    addTab(tabData) { /* ... */ }
    removeTab(tabId) { /* ... */ }
    switchTab(tabId) { /* ... */ }
    refreshTab(tabId) { /* ... */ }
}
```

핵심 기능:
- 최대 8개 탭 동시 관리
- 탭 간 빠른 전환 (Ctrl+1~8)
- 탭 닫기 시 이벤트 리스너 해제
- 탭 상태 localStorage 캐싱
- Excel 재시작/TaskPane reload 시 캐시된 탭의 이벤트 리스너 재등록
- 캐시 복원 후 즉시 원본 범위를 재조회해 stale 상태 제거
- 각 셀에 `baseValue`, `currentDisplayText`, `dirty`, `conflict` 메타데이터 유지

---

### 5. 그리드 렌더러

#### [NEW] src/taskpane/modules/gridRenderer.js

순수 DOM 기반 편집 가능 그리드:
- `<table>` 기반 렌더링 (경량, 빠른 렌더링)
- 각 셀 = `<td>` + `<input>` (포커스 시 활성화)
- 방향키/Tab/Enter로 셀 간 이동
- 더블클릭 또는 타이핑 시작으로 편집 모드 진입
- 대규모 범위 대응: 가상 스크롤링 (1000행 이상 시)

---

### 6. 고객사 배포 설정

#### [NEW] 고객사 manifest 생성 스크립트
- `npm run build:customer -- --base-url https://cellfocus.customer.example`
- `dist/` 정적 파일을 생성하고, 고객사 HTTPS base URL이 주입된 `dist/manifest.xml`을 생성
- `SourceLocation`, `Commands.Url`, `ExtendedOverrides`, 아이콘 URL, `AppDomain`을 고객사 도메인으로 통일
- 고객사 지원 페이지가 있으면 `--support-url`로 별도 지정

#### 배포 플로우:
```
코드 빌드 → npm run build:customer → dist/ 생성 → 고객사 HTTPS 정적 호스팅에 업로드
                                                          ↓
                      고객사 관리자: dist/manifest.xml을 Microsoft 365 Admin 중앙 배포
                                                          ↓
                                    사용자: Windows Excel Desktop에서 Add-in 실행
```

주의: 이 플로우에서 호스팅 계층은 고객사가 소유한다. 사용자의 작업 화면은 브라우저가 아니라 Windows Excel Desktop이며, CellFocus는 Excel 창 오른쪽 TaskPane으로 열린다.

---

## 파일 구조 최종

```
excel-new-tab-specific-cell/
├── manifest.xml                    # Add-in 매니페스트 (Shared Runtime + ContextMenu)
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html          # 메인 UI (탭 바 + 그리드)
│   │   ├── taskpane.css           # 프리미엄 다크 테마 스타일
│   │   ├── taskpane.js            # 엔트리포인트 + 초기화
│   │   └── modules/
│   │       ├── tabManager.js      # 탭 CRUD & 상태 관리
│   │       ├── gridRenderer.js    # 편집 가능 셀 그리드
│   │       ├── syncEngine.js      # 양방향 데이터 동기화
│   │       └── utils.js           # 유틸리티 (범위 파싱, ID 생성 등)
│   └── commands/
│       └── commands.js            # 컨텍스트 메뉴 & 단축키 핸들러
├── src/shortcuts.json             # 키보드 단축키 정의
├── assets/
│   └── icon-*.png                 # Add-in 아이콘 (16/32/80px)
├── webpack.config.js              # 빌드 설정
├── package.json
└── README.md                      # 설치 가이드
```

---

## Verification Plan

### Automated Tests
1. **로컬 Sideload 테스트**
   - `npm start` → Excel Desktop에 자동 Sideload
   - 셀 선택 → 우클릭 → "CellFocus로 열기" 확인
   - Ctrl+Shift+F 단축키 동작 확인

2. **동기화 테스트**
   - 셀 값 변경 → Add-in 그리드 자동 업데이트 확인
   - Add-in 그리드 편집 → 원본 셀 업데이트 확인
   - 무한 루프 방지 확인 (enableEvents)
   - 쓰기 직전 원본 셀 값이 바뀐 경우 즉시 덮어쓰지 않고 conflict 상태가 되는지 확인

3. **탭 관리 테스트**
   - 다중 범위 탭 생성/삭제
   - 서로 다른 시트의 범위 동시 열기
   - 최대 탭 개수 도달 시 경고
   - Excel/TaskPane 재시작 후 localStorage 탭 복원 + 이벤트 리스너 재등록 확인

4. **SharePoint 공동편집 테스트**
   - SharePoint/OneDrive에 저장된 `.xlsx`를 사용자 A/B가 Excel Desktop에서 동시에 열기
   - A가 CellFocus로 `A1:D10`을 열고, B가 같은 범위 셀을 원본 Excel에서 수정
   - A의 CellFocus 탭이 원격 변경을 감지해 최신 값으로 갱신하는지 확인
   - A가 오래된 값을 보고 같은 셀을 편집하려 할 때 silent overwrite가 아니라 conflict UI가 뜨는지 확인
   - A가 편집 중인 셀을 B가 수정했을 때 A의 입력 중 DOM이 사라지지 않는지 확인
   - AutoSave ON/OFF, 네트워크 일시 불안정 상태에서 상태바가 거짓 `Synced`를 표시하지 않는지 확인

### Manual Verification
- Excel Desktop (Windows) + SharePoint/OneDrive 저장 `.xlsx`를 1차 기준으로 테스트
- Excel Web은 보조 호환성 테스트로만 확인
- 대규모 범위 (1000+ 셀) 성능 테스트
- 다크/라이트 테마 전환 확인
