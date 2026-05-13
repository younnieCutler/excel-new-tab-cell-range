# CellFocus 고객사 배포 가이드

CellFocus v1의 기본 배포 모델은 AppSource 공개 출시가 아니라 **고객사 테넌트 소유 HTTPS 호스팅 + Microsoft 365 Admin 중앙 배포**입니다.

## 운영 원칙

- 사용자는 SharePoint/OneDrive의 `.xlsx`를 **Windows Excel Desktop**에서 엽니다.
- CellFocus는 브라우저 앱이 아니라 Excel 창 안의 Office Add-in TaskPane으로 열립니다.
- 벤더 서버는 사용하지 않습니다. 고객사가 Add-in 정적 파일과 manifest를 자기 테넌트/인프라에서 운영합니다.
- CellFocus 자체 backend, 외부 API, telemetry는 없습니다.
- Office.js 런타임은 Microsoft Office Add-in 플랫폼의 필수 의존성입니다.

## 고객사 배포 산출물 생성

```bash
npm install
npm run build:customer -- --base-url https://cellfocus.customer.example
```

지원 페이지 URL이 별도로 있으면 다음처럼 지정합니다.

```bash
npm run build:customer -- \
  --base-url https://cellfocus.customer.example \
  --support-url https://intranet.customer.example/cellfocus/support
```

생성 결과:

```text
dist/
├── manifest.xml
├── taskpane.html
├── commands.html
├── shortcuts.json
├── taskpane.bundle.js
├── commands.bundle.js
├── taskpane.css
└── assets/
```

## 고객사 호스팅

고객사는 `dist/` 전체를 고객사 소유 HTTPS 정적 호스팅에 배포합니다.

가능한 호스팅 예:

- Azure Static Web Apps
- Azure Storage static website + HTTPS front door
- 고객사 IIS/NGINX 정적 호스팅
- 사내 포털의 정적 파일 호스팅

필수 조건:

- `https://.../taskpane.html` 접근 가능
- `https://.../commands.html` 접근 가능
- `https://.../shortcuts.json` 접근 가능
- `https://.../assets/icon-*.png` 접근 가능
- 인증이 필요한 사내 호스팅을 쓰는 경우 Excel Desktop의 WebView에서 접근 가능해야 함

## Microsoft 365 Admin 중앙 배포

고객사 관리자는 `dist/manifest.xml`을 사용해 조직에 Add-in을 배포합니다.

권장 흐름:

1. Microsoft 365 admin center 접속
2. Integrated apps 또는 Office Add-ins 배포 메뉴로 이동
3. `dist/manifest.xml` 업로드
4. 배포 대상 사용자 또는 그룹 선택
5. Windows Excel Desktop에서 Add-in 노출 확인
6. SharePoint/OneDrive `.xlsx`를 Excel Desktop에서 열고 CellFocus 실행 확인

## 검증

```bash
npm run validate:customer
```

수동 검증:

- Windows Excel Desktop에서 Add-in이 리본/컨텍스트 메뉴에 표시되는지 확인
- SharePoint `.xlsx`를 Desktop 앱으로 열고 선택 범위를 CellFocus로 열 수 있는지 확인
- 공동편집 중 원격 변경이 CellFocus에 반영되는지 확인
- 충돌 가능성이 있는 쓰기에서 silent overwrite가 발생하지 않는지 확인

## AppSource를 v1 기본 경로로 쓰지 않는 이유

AppSource/Office Store 출시는 공개 Marketplace 심사, 퍼블리셔 등록, 공개 지원/개인정보 링크, 공개 HTTPS 호스팅 운영이 필요합니다. 이 모델은 "고객사 환경에서 서버 및 유지 책임을 고객사에 전가"하는 요구와 맞지 않습니다.

고객사가 조직 내부에서만 쓰는 제품이면 Microsoft 365 Admin 중앙 배포가 더 적합합니다.
