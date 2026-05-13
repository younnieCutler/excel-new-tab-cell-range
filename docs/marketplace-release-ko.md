# CellFocus 공개 배포 가이드

CellFocus를 공개 배포하려면 Microsoft Marketplace/AppSource 제출을 기준으로 준비합니다. 이 경로는 고객사 중앙 배포와 다릅니다.

## 공개 배포 원칙

- 사용자는 Microsoft Marketplace 또는 Office 안의 Add-ins UI에서 CellFocus를 설치합니다.
- Add-in의 HTML/JS/CSS는 공개 HTTPS 호스팅에 올라가야 합니다.
- `manifest.xml`에는 공개 접근 가능한 `SourceLocation`, 아이콘 URL, `shortcuts.json`, `SupportUrl`이 들어가야 합니다.
- Partner Center 계정, Marketplace 등록 정보, 스크린샷, 공개 지원 페이지, 개인정보 처리방침이 필요합니다.
- 공개 배포 후에도 Excel Desktop TaskPane 방식은 유지합니다. 독립 네이티브 창으로 바뀌지 않습니다.

## 공개 배포 산출물 생성

GitHub Pages로 공개 배포할 경우:

```bash
npm run build:github-pages
```

이 저장소의 기본 공개 URL:

```text
https://younnieCutler.github.io/excel-new-tab-cell-range/
```

`main` 브랜치에 push하면 GitHub Actions가 `dist/`를 `gh-pages` 브랜치에 배포합니다.

다른 공개 HTTPS 호스팅을 사용할 경우:

```bash
npm install
npm run build:marketplace -- \
  --base-url https://cellfocus.example.com \
  --support-url https://cellfocus.example.com/support.html
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
├── support.html
├── privacy.html
└── assets/
```

`dist/` 전체를 공개 HTTPS 호스팅에 배포한 뒤 `dist/manifest.xml`을 Partner Center 제출에 사용합니다.

## 검증

일반 manifest 검증:

```bash
npm run validate:customer
```

Marketplace production-level 검증:

```bash
npm run validate:marketplace
```

`validate:marketplace`는 Microsoft 검증 서비스에 접근하므로 네트워크가 필요합니다. 또한 manifest의 URL에 Microsoft 검증 서비스가 직접 접근하므로, `dist/`를 실제 공개 HTTPS 호스팅에 먼저 배포한 뒤 실행해야 합니다.

## Partner Center 제출 전 체크리스트

- Partner Center 계정 준비
- 공개 HTTPS 호스팅 준비
- `dist/manifest.xml`의 모든 URL이 공개 접근 가능한지 확인
- 지원 페이지 공개
- 개인정보 처리방침 공개
- 앱 설명, 검색 키워드, 카테고리, 아이콘, 스크린샷 준비
- Windows Excel Desktop에서 SharePoint `.xlsx`를 열고 TaskPane 동작 확인
- Excel on the web / Mac 등 manifest가 노출하는 플랫폼에서 최소 실행 가능 여부 확인

## 주의사항

Microsoft Marketplace 검증은 manifest에 선언된 플랫폼 기준으로 동작 가능성을 봅니다. 현재 manifest는 Excel on Windows 외 플랫폼도 가능하다고 분석될 수 있으므로, 실제 공개 제출 전에는 Windows 외 환경에서 기능이 깨지지 않는지 확인해야 합니다.

Office.js Add-in 특성상 공개 배포를 해도 서버가 사라지지 않습니다. Marketplace는 설치/발견 채널이고, Add-in UI 파일은 계속 공개 HTTPS 호스팅에서 제공해야 합니다.
