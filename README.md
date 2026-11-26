# KAHI 메일 도우미 (Outlook Add-in)

공유 메일함 **sunnyahn@coreatech.kr** 의 메일을 열면 오른쪽 작업창에 그 메일의 **국문 번역 · 요약 · 확인 필요사항(회신 판단·할 일·기한)** 과 크리에이터 정보를 보여 주는 읽기 전용 Outlook 추가 기능.
번역·요약은 구글 드라이브에 수집된 메일을 Apps Script `KAHI 메일 AI`(Claude API)가 미리 만들어 두고, 이 화면은 웹 앱 조회 주소(`?a=mail&id=<internetMessageId>&k=<조회 키>`)로 결과를 읽기만 한다.

- 권한 `ReadItem` — 메일을 읽기만 하고 수정·이동·발송하지 않는다.
- 공유 메일함 지원 `SupportsSharedFolders`, 창 고정 `SupportsPinning`
- 비밀 정보 없음: 연결 코드(웹 앱 주소 + 조회 키)는 각자 작업창에 한 번 붙여 넣으면 Outlook 계정 설정에 저장된다. 이 저장소에는 키가 들어가지 않는다.

## 배포
1. 이 저장소를 **Public** 으로 만들고 `main` 에 올린다.
2. **Settings → Pages → Build and deployment → Source: GitHub Actions** 를 한 번 선택한다.
3. `main` 에 올릴 때마다 Actions 가 화면 파일과 `manifest.xml` 을 배포한다.
   - Outlook 등록 주소: `https://<계정>.github.io/kahi-mail-addin/manifest.xml`

## Outlook 에 등록
- 본인: Outlook 웹 → https://aka.ms/olksideload → 내 추가 기능 → 사용자 지정 추가 기능 → URL에서 추가 → 위 manifest.xml 주소
- 팀 4명: Microsoft 365 관리 센터 → 설정 → 통합 앱 → 사용자 지정 앱 업로드 → 매니페스트 URL → 대상 사용자 지정

## 파일
| 파일 | 역할 |
|---|---|
| taskpane.html / .js / .css | 작업창 화면 |
| help.html | 사용 안내 |
| manifest.template.xml | Outlook 등록 파일 원본 (`__BASE__` → 배포 주소) |
| .github/workflows/pages.yml | Pages 배포 + manifest.xml 생성 |

화면을 고친 뒤 `main` 에 올리면 1~2분 안에 반영된다. 매니페스트(버전·권한)를 바꿀 때만 Outlook 에 다시 등록한다.
