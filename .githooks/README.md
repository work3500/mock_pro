# Git 훅 (.githooks)

이 디렉터리는 저장소에 **버전 관리되는 Git 훅**을 담고 있습니다.
기본 `.git/hooks` 는 저장소에 커밋되지 않으므로, 팀원 간 공유를 위해
훅을 이 디렉터리에서 관리하고 `core.hooksPath` 로 연결합니다.

## 설치

아래 중 한 가지 방법으로 활성화합니다.

```bash
# 방법 1: 설치 스크립트 실행
./.githooks/install.sh

# 방법 2: 수동 설정
git config core.hooksPath .githooks
```

설정은 로컬 저장소(`.git/config`)에만 적용되므로, 저장소를 새로
클론한 사람은 위 명령을 한 번 실행해야 합니다.

## 포함된 훅

### `pre-commit`

커밋 직전에 **스테이징된 추가 내용**을 검사하여 민감정보가
저장소에 올라가는 것을 막습니다. 다음과 같은 패턴을 탐지합니다.

- AWS 액세스 키 (`AKIA...`)
- 개인 키(PEM) 블록 (`-----BEGIN ... PRIVATE KEY-----`)
- Slack / GitHub / Google API 토큰
- 하드코딩된 `password`, `secret`, `api_key`, `token` 등

민감정보가 감지되면 커밋이 차단됩니다.

### 검사 건너뛰기

오탐이 명확한 경우에만 다음 방법으로 우회할 수 있습니다.

```bash
SKIP_HOOKS=1 git commit -m "..."
# 또는
git commit --no-verify -m "..."
```

## 비활성화

```bash
git config --unset core.hooksPath
```
