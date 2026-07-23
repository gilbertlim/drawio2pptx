# drawio2pptx

draw.io 다이어그램을 통이미지가 아니라 **요소 하나하나가 편집되는 PowerPoint 객체**로 넣습니다.

박스는 박스로, 화살표는 화살표로, 글자는 글자로 들어갑니다. 서버 아이콘 하나만 옮기거나, 영역 색을
바꾸거나, 라벨 오타를 고치는 데 draw.io로 돌아가 다시 export할 필요가 없습니다.

![왼쪽은 draw.io, 오른쪽은 같은 다이어그램이 개별 PowerPoint 객체로 들어간 모습](examples/preview.png)

<sub>오른쪽 패널은 저장된 `.pptx`를 이 도구가 다시 그려낸 것입니다([결과 확인](#결과-확인) 참고).
PowerPoint 스크린샷이 아닙니다.</sub>

실행 방법은 두 가지입니다. Claude Code에 말로 시키거나, 파이썬 CLI를 직접 부르거나. 하는 일은 같으니
지금 작업하는 방식에 맞는 쪽을 고르면 됩니다.

---

## 두 방법 공통: draw.io 데스크톱 설치

도형 렌더링은 draw.io가 합니다. 이건 우회할 방법이 없습니다. PowerPoint는 mxGraph 스텐실을 표현할
수단이 없고, 그걸 제대로 그려주는 다른 도구도 없습니다.

```bash
brew install --cask drawio            # macOS
sudo snap install drawio              # Linux
winget install JGraph.Draw            # Windows
```

일반적인 설치 경로는 자동으로 찾습니다. 특이한 위치에 있다면 `DRAWIO_BIN` 환경변수를 쓰거나
`--drawio /경로/drawio`를 넘기세요.

---

# 방법 1 — 프롬프트로 실행

이미 Claude Code에서 작업 중이고, 다이어그램 삽입이 보고서 덱을 만드는 큰 흐름의 한 단계일 때 좋습니다.

### 최초 1회 설정

```bash
git clone https://github.com/gilbertlim/drawio2pptx && cd drawio2pptx
uv tool install .                                          # drawio2pptx 명령을 PATH에 등록
ln -s "$PWD/.claude/skills/drawio2pptx" ~/.claude/skills/drawio2pptx
```

심볼릭 링크를 걸면 어느 프로젝트에서든 스킬이 잡힙니다. 이 레포 안에서 작업할 때는 링크가 없어도
`.claude/skills/`에서 자동으로 인식합니다.

### 그다음엔 그냥 말하면 됩니다

> 이 구성도 ppt에 넣어줘

> architecture.drawio를 보고서 덱 2번 슬라이드에 넣어줘, 원래 내용은 지우고

> 페이지별로 슬라이드 하나씩 만들어줘

Claude가 [스킬 문서](.claude/skills/drawio2pptx/SKILL.md)를 읽고 플래그를 고른 뒤 변환하고, 완료라고
말하기 전에 결과를 눈으로 확인합니다. 스킬에는 한계를 먼저 밝히라는 지시와, 사용자의 다이어그램을
공개 레포나 공유 아티팩트에 함부로 복사하지 말라는 규칙도 들어 있습니다.

결과로는 객체 개수가 돌아오고, 검증을 요청했다면 `check.png`도 함께 나옵니다.

```
report.pptx  (slide 2, 104 objects: 10 shapes, 22 connectors, 39 text boxes, 33 icons)
```

---

# 방법 2 — 파이썬으로 실행

스크립트에 엮거나, 여러 파일을 한 번에 돌리거나, 어떤 플래그가 실행됐는지 정확히 보고 싶을 때 좋습니다.

### 설치

```bash
git clone https://github.com/gilbertlim/drawio2pptx && cd drawio2pptx
pip install .
```

시스템 파이썬을 건드리기 싫으면 `uv tool install .` 또는 `pipx install .`을 쓰세요. 셋 중 하나만 하면
됩니다.

아무것도 설치하지 않고 클론에서 바로 돌려볼 수도 있습니다.

```bash
uv run --with python-pptx --with pillow python -m drawio2pptx examples/sample.drawio
```

### 잘 되는지 먼저 확인

레포에 예제가 들어 있으니 이걸로 한 번 돌려보세요. 아래 명령들에 나오는 `diagram.drawio`는 실제 파일이
아니라 **자리표시자**입니다. 본인 파일 경로로 바꿔서 쓰세요.

```bash
drawio2pptx examples/sample.drawio -o hello.pptx
# hello.pptx  (slide 1, 31 objects: 3 shapes, 8 connectors, 11 text boxes, 9 icons)
```

### 명령줄

```bash
# diagram.drawio -> diagram.pptx, 16:9 슬라이드 한 장. 이게 전부입니다.
drawio2pptx diagram.drawio

# 출력 파일명 지정
drawio2pptx diagram.drawio -o deck.pptx

# 이미 있는 덱의 2번 슬라이드에 넣기, 그 슬라이드 기존 내용은 지우고
drawio2pptx diagram.drawio --into deck.pptx --slide 2 --replace

# 여러 페이지짜리 다이어그램을 페이지당 슬라이드 하나로
drawio2pptx diagram.drawio --all-pages -o deck.pptx

# 변환하고, 제대로 나왔는지 눈으로 볼 수 있게
drawio2pptx diagram.drawio --verify check.png
```

| 플래그 | 용도 |
| --- | --- |
| `--page N` / `--all-pages` / `--list-pages` | 여러 페이지 다이어그램 |
| `--into DECK --slide N [--replace]` | 새로 만들지 않고 기존 덱에 삽입 |
| `--slide-size 16:9 \| 4:3 \| 16:10 \| auto \| 13.333x7.5` | 새 덱의 슬라이드 크기 |
| `--margin 0.04` | 슬라이드를 꽉 채우지 않고 여백 두기 |
| `--ea-font "Apple SD Gothic Neo"` | CJK 폰트 고정. 한글, 일본어 라벨이 다른 PC에서 틀어지지 않게 |
| `--font Arial` | 라틴 폰트 강제 지정 |
| `--scale 8` | 아이콘 해상도 상향 (기본 6, 슬라이드 폭 기준 약 370dpi) |
| `--drawio PATH` | 자동 탐지가 실패할 때 |
| `--keep-workdir` | 뭔가 이상할 때 중간 렌더 결과를 남김 |

나머지는 `drawio2pptx --help`에 있습니다.

### 라이브러리로 사용

```python
from drawio2pptx import convert

result = convert("diagram.drawio", "deck.pptx", slide_size="16:9", margin=0.03)
print(result.path, result.counts)
# deck.pptx {'rect': 10, 'picture': 28, 'line': 22, 'text': 39}
```

`convert()`는 출력 경로, 기록한 슬라이드 번호, 객체 개수, 다이어그램 좌표계 기준 콘텐츠 경계를 담은
`Result`를 돌려줍니다.

---

## 실제로 뭐가 나오나

| draw.io | PowerPoint |
| --- | --- |
| 일반 사각형과 컨테이너 | 네이티브 도형. 채우기, 테두리, 점선 그대로 |
| 라벨 | 진짜 텍스트 상자. 폰트, 크기, 굵기, 색 유지 |
| 연결선 (꺾인 경로 포함) | 자유형 커넥터. 원본 화살촉 그대로 |
| AWS / GCP / Cisco / Veeam 스텐실 | 아이콘당 고해상도 PNG 하나, 픽셀 단위로 배치 |
| 내장 이미지 | 원본 해상도로 추출, 크롭 반영 |

래스터로 남는 건 스텐실 아이콘 하나뿐입니다. 각각 개별 이미지라 옮기고 크기를 바꿀 수는 있지만,
PowerPoint에서 색을 바꾸지는 못합니다. 나머지는 전부 네이티브입니다.

## 결과 확인

변환 오류는 눈으로 보이는 종류입니다. 그러니 보세요.

```bash
drawio2pptx diagram.drawio --verify check.png
```

`check.png`는 세 단으로 나옵니다. 위는 draw.io 자체 export, 가운데는 저장된 `.pptx`를 다시 그려낸 것,
아래는 차이 맵입니다.

아래 패널은 규칙 하나만 기억하고 보면 됩니다. 글자와 선 테두리를 따라 나오는 얇은 윤곽선은 정상입니다.
검증 렌더러의 텍스트 배치가 PowerPoint와 다르고, 앞으로도 같아지지 않을 것이기 때문입니다. **면으로
꽉 찬 영역이 보이면 도형이 실제로 잘못 놓인 것입니다.**

## 동작 원리

세 가지가 전부를 떠받칩니다.

**프레임 사각형이 좌표계를 고정합니다.** draw.io는 export할 때 항상 그래프 경계에 맞춰 잘라내기 때문에,
같은 다이어그램을 두 번 렌더해도 원점이 다르고 시트에서 아이콘 하나를 안정적으로 오려낼 수 없습니다.
그래서 export 직전에 다이어그램 복사본에 보이지 않는 사각형을 끼워 넣고, SVG에서 그게 실제로 어디에
그려졌는지 다시 읽습니다. 그 읽은 값이 원점입니다. 라벨이 프레임 밖으로 튀어나와 캔버스가 커져도
매핑이 깨지지 않는 이유가 이것입니다.

**라벨 위치와 연결선 경로는 draw.io의 SVG export에서 `data-cell-id`로 읽어옵니다.** mxGraph의 라벨
배치와 직교 라우터를 다시 구현할 가치는 없고, SVG에 이미 답이 들어 있습니다.

**겹치지 않는 스텐실은 렌더를 공유합니다.** draw.io를 한 번 띄우는 데 몇 초씩 듭니다. 겹침 여부로
묶었더니 아이콘 19개짜리 다이어그램이 19번에서 2번으로, 약 75초에서 26초로 줄었습니다.

## 한계

- 곡선과 둥근 연결선은 꼭짓점만 남기고 직선화됩니다.
- 스텐실 아이콘 색은 PowerPoint에서 바꿀 수 없습니다.
- 회전된 도형과 스윔레인 컨테이너는 아직 처리하지 않습니다. 오류가 나는 게 아니라 잘못된 위치에
  놓입니다.
- 압축 저장된 `.drawio`는 먼저 비압축 XML로 다시 저장해야 합니다(draw.io: **Extras → Edit Diagram**에서
  *Compressed* 해제). 압축 파일을 만나면 도구가 알려줍니다.

## 개발

```bash
pip install -e ".[dev]"
pytest -q
ruff check .
```

종단 테스트는 draw.io 데스크톱이 필요하고 없으면 스스로 건너뜁니다. 파싱과 좌표 계산 테스트는 어디서든
돕니다. 코드를 고치기 전에 [CLAUDE.md](CLAUDE.md)에 정리된 불변 조건을 먼저 보세요.

렌더러를 손봤다면 README 이미지를 다시 만드세요.

```bash
drawio2pptx examples/sample.drawio -o /tmp/s.pptx --verify /tmp/check.png
python tools/make_preview.py /tmp/check.png examples/preview.png "PowerPoint — 31 separate objects"
```

## 라이선스

MIT
