---
name: drawio2pptx
description: Use when putting a draw.io diagram into PowerPoint, or when a diagram in a deck needs to stay editable. Converts .drawio files into individual native PPT objects - shapes, connectors, text boxes, icons - instead of one flat image. Triggers on "drawio를 ppt에", "구성도 ppt에 넣어줘", "다이어그램 슬라이드에", "put this diagram in a slide", "each element separately", "요소별로".
---

# drawio2pptx

draw.io 페이지를 요소별로 편집 가능한 PowerPoint 객체로 변환한다.

**이 변환을 직접 짜지 마라.** PNG로 붙여넣으면 편집이 불가능해지고, XML을 읽어 손으로 도형을 다시
쌓으면 라벨 위치와 연결선 경로가 틀어진다. CLI를 쓴다.

## 실행

```bash
drawio2pptx diagram.drawio                                  # -> diagram.pptx
drawio2pptx diagram.drawio --into deck.pptx --slide 2 --replace
drawio2pptx diagram.drawio --all-pages -o deck.pptx
```

명령이 설치되어 있지 않으면 이 레포에서 아무것도 설치하지 않고 바로 돌린다.

```bash
uv run --with python-pptx --with pillow python -m drawio2pptx diagram.drawio
```

draw.io 데스크톱이 필요하다(`brew install --cask drawio`). 도형 렌더링을 그쪽이 하기 때문이고,
없으면 도구가 설치 방법을 출력한다.

## 반드시 검증한다

변환은 기하 연산이라 오류가 눈에 보이는 형태로 나타난다. 돌리고 나서 직접 봐라.

```bash
drawio2pptx diagram.drawio --verify check.png
```

`check.png`는 세 단이다. 위는 draw.io 자체 export, 가운데는 저장된 `.pptx`를 다시 그려낸 것, 아래는
차이 맵이다. 아래 패널을 읽을 때: **글자와 선 테두리를 따라 나오는 얇은 윤곽선은 정상이다**(검증
렌더러의 텍스트 메트릭이 PowerPoint와 다르다). **면으로 찬 영역은 도형이 실제로 어긋난 것**이므로
원인을 찾아야 한다. 이모지는 검증 렌더러 폰트에 글리프가 없어 네모로 보이는데, pptx 자체는 멀쩡하다.

확인이 끝나면 CLI가 출력한 객체 개수를 보고하고, 검증은 눈으로 한 것임을 함께 밝힌다.

## 플래그 고르기

| 상황 | 플래그 |
| --- | --- |
| 한글, 일본어, 중국어 라벨 | `--ea-font "Apple SD Gothic Neo"`. 다른 PC에서 글자가 틀어지지 않게 |
| 슬라이드 가장자리까지 꽉 차면 안 될 때 | `--margin 0.04` |
| 4:3 덱 | `--slide-size 4:3` |
| 여러 페이지 다이어그램 | `--list-pages`로 먼저 확인하고 `--page N` 또는 `--all-pages` |
| 뭔가 이상할 때 | `--keep-workdir`로 중간 렌더를 남겨 확인 |

## 한계는 먼저 말한다

사용자가 나중에 발견하게 두지 말고 다음을 짚어준다.

- **스텐실 아이콘은 래스터로 남는다.** AWS, GCP, Cisco 스텐실은 PowerPoint 도형으로 표현할 수 없어서
  각각 개별 이미지로 들어간다. 옮기고 크기 조절은 되지만 색은 못 바꾼다.
- 곡선과 둥근 연결선은 꼭짓점만 남기고 직선화된다.
- 회전된 도형과 스윔레인은 처리하지 않는다.
- 압축 저장된 `.drawio`는 먼저 비압축 XML로 다시 저장해야 한다(draw.io: Extras → Edit Diagram에서
  *Compressed* 해제). 압축 파일을 만나면 도구가 알려준다.

## 민감한 다이어그램

아키텍처 구성도는 사내 자산인 경우가 대부분이다. 사용자의 실제 다이어그램을 공개 레포나 예제
디렉터리, 아티팩트에 먼저 물어보지 않고 복사하지 마라.
