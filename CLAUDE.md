# drawio2pptx 작업 가이드

draw.io 페이지 하나를 개별 PowerPoint 객체로 변환하는 도구다. 무엇을 하는지는 README에 있고,
이 문서는 **뭘 건드리면 조용히 깨지는지**를 다룬다.

## 실행

```bash
pip install -e ".[dev]"
pytest -q                 # 40개. 종단 테스트는 draw.io 데스크톱 없으면 스스로 건너뜀
ruff check .
drawio2pptx examples/sample.drawio -o /tmp/s.pptx --verify /tmp/check.png
```

이 머신에 설치된 실행 파일은 `uv tool install .`로 넣은 것이다. 소스를 고친 뒤에는
`uv tool install --reinstall .`을 돌려야 PATH의 `drawio2pptx`가 갱신된다.

## 모든 것이 얹혀 있는 불변 조건 하나

draw.io는 PNG든 SVG든 항상 그래프 경계에 맞춰 잘라낸다. 내용이 바뀌면 원점이 따라 움직이고, 같은
다이어그램을 두 번 렌더해도 좌표계가 공유되지 않는다. 그래서 시트에서 아이콘 하나만 오려내는 게
그 자체로는 불가능하다.

해법: `model.with_frame()`이 export 직전에 다이어그램 복사본으로 보이지 않는 사각형(`__d2p_frame`)을
끼워 넣는다. 그리고 `svgmap.SvgMap.frame_rect()`가 **그 사각형이 실제로 그려진 위치**를 읽어오는데,
요청한 프레임 값이 아니라 이 읽은 값이 원점이다. 가정하지 않고 되읽기 때문에 라벨이 프레임 밖으로
넘쳐 캔버스가 커져도 매핑이 깨지지 않는다.

모든 요소의 좌표는 **SVG 좌표계**에 산다. 셀 박스는 모델에서 draw.io 좌표로 들어와 `build.collect()`
에서 이동되고, 라벨과 연결선 경로는 처음부터 SVG 공간으로 들어온다. 이 둘을 섞으면 전체가 프레임
패딩만큼 조용히 어긋나는데, 테스트는 경계만 검사하므로 잡지 못한다. `--verify`를 돌려야 한다.

## 알아내는 데 시간을 쓴 것들

- **draw.io CLI의 `-p`는 1-based다.** 0을 넘기면 조용히 1페이지가 나오므로, off-by-one이 오류가 아니라
  "엉뚱한 도형이 렌더됨"으로 나타난다.
- **이미지 크롭은 별도 스타일 키에 있다.** `image=data:...` URI 안이 아니라 `clipPath=inset(위% 오른쪽%
  아래% 왼쪽%)`이다. 놓치면 크롭된 로고가 작고 어긋나게 들어간다. `build.extract_bitmap()`이 처리한다.
- **스타일에 `fillColor`/`strokeColor`가 없으면 draw.io 기본값(흰 채움, 검정 테두리)이다.** 없는 걸
  "none"으로 읽으면 도형이 통째로 사라지고 라벨만 남는다. 기본값을 코드에 다시 구현하지 말고
  `svgmap.paint()`로 SVG가 실제 칠한 값을 읽어라. 모서리 반경(`rx`)도 거기서 같이 나온다.
- **색상은 `light-dark(밝은값, 어두운값)` 형태로 들어온다.** 앞쪽을 써야 한다. 색 파싱은
  `svgmap.normalize_color()` 한 곳에서만 해야 한다.
- **HTML 라벨은 안쪽 선언이 이긴다.** 바깥 div가 `font-weight: bold`라도 안쪽
  `<span style="font-weight: normal">`이 덮어쓴다. 그래서 `_label_from_foreign_object()`는 굵기, 크기,
  색, 폰트를 모두 **마지막** 매치로 잡는다.
- **python-pptx의 도형과 자유형은 테마 `effectRef`를 상속한다.** 모든 박스와 커넥터에 드롭 섀도가
  붙는다는 뜻이다. `_Emitter._no_theme_effects()`가 막고 `test_no_shape_inherits_a_theme_shadow`가
  지킨다. `add_shape`나 `build_freeform`을 새로 부르는 코드는 똑같이 처리해야 한다.
- **PowerPoint 화살촉은 선 두께에 비례한다.** draw.io는 두께와 무관하게 고정 크기로 그린다. 1px 선에
  `headEnd`/`tailEnd`를 붙이면 촉이 안 보일 만큼 작아지고, 짧은 선은 아예 사라진다. 그래서 SVG에서
  draw.io가 그린 화살촉 폴리곤을 읽어 `custGeom`의 닫힌 채움 서브패스로 넣는다(`_rebuild_with_arrowheads`).
  한 도형 안에 넣어야 선과 촉이 같이 움직인다. 서브패스를 추가하면 도형 경계도 다시 잡아야 하고,
  안 그러면 촉이 잘린다.
- **`a:ln`의 자식 순서는 스키마로 강제된다.** solidFill, prstDash, headEnd, tailEnd 순서다. 어기면
  PowerPoint가 "복구하시겠습니까"를 띄우는 파일이 나온다.
- **draw.io 앱 창이 떠 있으면 CLI export가 그 뒤에 물려 멈춘다.** user-data-dir을 잡고 있기 때문이다.
  테스트가 갑자기 300초 타임아웃으로 죽으면 십중팔구 이것이니, 코드를 의심하기 전에 `pgrep -fl draw.io`
  부터 보라. `run()`이 타임아웃을 잡아 이 힌트를 출력한다.
- **이 머신에서는 PowerPoint와 Keynote의 AppleScript 자동화가 막혀 있다.** `open`은 성공하는데 열린
  문서가 0개다. 여기에 시간 쓰지 말 것. `verify.py`가 존재하는 이유가 이것이다.

## 레이어 묶기

`stencils.plan_layers()`는 겹치지 않는 도형들을 draw.io 한 번 실행에 몰아넣는다. 실행 한 번에 몇 초씩
들기 때문이다. 불변 조건은 **같은 레이어 안에서 어떤 셀의 크롭 박스도 다른 셀이 그린 픽셀에 닿으면
안 된다**는 것이다. 닿으면 크롭이 옆 도형을 물고 들어온다.

`aws4.group`이 까다로운 경우다. 셀 박스는 컨테이너 전체인데 실제로 칠해지는 건 왼쪽 위 25px 남짓의
배지뿐이다(테두리는 네이티브로 그리고, 렌더 복사본에서는 `strokeColor`를 `none`으로 바꾼다).
`crop_box()`가 그룹 셀을 `GROUP_ICON_EXTENT`로 줄여서, 컨테이너 안에 있는 도형들과 충돌하지 않게
한다. 이 상수를 키우면 자식 도형을 물기 시작한다.

## verify.py는 의도적으로 근사치다

저장된 `.pptx`를 PIL로 다시 그려서 도형이 엉뚱한 자리에 갔는지 잡는 용도다. PowerPoint를 흉내 내려는
게 아니고, 텍스트 메트릭이 다르다. 차이 패널에 글자와 선 테두리를 따라 나오는 얇은 윤곽선은 정상이고
문제없다. 면으로 찬 영역이 나오면 그건 실제로 어긋난 것이다. 윤곽선 노이즈를 "고치려" 들지 말 것.
그 길 끝에는 텍스트 레이아웃 엔진 재구현이 있다.

## 새 도형 타입 지원 추가

`stencils.needs_render()`가 draw.io에게 그리게 할지, 네이티브 객체로 만들지를 가른다. PowerPoint가
정확히 표현할 수 있을 때만 네이티브로 보내라. 채우기와 테두리, 점선이 있는 사각형 정도다. 나머지는
렌더한다. 회전된 도형과 스윔레인은 미지원이고, 실패하는 대신 잘못된 위치에 놓인다. 이쪽을 건드린다면
`collect()`에 분기와 테스트를 함께 추가할 것.

## 남의 실제 다이어그램을 커밋하지 말 것

`examples/sample.drawio`는 일부러 일반적인 내용이다(DMZ, WAS, DBMS). 실무에서 나온 아키텍처 구성도는
사내 자산이고 규제기관 제출용인 경우도 있다. 이 레포에도, 테스트 픽스처에도, 공개 아티팩트에도 들어가면
안 된다. 버그 재현에 실제 다이어그램이 필요하면 gitignore된 `fixtures-private/` 아래에 두고 쓴다.
