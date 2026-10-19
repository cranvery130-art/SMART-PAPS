# SMART PAPS 앱 아이콘

양쌤스튜디오 앱 아이콘 가족(배구 경기분석·스코어 대시보드와 같은 틀)이에요. 따뜻한 종이색 바탕에
체육관 마루 띠와 코트 주황 선을 깔고, 가운데에 남색 초시계(주황 부분은 측정 중인 시간)를 그렸어요.
마루 위에는 PAPS 글자를 넣었어요. 브라우저 탭에 쓰는 16·32px 아이콘은 작은 크기에서도 알아볼 수
있도록 마루 띠 없이 초시계만 크게 넣었어요.
글자(Jua)는 필요한 글자만 담은 글꼴을 SVG 안에 넣어 두어서 따로 글꼴이 없어도 똑같이 보여요.

| 원본 SVG | 용도 | 만들어진 PNG (`public/icons/`) |
|---|---|---|
| `icon-any.svg` | 크롬 바로가기·바탕화면 | `icon-192.png`, `icon-512.png` |
| `icon-maskable.svg` | 안드로이드 마스크용(가장자리까지 배경, 내용은 안전영역 80% 안) | `icon-maskable-512.png` |
| `icon-full.svg` | 아이폰 홈 화면(iOS가 모서리를 직접 깎음) | `apple-touch-icon.png` (180) |
| `icon-small.svg` | 브라우저 탭 파비콘(초시계 단독) | `favicon-32.png`, `favicon-16.png` |

아이콘이 바뀌면 `index.html`·`public/manifest.webmanifest`의 아이콘 주소 끝 `?v=` 숫자를
하나 올려야 브라우저가 새 아이콘을 받아와요.
