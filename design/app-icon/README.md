# SMART PAPS 앱 아이콘

앱 스플래시 화면을 그대로 옮긴 디자인입니다. 검정 바탕에 골드 빛번짐, 은메달·금 트로피·은 상장
엠블럼(앱에서 쓰는 Lucide 아이콘과 같은 선 모양)과 그 아래 흰 글씨 PAPS. 브라우저 탭에 쓰는
16·32px 아이콘은 작은 크기에서도 알아볼 수 있도록 트로피만 크게 넣었습니다.
글자는 모두 도형(path)으로 바꿔 두어서 폰트가 없어도 똑같이 보입니다.

| 원본 SVG | 용도 | 만들어진 PNG (`public/icons/`) |
|---|---|---|
| `icon-any.svg` | 둥근 모서리, 크롬 바로가기·바탕화면 | `icon-192.png`, `icon-512.png` |
| `icon-maskable.svg` | 안드로이드 마스크용(가장자리까지 배경, 내용은 안전영역 80% 안) | `icon-maskable-512.png` |
| `icon-full.svg` | 아이폰 홈 화면(iOS가 모서리를 직접 깎음) | `apple-touch-icon.png` (180) |
| `icon-small.svg` | 브라우저 탭 파비콘(트로피 단독) | `favicon-32.png`, `favicon-16.png` |

사용 폰트: Archivo 700 — PAPS. 아이콘이 바뀌면 `index.html`·`public/manifest.webmanifest`의
아이콘 주소 끝 `?v=` 숫자를 하나 올려야 브라우저가 새 아이콘을 받아옵니다.
