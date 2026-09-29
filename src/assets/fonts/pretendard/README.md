# Pretendard (동적 서브셋) — 사이트에 포함한 글꼴

- 출처: npm `pretendard@1.3.9` (https://github.com/orioncactus/pretendard) — `dist/web/variable/` 의
  `pretendardvariable-dynamic-subset.css` + `woff2-dynamic-subset/*.woff2`(92개)를 **수정 없이** 복사.
- 라이선스: SIL Open Font License 1.1 — `LICENSE.txt`.
- 왜 동적 서브셋인가(D-024): 글자 범위(unicode-range)별로 잘게 나뉘어 있어 브라우저가 페이지에 실제로
  나오는 글자 조각만 내려받는다. 전체 가변 글꼴(약 2MB)을 한 번에 받는 것보다 모바일 첫 로딩이 가볍다.
- 구글 글꼴 서버 등 외부 서버를 부르지 않는다(D-022, D-020).
- 버전을 올릴 때: 같은 경로의 파일을 통째로 교체하고 이 문서의 버전을 고친다.
