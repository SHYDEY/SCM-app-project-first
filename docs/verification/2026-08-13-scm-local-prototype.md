# SCM 로컬 프로토타입 검증 기록

## 검증 환경

- URL: `http://localhost:3000`
- 실행: `node server.mjs`
- 데이터 모드: 가상데이터
- 사용자 모델: 단일 사용자

## 자동 검증

`npm.cmd test` 실행 결과:

- 총 6개 테스트
- 통과 6개
- 실패 0개

검증한 항목:

- 7단계 워크플로 상태 생성
- 빈 작업공간의 0금액·0예외 상태
- 데모 수요 단계의 진행 상태
- 전체 화면 레지스트리
- 상태 JSON 직렬화
- 앱 셸의 전체 네비게이션 라벨

## 브라우저 검증

- 대시보드 화면 표시 확인
- 데모 금액·수량·예외 KPI 표시 확인
- 수요 확인 화면 이동 및 4개 샘플 행 표시 확인
- 보고·출력 화면 이동 및 JSON/업로드 파일 버튼 표시 확인
- 프로토타입 경계 문구 표시 확인

## 디자인 검증

- 화이트 기반 글로벌 헤더·상단 네비게이션 적용 확인
- 후지필름 그린 포인트(`#018463`) 적용 확인
- 블랙 타이포그래피와 넓은 여백 적용 확인
- 직선형 카드·얇은 보더·작은 라운드 적용 확인
- 대시보드 히어로, KPI, 워크플로, 예외 영역의 재배치 확인
- 주요 버튼의 다크 그레이 기본색과 그린 그라디언트 호버 규칙 적용
- 모바일 767px 기준 세로 스택·가로 스크롤 네비게이션 CSS 적용

## 다음 단계 승인 대기 범위

- 실제 `.xlsx` 파일 읽기 및 시트 매핑
- 업로드 검증 결과 저장
- 기준정보·BOM 기반 계산 엔진
- MOQ·Flex·필수 옵션·Common품 자동 검증
- 사장 보고용 실제 Excel 생성
- Supabase 저장 및 Vercel 배포
# UI interaction verification (2026-08-13)

- Demand source summary cards expose `data-detail` click actions.
- Clicking the OL submission card renders a `.detail-panel` with title, description, and detail values.
- The detail panel close button removes the panel without changing workflow state.
- `.sidebar` and `.topbar` report `position: sticky`; after scrolling, both remain at viewport top.
- Browser screenshot verified the detailed demand panel and fixed navigation at `http://localhost:3000`.
- Automated test result: 9 passed, 0 failed.
