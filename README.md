# SCM 수급·발주 관리 로컬 프로토타입

PRD 기준 1차 구현본입니다. 로그인 없이 한 명의 담당자가 가상데이터 또는 빈 작업공간으로 전체 발주 플로우를 확인할 수 있습니다.

## 실행

```powershell
npm start
```

브라우저에서 [http://localhost:3000](http://localhost:3000)을 엽니다.

## 이번 단계 범위

- 대시보드 중심 화면
- 가상데이터 시작 / 빈 작업공간 시작
- 데이터 업로드 화면과 업로드 이력 프로토타입
- 수요·재고·기기 발주·옵션 발주·예외·보고 화면
- JSON 백업 및 System 업로드용 프로토타입 CSV 다운로드
- localStorage 자동 저장

## 다음 단계 후보

- 실제 Excel `.xlsx` 파싱 및 업로드 검증
- 기준정보·BOM 기반 계산 엔진
- MOQ·Flex·필수 옵션·Common품 검증 자동화
- 사장 보고용 Excel 생성
- Supabase 저장 어댑터와 Vercel 배포

## 테스트

```powershell
npm test
```

# SCM-app-project-first
