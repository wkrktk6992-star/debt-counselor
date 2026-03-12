# 채무조정 상담도구 - 핸드오프 문서

## 프로젝트 개요
- **위치**: `C:\Users\dks20240502\Desktop\ai 출력물\debt-counselor\`
- **배포**: Vercel 자동배포 (GitHub push → `https://debt-counselor.vercel.app`)
- **GitHub**: `https://github.com/wkrktk6992-star/debt-counselor.git`
- **구조**: 단일 HTML 파일 (`index.html`, ~1289줄) + 정적 배포
- **Preview 서버**: `.claude/launch.json`에 설정됨 (`npx serve` port 3456)

---

## 현재 상태 (2026-03-12, 업데이트)

### 완료된 작업
1. **ExcelJS 템플릿 기반 엑셀 내보내기로 전환**
   - 기존: SheetJS + AOA(배열) 방식 → 서식 전혀 안 맞음 (사용자 불만)
   - 변경: ExcelJS + 원본 xlsx 템플릿 base64 임베드 → 원본 서식 유지
   - CDN: `https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js`
   - 템플릿 base64: `TEMPLATE_B64` 상수 (56,744자, index.html line 1260)

2. **제도 선택 UI 추가**
   - 분석 결과 패널에 "최종 제도 선택" 드롭다운 추가
   - 옵션: 개인파산·면책 / 개인회생 / 새출발기금 / 신용회복위원회
   - 분석 실행 시 최우선 추천 제도가 자동 선택됨

3. **상담내용/SOLUSION/판정의견 자동생성**
   - `generateNarrative(d, programId)` — 상담일지 D15 영역 내러티브 (가/나/다 구조)
   - `generateSolusion(programId)` — 평가표 C55
   - `generateVerdict(programId)` — 평가표 C64
   - textarea로 표시되어 사용자가 직접 수정 가능

4. **ExcelJS `exportExcel()` async 함수**
   - 원본 템플릿 로드 → 셀 값만 채움 → 다운로드
   - 상담일지 매핑: D5(성명), H5(나이), I5(생년월일), H7(연락처), D9(주소), B14/D14(상담일/내용), D15(내러티브), D39-D45(채무), H45(재산), A54/C54/E54(소득), C57/E57(가족), F60-F62(특기사항), C64(진술)
   - 평가표 매핑: C5(이름), E5(생년월일), G5(전화), C7(사업상태), G14/G22/G27/G38/G40/G43/G48(후속프로세스), F23(정기소득), A56-A62(체크박스 ■/□), C55(SOLUSION), C64(판정의견)

5. **evalMap 버그 수정 (shin 키 추가)**
   - `evalMap`에 `shin` 키가 없어 신용회복위원회 선택 시 체크박스 미매핑 → 수정
   - 연체 기간에 따라 workout(90일+) / preworkout(31~90일) / quick(31일 이하)로 분기
   - `generateSolusion()`, `generateVerdict()`도 하위 제도명 반영

6. **셀 좌표 검증 완료**
   - Python openpyxl 검증 스크립트로 46개 셀 전수 검증 (25 상담일지 + 21 평가표)
   - Merged ranges 보존 확인 (상담일지 86개, 평가표 67개)
   - 서식(font/border) 보존 확인
   - 결과: 46 PASS / 0 FAIL

7. **Git 커밋 & 푸시 완료**

### 수동 브라우저 테스트 필요
- 브라우저에서 예시 데이터 입력 → 분석 → **4개 제도 각각 선택**하여 내러티브 확인
- 엑셀 내보내기 → Excel에서 열어 서식/값 확인
- 특히 신용회복위원회 선택 시 연체 기간별 하위 제도 분기 확인

---

## 핵심 파일

| 파일 | 용도 |
|------|------|
| `index.html` | 메인 앱 (단일 파일, HTML+CSS+JS) |
| `_clean_template.xlsx` | 데이터 셀 비운 엑셀 템플릿 (base64 원본) |
| `_template_b64.txt` | 위 파일의 base64 인코딩 (JS에 임베드됨) |
| `_patch_html.py` | 임시 - base64를 HTML에 삽입한 스크립트 |
| `.claude/launch.json` | Preview 서버 설정 |

---

## 주요 함수 구조 (index.html)

```
doCalc()          → 입력값 수집 + 4개 제도 점수 계산
renderResult(d)   → 결과 카드/비교표/상세/상담일지 렌더 + 제도 선택 UI
onProgramSelect() → 드롭다운 변경 시 내러티브/SOLUSION/판정의견 재생성
generateNarrative(d, programId) → 상담일지 상담내용 텍스트 (가/나/다 구조)
generateSolusion(programId)     → 평가표 SOLUSION 텍스트
generateVerdict(programId)      → 평가표 판정의견 텍스트
exportExcel()     → async, ExcelJS로 템플릿 로드 → 셀 채움 → xlsx 다운로드
doReset()         → 전체 초기화
```

---

## 원본 엑셀 템플릿 위치
```
G:\공유 드라이브\02. 소상공인 법률자문 및 채무조정 용역\
  2025_소상공인 법률지원 용역\2. 채무조정\2-3. 개인회생\
  채무조정_21645_이향지[완]\채무조정_21645_이향지.xlsx
```
- 7개 시트: 신청서, 상담일지(64행), 평가표(63행), 신용위안내문, 결과보고서, 온라인신청사항, 유효성
- 상담일지: 84개 merged range, 평가표: 83개 merged range
- 현재 상담일지 + 평가표 2개 시트만 데이터 채움

---

## 다음 단계 (우선순위순)

### 1. 엑셀 내보내기 실제 테스트
```
- 브라우저에서 예시 데이터 입력 → 분석 → 제도 선택 → 엑셀 내보내기
- 다운로드된 xlsx를 Excel에서 열어 확인:
  · 서식(병합, 테두리, 색상) 원본과 동일한지
  · 데이터가 올바른 셀에 들어갔는지
  · 한글 깨짐 없는지
```

### 2. 셀 좌표 보정
```
- 상담일지와 평가표의 정확한 셀 좌표 재검증
- merged cell 영역에 값을 쓰면 ExcelJS가 에러를 뱉거나 무시할 수 있음
- 필요시 Python openpyxl로 원본 분석:
  python -c "
  from openpyxl import load_workbook
  wb = load_workbook('원본.xlsx')
  ws = wb['상담일지']
  for m in ws.merged_cells.ranges:
      print(m)
  "
```

### 3. Git 커밋 & 푸시
```bash
cd "C:\Users\dks20240502\Desktop\ai 출력물\debt-counselor"
echo "_clean_template.xlsx" >> .gitignore
echo "_template_b64.txt" >> .gitignore
echo "_patch_html.py" >> .gitignore
echo "_make_clean_template.py" >> .gitignore
git add index.html .gitignore
git commit -m "ExcelJS 템플릿 기반 엑셀 내보내기 + 제도 선택 UI + 상담내용 자동생성"
git push origin main
```

### 4. (선택) 추가 개선
- 나머지 5개 시트(신청서, 신용위안내문 등)도 데이터 채우기
- 상담내용 내러티브를 더 정교하게 (변호사 실제 작성 스타일 참고)
- Claude API 연동하여 AI 기반 내러티브 생성 (더 자연스러운 문체)

---

## 주의사항
- `index.html`에 base64 템플릿(56KB)이 임베드되어 있어 파일이 큼 (~110KB 총)
- ExcelJS CDN이 로드 안 되면 엑셀 내보내기 전체 불가 → 오프라인 사용 시 CDN 파일 로컬 번들 필요
- `exportExcel()`은 `async` 함수 — 버튼 onclick에서 호출 시 정상 동작하지만 `await` 없이 호출되므로 에러 catch는 함수 내부에서 처리
- 사용자는 "기존 엑셀 파일과 병합이나 모양이 전부 다르다"고 강하게 불만 표시했으므로, 서식 일치가 최우선 품질 기준
