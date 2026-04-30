# 생기부 특기사항 자동 검토기

PyQt5 기반 오프라인 전용 Windows 데스크톱 앱입니다.  
학교생활기록부 서술형 항목을 일괄 분석해 Byte 초과, 금지/주의 표현, 대체 권고 표현을 제공합니다.

## 핵심 특성
- 오프라인 전용(외부 API/클라우드 호출 없음)
- 원본 Excel 미변경, 결과 파일 별도 저장
- `xlsx`, `xlsm` 지원(매크로 실행하지 않음)
- 대용량 분석을 위한 QThread 기반 비동기 처리와 취소 기능

## 설치
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 실행
```bash
python main.py
```

## 테스트
```bash
pytest -q
```

## 샘플 입력 파일 생성
```bash
python scripts/generate_sample_excel.py
```

## PyInstaller 빌드
```bash
pyinstaller --noconfirm --clean build/app.spec
```

## 설정 파일 위치
- 기본 설정 원본: `config/*.json`
- 실행 시 사용자 설정 복사본: `data/config/*.json`

## 규칙 편집 방법
- `data/config/forbidden_rules.json`: 금지/주의 표현
- `data/config/suggestion_rules.json`: 대체 권고 사전
- `data/config/allowlist.json`: 허용 예외 용어
- `data/config/category_rules.json`: 카테고리/문맥 규칙

## 출력 컬럼
- `byte_count`
- `byte_limit`
- `overflow_yn`
- `verdict`
- `hit_terms`
- `suggested_rewrite`
- `review_note`

## 결과 시트
- `요약`
- `위반목록`
- `검토필요목록`
- `설정 스냅샷`

## 스크린샷 예시 설명
- 좌측 패널: 파일/시트/헤더/검토열/항목유형 선택
- 중앙 표: 원본 미리보기와 셀 선택
- 우측 패널: 선택 셀 상세 분석 근거
- 하단 탭: Byte 분석, 금지표현, 대체표현, 설정, 로그

## 제한 사항
- `.xls`는 미지원
- 문맥 판별은 규칙 기반 최소 분석이며 완전한 의미 이해 모델은 아님
- 매우 큰 파일에서 메모리 사용량 증가 가능
