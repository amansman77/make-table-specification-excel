# 테이블 명세서 생성기 (feat. Excel)

엑셀 형식으로 테이블 명세서를 작성하는 스크립트입니다.

## 사용법

1. `template-sample.xlsx` 와 같이 명세서를 위한 템플릿을 지정합니다.
2. config.py에 다음 사항들을 설정합니다.
- FILE['file_name'] : output을 위한 파일명
- DATABASE : 명세서 작성을 위한 데이터베이스 정보
- CELL_INFO['template_sheet_name'] : 템플릿 시트명
- CELL_INFO['TABLE_ENGLISH_NAME'] : 테이블 명(영문)
- CELL_INFO['TABLE_KOREAN_COMMENT'] : 테이블 명(한글)
- CELL_INFO['TABLE_COMMENT'] : 테이블 설명
- CELL_INFO['START_COLUMN_INDEX'] : 컬럼정보의 시작위치(이하 row의 수는 충분히 확보해야합니다.)
- CELL_INFO['COLUMN_NUMBER'] : 컬럼의 순번 위치
- CELL_INFO['COLUMN_NAME'] : 컬럼명 위치
- CELL_INFO['DATA_TYPE'] : 데이터 타입 위치
- CELL_INFO['DATA_LENGTH'] : 데이터 길이 위치
- CELL_INFO['IS_NULLABLE'] : NULL 여부 위치
- CELL_INFO['COLUMN_KEY'] : 컬럼 키 위치
- CELL_INFO['EXTRA'] : 기타 정보 위치
- CELL_INFO['START_INDEX_INDEX'] : 인덱스정보의 시작위치(이하 row의 수는 충분히 확보해야합니다.)
- CELL_INFO['INDEX_NUMBER'] : 인덱스의 순번 위치
- CELL_INFO['INDEX_NAME'] : 인덱스명 위치
- CELL_INFO['COLUMN_NAME'] : 인덱스를 건 컬럼 위치
- CELL_INFO['INDEX_TYPE'] : 인덱스 유형
- CELL_INFO['COLUMN_NAME_2'] : 인덱스를 건 컬럼 위치
