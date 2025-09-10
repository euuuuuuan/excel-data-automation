# 📊 Excel Automation Tool for Product Management

## 💡 프로젝트 개요
이 프로젝트는 **엑셀 파일에 분산되어 있는 상품 정보를 자동으로 통합**하고, 특정 기준에 따라 상품의 **'전체숨김여부' 상태를 업데이트**하는 자동화 도구입니다.  
복잡한 수작업을 없애고 **데이터 관리 효율을 극대화**하여, 상품 재고 및 상태를 정확하게 파악할 수 있도록 돕습니다.  

특징적으로, **엑셀의 원본 서식(셀 색상, 글꼴, 열 너비 등)을 그대로 유지**하면서 필요한 데이터만 정확하게 수정하고, 원본 파일 손상 없이 **새로운 결과 파일을 생성**합니다.  
또한, 실제 민감한 개인정보 및 회사 대외비 정보는 포함되지 않도록, **더미 데이터 생성 스크립트**(`create_dummy.py`)를 통해 안전하게 포트폴리오화하였습니다.  

---

## ⚙️ 주요 기능
- **상품 코드 추출**  
  여러 월별 정산내역 파일(예: `4월~7월 정산내역`)을 자동 스캔하여 모든 시트에서 고유한 상품코드를 추출  

- **데이터 통합 및 비교**  
  추출된 상품코드를 `상품리스트` 파일과 비교하여 정산내역에 포함된 상품을 식별  

- **'전체숨김여부' 상태 업데이트**  
  정산내역에 존재하는 상품코드를 상품리스트 파일에서 찾아, 해당 상품의 상태를  
  `on` → **`off`** 로 자동 변경  

- **원본 서식 보존**  
  `openpyxl`을 활용해 엑셀 파일의 원본 양식(셀 색상, 폰트, 정렬 등)을 그대로 유지  

- **안전한 파일 관리**  
  원본 파일은 절대 수정하지 않고, 모든 변경사항이 적용된 **새로운 결과 파일을 생성**하여 데이터 손실 방지  

---

## 🛠 기술 스택
[![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)](https://www.python.org/)  
[![pandas](https://img.shields.io/badge/pandas-150458?style=flat&logo=pandas&logoColor=white)](https://pandas.pydata.org/)  
[![openpyxl](https://img.shields.io/badge/openpyxl-1F6FEB?style=flat&logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)  
[![glob](https://img.shields.io/badge/glob-000000?style=flat&logo=python&logoColor=white)](https://docs.python.org/3/library/glob.html)  

---

## 🚀 사용법

### 1. 파일 준비
- `더미_상품리스트.xlsx` 파일을 프로젝트 폴더에 위치  
- `더미_*월_정산내역.xlsx` 형식의 월별 정산내역 파일들을 동일한 폴더에 위치  

### 2. 의존성 설치
```
pip install pandas openpyxl
```
3. 실행
```
python test.py
```

4. 결과
성공적으로 실행되면,
➡️ 더미_상품리스트_업데이트.xlsx 파일이 새로 생성됩니다.
<img src="https://github.com/euuuuuuan/excel-data-automation/blob/main/docs/screenshots/teminal_excute_result.png" alt="결과 이미지">

📝 코드 설명
glob.glob()
지정된 패턴(더미_*월_정산내역.xlsx)에 맞는 모든 파일 경로를 수집

pd.read_excel()
각 월별 파일을 DataFrame으로 로드 (헤더 위치를 정확히 지정해 열 이름 인식 보장)

openpyxl.load_workbook()
상품리스트 파일을 Workbook 객체로 불러와 원본 서식 유지

set()
추출된 상품코드를 중복 없이 효율적으로 관리

ws.cell()
특정 셀에 직접 접근하여 값 수정

wb.save()
원본 파일은 보존하고, 새로운 파일명으로 저장

🔒 데이터 보안
본 프로젝트는 실제 데이터가 아닌 **더미 데이터(create_dummy.py)**를 사용하여 제작되었습니다.




