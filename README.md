# sCO2 Raw Data to FEM-ready Workbook

초임계 CO₂ 실험에서 수집한 RTD, SYS, HX 데이터를 하나의 Excel 데이터셋으로 병합하고,
매크로 템플릿을 이용해 계산이 완료된 `.xlsm` 결과 파일을 만드는 Windows용 자동화 도구입니다.

## 주요 기능

- RTD 센서 수를 자동으로 판별하고 T/t/R 데이터 묶음을 정렬합니다.
- RTD 측정 시간 범위에 맞춰 SYS 및 HX 데이터를 필터링합니다.
- 실험 조건별 `*_Merged.xlsx` 파일을 생성합니다.
- `MergedData` 시트의 실제 데이터 범위만 읽어 `exp!T4`부터 한 번에 입력합니다.
- `Batch_Clear_Data`의 MsgBox/InputBox를 해당 Excel 프로세스의 HWND로 식별해 처리합니다.
- `calculator` 매크로 실행 전후로 입력 데이터가 유지되는지 확인합니다.
- 저장한 XLSM을 같은 Excel 인스턴스에서 다시 열어 최종 데이터를 검증합니다.
- 파일별 성공, 실패, 충돌 건너뜀 결과를 마지막에 요약합니다.

## 실행 환경

- Windows
- Python 3.10 이상 권장
- Microsoft Excel 데스크톱 앱
- PyCharm 또는 PowerShell/명령 프롬프트

Excel COM 자동화와 Windows 창 제어를 사용하므로 macOS 및 Linux에서는 XLSM 후처리를
실행할 수 없습니다. `run_xlsm_automation=False`로 설정하면 Excel 자동화 없이
`Merged.xlsx` 생성 단계까지만 사용할 수 있습니다.

## 설치

저장소를 복제한 뒤 프로젝트 폴더에서 다음 명령을 실행합니다.

```powershell
git clone https://github.com/shPark-1205/sCO2_raw_to_FEM-ready.git
cd sCO2_raw_to_FEM-ready
python -m pip install -r requirements.txt
```

직접 패키지를 설치하려면 다음 명령을 사용할 수 있습니다.

```powershell
python -m pip install pywin32 pywinauto openpyxl pandas numpy
```

## 입력 폴더 구성

`CONFIG["base_dir"]`은 한 실험 날짜/조건의 최상위 원본 폴더를 가리켜야 합니다.

```text
base_dir/
├─ _SYS*.xlsx
├─ RTD/
│  ├─ ..._RTD-<실험조건>-r1mm/
│  │  └─ *.xlsx
│  ├─ ..._RTD-<실험조건>-r2mm/
│  │  └─ *.xlsx
│  └─ ...
└─ HX/
   └─ *<실험조건>*.xlsx
```

RTD 하위 폴더 이름에서 `_RTD-`와 `-r숫자mm` 사이의 문자열을 실험 조건으로 사용합니다.
같은 실험 조건의 RTD 폴더들은 하나의 그룹으로 처리됩니다.

## 설정

실행 전 [main.py](main.py)의 `CONFIG`를 환경에 맞게 수정합니다.

| 항목 | 설명 |
|---|---|
| `base_dir` | RTD, HX, SYS 원본이 있는 최상위 폴더 |
| `xlsm_template_path` | VBA 매크로가 포함된 `.xlsm` 템플릿 |
| `xlsm_output_dir` | 완성된 XLSM을 저장할 최상위 폴더 |
| `run_xlsm_automation` | `False`이면 Merged.xlsx까지만 생성 |
| `overwrite_existing_xlsm` | 기존 XLSM 덮어쓰기 여부. 기본값은 `False` |
| `excel_visible` | 자동화 중 Excel 창 표시 여부 |
| `batch_clear_confirmation` | 삭제 확인 InputBox에 입력할 문자 |
| `dialog_timeout_seconds` | 각 VBA 팝업의 최대 대기시간 |
| `dialog_diagnostic_logging` | 발견한 창과 자식 컨트롤의 상세 로그 출력 여부 |
| `calculation_timeout_seconds` | calculator 이후 계산 완료 최대 대기시간 |
| `batch_clear_macro` | 데이터 초기화 매크로 이름 |
| `calculator_macro` | 계산 매크로 이름 |
| `chiller_setting_temperature_c` | `exp!E5`에 입력할 Chiller 설정온도 |
| `pump_input_hz` | `exp!E6`에 입력할 Pump 주파수 |
| `heater_temperature_c` | `exp!E8`에 입력할 Heater 온도 |
| `environment_temperature_c` | `exp!E9`에 입력할 환경온도 |

`overwrite_existing_xlsm=False`이면 같은 이름의 결과가 이미 있을 때 해당 파일은
변경하지 않고 건너뜁니다.

## 파일명 규칙

Merged 파일은 다음 형식을 사용합니다.

```text
YYYYMMDD_<압력>MPa_<온도>C_<Re>_<전압>V_Hd<H/d>_Merged.xlsx
```

예시:

```text
20260915_7.771MPa_31.3C_40000_8V_Hd6.00_Merged.xlsx
```

이 파일은 다음 경로와 이름의 XLSM으로 변환됩니다.

```text
xlsm_output_dir/
└─ 20260915_7.771MPa_31.3C_40000/
   └─ 20260915_7.771MPa_31.3C_40000_8V_Hd6.00.xlsm
```

형식이 일치하지 않거나 날짜가 유효하지 않으면 값을 추측하지 않고 해당 파일을
건너뜁니다. 전압과 H/d에는 정수 또는 소수를 사용할 수 있습니다.

## 처리 과정

### 1. Merged.xlsx 생성

1. RTD 파일에서 측정 시작 및 종료 시각을 구합니다.
2. RTD 센서 열을 T, t, R 묶음으로 나눕니다.
3. 센서가 10개보다 적으면 각 묶음에 빈 열을 추가해 열 배치를 맞춥니다.
4. SYS와 HX 데이터를 RTD 시간 범위로 필터링합니다.
5. 결과를 `MergedData` 시트에 값으로 저장합니다.

### 2. XLSM 생성 및 계산

각 Merged 파일은 다음 순서로 하나씩 처리됩니다.

1. XLSM 템플릿을 최종 출력 경로로 복사합니다.
2. `DispatchEx("Excel.Application")`로 독립 Excel 인스턴스를 만듭니다.
3. `Batch_Clear_Data`를 실행하고 세 개의 팝업을 순서대로 처리합니다.
4. 파일명 및 CONFIG 값을 `exp` 시트에 입력합니다.
5. `MergedData` 값을 `exp!T4`부터 2차원 배열로 일괄 입력합니다.
6. 입력값을 다시 읽고 `CountA`와 대표 셀 다섯 개를 검증합니다.
7. `calculator`를 실행하고 입력 데이터가 유지되는지 다시 확인합니다.
8. 저장 후 XLSM을 read-only로 재오픈해 디스크 기록 결과를 최종 확인합니다.
9. workbook과 Python이 만든 Excel 인스턴스를 종료합니다.

## exp 시트 입력 위치

| 셀/범위 | 입력값 |
|---|---|
| `B1` | 파일명에서 추출한 H/d 문자열 |
| `E3` | 파일명의 날짜를 `YY.MM.DD`로 변환한 문자열 |
| `E4` | 파일명에서 추출한 전압 |
| `E5` | Chiller 설정온도 |
| `E6` | Pump 입력 주파수 |
| `E8` | Heater 온도 |
| `E9` | 환경온도 |
| `T4` 이후 | MergedData 전체 값 범위 |

## 실행

```powershell
python main.py
```

정상 실행 시 파일마다 다음과 같은 단계 로그가 표시됩니다.

```text
[XLSM 1/3] 처리 시작: ...
  [1/6] 템플릿 복사 완료
  [2/6] Batch_Clear_Data 실행 및 팝업 처리 완료
  [3/6] exp 기본값 입력 완료
  [4/6] MergedData 범위 할당 및 read-back 검증
  [5/6] calculator 실행 및 계산 완료
  [6/6] 저장 및 Excel 종료 완료
```

## 데이터 무결성 검증

COM 범위 할당이 예외 없이 반환됐다는 사실만으로 성공 처리하지 않습니다.

- 원본의 비어 있지 않은 셀 수를 계산합니다.
- 첫 값, 25%, 50%, 75%, 마지막 값에 해당하는 대표 셀을 검사합니다.
- 입력 직후, 임시 저장 후, calculator 후, 최종 저장 후에 값을 다시 읽습니다.
- source와 destination의 `CountA` 차이가 허용 범위인지 확인합니다.
- 파일을 닫고 read-only로 다시 열어 같은 검사를 반복합니다.

검증에 실패한 XLSM은 성공 수에 포함하지 않습니다. 조사할 수 있도록 불완전한 복사본은
삭제하지 않고 콘솔에 경로를 출력합니다.

## Excel 보안 설정

매크로 실행이 차단되면 Excel에서 다음 메뉴로 이동합니다.

```text
파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 신뢰할 수 있는 위치 → 새 위치 추가
```

`xlsm_output_dir`을 신뢰할 수 있는 위치로 추가하고 필요하면 하위 폴더도 신뢰하도록
설정합니다. 프로그램은 매크로 보안 수준을 변경하거나 보안 정책을 우회하지 않습니다.

## 팝업 처리 문제 확인

`dialog_diagnostic_logging=True`이면 다음 정보가 출력됩니다.

- 감시 중인 Excel PID와 메인 HWND
- 팝업의 제목, 클래스 이름과 HWND
- 각 자식 컨트롤의 클래스, 텍스트와 HWND
- `WAIT_CONFIRM → WAIT_INPUT → WAIT_COMPLETION → COMPLETED` 상태 전이

Excel과 PyCharm의 실행 권한이 다르면 Windows가 UI 메시지를 차단할 수 있습니다.
이 경우 두 프로그램을 모두 일반 권한으로 실행하거나 둘 다 같은 권한 수준으로 실행합니다.

## 테스트

Excel을 실행하지 않는 단위 테스트:

```powershell
python -m unittest -v
```

문법 검사:

```powershell
python -m py_compile main.py test_main.py
```

단위 테스트는 파일명 파싱, 출력 경로, 실제 데이터 범위 판별, 내부 빈 셀 유지,
시간 문자열과 Excel 시간 일련값 비교, CountA/대표 셀 검증을 확인합니다.

## 주의 사항

- 자동화가 실행되는 동안 생성된 Excel 창과 VBA 팝업을 수동으로 조작하지 마세요.
- 여러 XLSM을 병렬 처리하지 마세요. 프로그램은 의도적으로 한 파일씩 처리합니다.
- 템플릿에는 `exp` 시트와 설정된 이름의 VBA 매크로가 있어야 합니다.
- 실행 중 강제 종료하면 Excel 프로세스가 남을 수 있으므로 가능하면 `Ctrl+C`로 중단하세요.
- 신뢰할 수 있는 위치에는 출처를 확인한 파일만 저장하세요.
