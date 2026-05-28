# LeakTrendViewer

AutoBase HMI 기반 가스 장비의 Pa 단위 LeakRate2 시계열 데이터를 조회·시각화하고, 기존 ExcelReport 흐름 대비 빠른 자동 보고서 생성을 지원하는 WPF 데스크톱 도구입니다.

## 기능

- **Access DB 조회**: `.mdb` 파일에서 LeakRate2(Pa) 데이터 조회 (ACE/Jet OleDb provider fallback, 잠김 시 임시 복사 재시도, 동적 스키마 탐색)
- **시계열 차트**: OxyPlot 로그 스케일 축 (0 이하 leak 값은 log floor 치환)
- **날짜/시간 필터링**: 범위 선택 + 빠른 선택 (오늘, 어제, 최근 7일 등) — 실제 사용 피드백 반영
- **데이터 내보내기**: CSV / Excel (EPPlus) 지원
- **컬럼 선택**: 필요한 컬럼만 선택해서 조회·내보내기 (화면·파일 순서 일관 유지)
- **CLI 자동 내보내기**: `--auto-export` 모드 — AutoBase 버튼에서 GUI 없이 직접 호출 가능
- **비동기 처리**: MDB 로딩·Excel export 비동기 + DataGrid virtualization으로 긴 시계열 조회 시 UI 멈춤 최소화

## 실행 환경

- Windows 10/11
- .NET 9.0
- OLEDB 드라이버 (ACE 또는 Jet)

## 빌드 및 실행

```bash
dotnet build
dotnet run
```

## 사용법

### GUI 모드
1. 앱 실행 후 데이터 폴더 선택 (기본: `C:\B_MilliData`)
2. 날짜/시간 범위 설정
3. "조회" 버튼 클릭
4. 차트에서 데이터 포인트 클릭하면 그리드와 연동

### CLI 자동 내보내기 모드

```bash
LeakTrendViewer.exe --auto-export --start "2024-01-01 00:00" --end "2024-01-31 23:59"
```

## 기술 스택

- **WPF** (.NET 9.0) — UI 프레임워크 (STA thread 기반 CLI export 지원)
- **OxyPlot.Wpf** — 로그 스케일 시계열 차트
- **EPPlus** — Excel 파일·차트 생성
- **System.Data.OleDb** — Access `.mdb` 직접 접근 (ACE/Jet provider 자동 fallback)

## 프로젝트 구조

```
NEWLeakTrendViewer/
├── Models/          — LeakRecord, ColumnInfo (데이터 모델)
├── Services/        — MdbDataLoader, MdbFileResolver, ExcelExporter
├── MainWindow.xaml  — 메인 UI (빠른 날짜 선택 버튼, DataGrid virtualization)
├── App.xaml.cs      — CLI --auto-export 진입점 (STA thread)
└── docs/            — 문서 및 스크립트
```

## 배경

기존 AutoBase ExcelReport 흐름은 긴 시계열 데이터를 행 단위로 처리해 시간·자원 소모가 컸습니다. 본 도구는 같은 `.mdb` 데이터를 OleDb로 직접 읽고, OxyPlot 로그 축으로 필요한 구간만 렌더링하며, EPPlus로 자동 보고서를 생성합니다. 빠른 날짜 선택 버튼·CLI 자동 내보내기·컬럼 일관성 유지는 실제 사용 과정에서 도출된 개선입니다.

## 라이선스

개인 프로젝트
