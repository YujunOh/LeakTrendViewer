# LeakTrendViewer

반도체 장비의 Leak Rate 데이터를 시각화하고 분석하는 WPF 데스크톱 앱입니다.

## 기능

- **Access DB 조회**: .mdb 파일에서 LeakRate2 데이터 조회
- **시계열 차트**: OxyPlot 기반 로그 스케일 차트
- **날짜/시간 필터링**: 범위 선택 + 빠른 선택 (오늘, 어제, 최근 7일 등)
- **데이터 내보내기**: CSV / Excel 형식 지원
- **컬럼 선택**: 필요한 컬럼만 선택해서 조회
- **CLI 자동 내보내기**: `--auto-export` 모드로 자동화 가능

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

- **WPF** (.NET 9.0) — UI 프레임워크
- **OxyPlot** — 차트 라이브러리
- **EPPlus** — Excel 파일 생성
- **OleDb** — Access DB 연결

## 프로젝트 구조

```
NEWLeakTrendViewer/
├── Models/          — 데이터 모델 (LeakData, etc.)
├── Services/        — DB 로더, CSV/Excel 내보내기
├── MainWindow.xaml  — 메인 UI
├── App.xaml         — 앱 설정
└── docs/            — 문서 및 스크립트
```

## 라이선스

개인 프로젝트
