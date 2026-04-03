# 유상사급 타처보관 확인서 생성 프로그램

PySide6 기반 데스크톱 앱으로 사업장별 `자재유형별수불집계` 파일을 읽어 거래처별 재고자산확인서 Excel/PDF를 생성한다.

## 실행

```powershell
python -m pip install -e .
python -m outsourced_inventory_confirmation
```

## 빌드

```powershell
pwsh .\packaging\build.ps1
```

- 실행 파일은 `dist\유상사급타처보관`에 생성된다.
- Inno Setup으로 `packaging\installer.iss`를 열어 설치 파일을 만들 수 있다.
