# auto_project_manager (과제관리자동화) 매출전표 다운로더 모듈
<br/>
<br/>

## 매출전표다운로더
### 필요 라이브러리 설치
* ```pip install -r requirement.txt```
<br/>
<br/>

### exe 파일로 빌드
* ```make```
<br/>
<br/>

### 사용방법
* 설정.json 파일에서 다운로드 경로와 받고자 하는 카드를 지정
* 빌드한 경우 dist 폴더에서 <b>관리자권한</b>으로 매출전표다운로더.exe 파일을 더블클릭하여 실행
* 빌드하지 않은 경우 CMD를 켜서 과제관리 폴더로 이동한 후 "python 매출전표다운로더.py" 로 실행
<br/>
<br/>

### 주의사항
* 관리자 권한으로 실행 요망
* 설정.json 파일을 수정할때에 notepad++ 사용하길 바라며, (CR LF), UTF-8 형식임을 확인 요망
* 다른 엑셀파일이나 매트릭스가 열려있는 경우 비정상 작동하는 경우가 존재
