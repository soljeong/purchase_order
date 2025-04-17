# purchase_order
구매의뢰 -> 발주 관리

- 새 레포 만들기
- 데브컨테이너 설정 추가 (clasp 설치)
- 코드스페이스 실행
- clasp 로그인
```
clasp login
```

- 이미 작성하고 있던 프로젝트를 clone 하기, 이미 작성해놓은 내용을 가져온다
- 앱스크립트의 scriptId, 이게 프로젝트를 부르는 이름인가
```
clasp clone 18tbKmogGD3Bb7dw5vVj-Y16eGzu1MRyiuZ-7lyzNNMgl8rVihJYV9lkK
```

- 이제 수정하면 push 가능할지
```
clasp push
clasp push -w
clasp open-container
clasp open-script
```

## HtmlService.createTemplateFromFile

- 앱스크립트에서 html 템플릿 이용할 수 있을지
- 시트의 데이터를 템플릿에 뿌려서 모달로 템플릿 표시

## 프로젝트
- 이슈를 먼저 만들고
- 커밋을 할 때 커밋 메시지에 해당 이슈의 번호를 넣어서
- 이슈와 커밋을 연결지을 수 있다?


## 발주서
- 문서에 표시되는 날짜를 어떤 기준으로 할지?