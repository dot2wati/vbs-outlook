
## Automation Anywhere 에서 사용위함
## Files 설명

1. getEmail.vbs
    - Outlook에서 이메일 읽어들임 
    - 조건1. 제목에 특정 텍스트와 동일
    - 조건2. 해당하는 폴더만 ex) RPA/MDG2004
    - AA 에 해당메일 Body 내용들 Return 해줌 comma로 묶어서
    - 메일내용 확인 후 삭제 (지운 편지함으로 이동, 코멘트 처리해둠)

2. appendLog.vbs
    - 해당 pathLogFile 에 append로 로그 남김 UTF-8

3. sendEmail.vbs
    - 파라미터 5개 제목, 내용, To, CC, 첨부파일
    - To, CC, 첨부파일은 모두 ;로 구분함. 여러개 인자 전달 가능
---

###### MEMO (날짜 구분 가능, 첨부파일 다운 가능, 메일제목이나 본문 키워드 포함여부 가능, 읽음처리 가능, 삭제가능)
- [MailItem object](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem)
- DateAdd
- email.Attachments.Count
- InStr(email.subject,"keyword")