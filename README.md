
## 🥨Automation Anywhere 에서 사용위함
- AA에서 호출해서 사용했는데 AA Run Script로 사용할경우 기본적으로 Parameter가 하나 전달하는것을 확인함
- 그 파라미터는 파일자체 Path 를 전달함을 확인
- 그리고 AA Run Script 통해 호출된 VBScript파일 내부에서 다른 .vbs 파일을 실행한 경우 에러발생함..


## Files 설명

**1. getEmail.vbs**
- Outlook에서 이메일 읽어들임 
- 조건1. 제목에 특정 텍스트와 동일
- 조건2. 해당하는 폴더만 ex) RPA/MDG/MDG2004
- AA 에 해당메일 Body 내용들 Return 해줌 comma로 묶어서
- 메일내용 확인 후 삭제 (지운 편지함으로 이동, 코멘트 처리해둠)
- appendText 통해서 csv 파일 남기기

**2. appendText.vbs**
- 해당 pathTextFile 에 append로 텍스트 남김

**3. sendEmail.vbs**
- 파라미터 5개 To, CC, 제목, 내용, 첨부파일
- To, CC, 첨부파일은 세미콜론(;) 구분자 포함하여 여러개 인자 전달 가능

---

###### MEMO - 날짜 구분 가능, 첨부파일 다운 가능, 메일제목이나 본문 키워드 포함여부 가능, 읽음처리 가능, 삭제가능
- [MailItem object](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem)
- DateAdd
- email.Attachments.Count
- InStr(email.subject,"keyword")


## 참고
- [에러핸들링](https://stackoverflow.com/questions/157747/vbscript-using-error-handling)
