# vsb_outlook

## Automation Anywhere 에서 사용위함
---

##### Files 설명

1. getEmail.vbs
    - Outlook에서 이메일 읽어들임 
    - 조건1. 제목에 특정 텍스트 포함 ex) "MDG2004"
    - 조건2. 해당하는 폴더만 ex) RPA/MDG2004
    - AA 에 해당메일 Body 내용들 Return 해줌 comma로 묶어서

2. appendLog.vbs
    - 해당 pathLogFile 에 append로 로그 남김 UTF-8