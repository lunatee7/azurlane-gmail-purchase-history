# 개요

벽람항로의 결제내역을 Gmail을 통해 받아와 Google Sheets에 추가합니다.

# 사용 방법

Google 스프레드시트에서 확장 프로그램 - Apps Script를 사용하여 코드를 추가합니다.
<br>

### 시트 양식

- [1계정용](https://docs.google.com/spreadsheets/d/1Q4tSpspTzBAjndvBI2WACsjHvfCYaS-lMYnJeGVbaSs/edit?usp=sharing) (계정이 1개인 경우)
- [2계정용](https://docs.google.com/spreadsheets/d/1jchQEnI8grjWwFmCWn_hxJg3HQsCYYBvM748pmlI21U/edit?usp=sharing)   (벽람항로 계정이 2개인 경우, 본계정과 부계정의 결제내역을 관리하고 싶을 경우)

# 유의 사항

- 벽람항로 한국 서버에서 결제한 내역만을 추출합니다.
- 결제 내역 이메일을 소지하고 있는 Google 계정으로 스크립트를 실행하여야 합니다.
- Gmail에 결제 내역 이메일이 존재하지 않을 경우 사용할 수 없습니다.
- 2계정용 시트의 경우 본/부계정의 결제여부는 본인이 직접 입력해야 합니다.

# 업데이트 내역

- 0.1.0
  - 시간대 문제 수정
