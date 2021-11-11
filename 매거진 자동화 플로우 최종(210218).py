from __future__ import print_function
import pickle
import os.path
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
import urllib.error as errors
from datetime import datetime
from apiclient import errors
from typeform import Typeform

SPREADSHEET_ID = "#####" #Google Sheet 아이디, 실제로 사용할 것
SPREADSHEET_NAME = '시트1'

# SPREADSHEET_ID = "#####" #테스트용
# SPREADSHEET_NAME = '메인'

RANGE_NAME = SPREADSHEET_NAME+'!A:Q'
MANAGER_EMAIL = "#####" #관리자 계정
GOOGLE_EMAIL = "#####" #구글 api 사용하는 계정
CREDS_FILE = "#####" #구글 api credential 파일
TOGOFORM_DAY = 30 #마지막 포스트 얼마나 이후에 다음 투고 물어볼지 (새로운 행 만들기)
TOGOFORM_RE_DAY = 7 #투고 폼 언제 다시 보낼지 간격
ESTIMATED_DATE_LEFT = 14 #다음 투고 예정일 얼마나 남았을 때, 초고 입력 안내할지 
ESTIMATED_DATE_LEFT_NOT_SUMBITTED = 3 #초고 제출 마감일이 n일 남았는데, 아직 초고가 제출되지 않은 경우 초고 독촉 메일 날리기 => 이 때의 n
NOT_YET_CHECK_FILE = 5 #필진에게 수정 알림 메일 보냈는데, 일정기간동안 필진이 안읽을 때, 관리자에게 직접 확인 요청
POST_ALARM = 30 #원고 최종 완료 후, 이만큼 지났는데도 POST DATE와 URL이 업데이트가 되어있지 않으면 관리자에게 메일 보내기

#<To 관리자>
def makeDocu(name, email):
    message_text = (name+"님의 원고가 작성될 문서를 작성 후, Google Sheet에 입력해주세요")
    message = create_message(
        GOOGLE_EMAIL, email, "[김박사넷] 매거진 문서 추가 관련 메일", message_text)
    return message

def submittedFirst(name, DOCUMENT_ID,email):
    message_text = (name+"님의 투고가 완료되었습니다.\n"
                    +"확인 후, stage 서버에 업로드 해주시고, ‘매거진 필진 관리’ 스프레드 시트의 Stage URL, Last edited 부분을 업데이트해주세요."
                    +"\n\n원고 URL :"+DOCUMENT_ID)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 투고 완료 알림", message_text)
    return message    

def requestModify(name, DOCUMENT_ID, email):
    message_text = (name+"님이 수정을 요청하였습니다.\n"
                    +"수정을 진행하시고, ‘수정 완료’ 댓글을 재활성화 해주세요.\n"
                    +"또한, ‘매거진 필진 관리’ 스프레드 시트의 Last edited를 꼭 업데이트 해주세요.\n\n"
                    +"원고 URL : " +DOCUMENT_ID)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 수정 요청 알림", message_text)
    return message

def finalSubmitted(name, DOCUMENT_ID,email):
    message_text = (name+"님의 최종 원고가 제출되었습니다.\n"
                    +"확인 후, 매거진을 발행해주세요.\n\n"
                    +"원고 URL : "+DOCUMENT_ID)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 최종 원고 제출 알림", message_text)
    return message

def notReadMail(name, DOCUMENT_ID, to_email, email):
    message_text = (name+"님이 계속해서 검토를 하지 않고 있습니다. 확인해주세요.\n\n"
                    +"원고 URL : "+ DOCUMENT_ID
                    +"\n필진 email : "+to_email)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 수정 없음 알림", message_text)
    return message

def postMagazineAlarm(author,email,DOCUMENT_ID):
    message_text = (author+"님의 매거진이 제출되었지만, 아직 업로드 되지 않았습니다.\n 매거진을 업로드 해주세요\n"
                    +"원고 URL : " + DOCUMENT_ID)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 업로드 요청 알림", message_text)
    return message


#<To 필진>
def firstMail(name,document_Link, email):
    message_text = (name+"님!"+" 안녕하세요, 김박사넷입니다!\n"
                    +"다음 URL은 필진님의 원고가 작성될 문서 URL입니다.\n"
                    +"해당 URL에 접속하여, 안내문을 꼭 읽으신 후, 초고를 작성해주시기 바랍니다.\n"
                    +"혹시 안내문을 읽으신 후, 도움이 꼭 필요한 부분이 있으시다면, 메일로 문의주시기 바랍니다.\n\n"
                    +"원고 URL :" + document_Link)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 초안 관련 메일", message_text)
    return message

def notSubmittedFirst(name, document_Link, email, days):
    message_text = (name+"님!"+" 안녕하세요, 김박사넷입니다!\n"
                    +"매거진 초고 마감일이 "+str(days)+"일 남았는데, 아직 초고가 제출되지 않았습니다.\n"
                    +"다음 URL은 필진님의 원고가 작성될 문서 URL입니다.\n"
                    +"해당 URL에 접속하여, 안내문을 꼭 읽으신 후, 초고를 작성해주시기 바랍니다.\n"
                    +"혹시 안내문을 읽으신 후, 도움이 꼭 필요한 부분이 있으시다면, 메일로 문의주시기 바랍니다.\n\n"
                    +"원고 URL :" + document_Link)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 초안 미제출 관련 메일", message_text)
    return message


def finishedModify(name, stageURL, document_Link, email):
    message_text = (name+"님의 원고 수정본입니다.\n"
                    +"매거진 URL을 확인 후, 원고 URL에 접속하여 안내문을 꼭 읽으신 후, 수정을 진행해주시기 바랍니다.\n"
                    +"수정하실 부분이 없으셔도 원고 URL에 접속하여 안내문을 따라주셔야합니다.\n"
                    +"혹시 안내문을 읽으신 후, 도움이 꼭 필요한 부분이 있으시다면, 메일로 문의주시기 바랍니다.\n"
                    +"매거진 URL 확인 시 다음의 아이디와 비밀번호를 사용해주세요 \n -> 아이디 : magazinetester\n -> 비밀번호 : qwer135!#%\n"
                    +"\n매거진 URL : " + stageURL
                    +"\n원고 URL : " + document_Link)
    message= create_message(GOOGLE_EMAIL, email, "[김박사넷] 매거진 수정 완료 알림 메일", message_text)
    return message

def togoPlanForm(Author, Author_email):
    message_text = (Author+ "님! 안녕하세요, 김박사넷입니다!\n"
                    +Author+"님의 매거진이 그립습니다.\n"
                    +"혹시 다음 매거진 일정이 있으시다면 다음 링크를 통해서 알려주세요!\n"
                    + "링크 : https://lim310851.typeform.com/to/S3UcGwHN")
    message= create_message(GOOGLE_EMAIL, Author_email, "[김박사넷] 다음 매거진 일정에 대해서", message_text)
    return message

def postAlarmMail(Author, Author_email, postURL):
    message_text = (Author +"님! 안녕하세요, 김박사넷입니다!\n"
                    +Author+"님의 매거진이 업로드 되었습니다.\n"
                    +"반응을 확인해주세요!\n"
                    +"URL : " + postURL)
    message = create_message(GOOGLE_EMAIL, Author_email, "[김박사넷] 매거진 업로드 알림", message_text)
    return message


##########################################################해당 선의 밑 부분은 수정하지 마세요!#######################################################



def create_message(sender, to, subject, message_text): #GMail API : 메시지 만드는 함수
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  message['cc'] = 'admin@phdkim.net'
  return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode("utf8")}



def send_message(service, user_id, message): #GMail API : create_message에서 만든 메시지를 보내는 함수
  try:
    message = (service.users().messages().send(userId=user_id, body=message).execute())
    print('Message Id: %s' % message['id'])
    return message
  except (errors.HTTPError):
    print('An error occurred')


def retrieve_comments(service, file_id): #Google Drive의 word API : word 문서의 댓글들을 읽어오는 함수
  try:
    comments = service.comments().list(fileId=file_id, includeDeleted = False, fields='*').execute()
    comments = comments['comments']
    return comments
  except (errors.HttpError):
    print ('An error occurred')
  return None


def main():
    SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/documents.readonly', 'https://mail.google.com/']
    creds = None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)
    service2 = build('gmail', 'v1', credentials=creds)
    service3 = build('sheets', 'v4', credentials=creds)
    service4 = build('drive', 'v3', credentials=creds)

    ########GSheets 읽어오기##############
    sheet = service3.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range = RANGE_NAME).execute()
    values = result.get('values', [])
    email = ""

    if not values:
        print('No data found.')
    else:
        RecentChecked = values[1][16] #가장 최근에 코드 실행했던 날짜

        RC_for_form = values[2][16] #Typeform 가장 최근에 읽은 답변의 submitted 시각
        sub_date_dict = dict()

        responses = Typeform("#####").responses 
        result: dict = responses.list(uid = "#####", since = RC_for_form) #Typeform의 답변을 가장 최근에 읽은 답변 다음부터 가져오기

        result = result['items']
        tokens = dict()
        submitted_date = dict()

        #Typeform 답변 중, '네'라고 답한 것들만 모으기
        for i in result: 
            answer = i['answers']
            if(len(answer)==4 and i['submitted_at'] != RC_for_form):
                if(answer[0]['choice']['label'] == '예'):
                    date = answer[1]['date'][0:10]
                    if 'other' in answer[3]['choice']:
                        print("수동으로 Typeform 체크해주세요 : ", end = "")
                        print(answer[3]['choice']['other'])
                        continue
                    author = answer[3]['choice']['label']
                    date = date.replace('-','.')
                    tokens[author] = i['token']
                    submitted_date[author] = i['submitted_at']
                    sub_date_dict[author] = date

        for i in range(1,len(values),1):
            row = values[i]
            if(len(row)>=15):
                LastEdit = row[5]
                if (row[13] == "TRUE" and row[14]=="TRUE" and row[15]=="TRUE"):
                    if(row[10]=="FALSE" and row[4] == "" and (datetime.now()-datetime.strptime(row[5],'%Y.%m.%d')).days >= POST_ALARM):
                        email = MANAGER_EMAIL
                        author = row[1]
                        DOCUMENT_ID = row[12]
                        message = postMagazineAlarm(author,email,DOCUMENT_ID)
                        send_message(service2,GOOGLE_EMAIL, message)
                    elif row[10]=="FALSE" and row[4] == "":
                        print("skip")
                    elif(row[10] == "FALSE" and datetime.strptime(row[4],'%Y.%m.%d') <= datetime.now()):
                        email = row[6]
                        author = row[1]
                        postURL = row[9]
                        message = postAlarmMail(author, email, postURL)
                        send_message(service2, GOOGLE_EMAIL, message)

                        request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                                    range=SPREADSHEET_NAME+'!K'+str(i+1), # 2
                                                                    valueInputOption='RAW',
                                                                    body={
                                                                        'values' : [['TRUE']]
                                                                    })
                        request.execute()
                    else:
                        continue
                elif (row[12] == "" and row[13] == "FALSE" and row[3] != "FALSE" and (datetime.strptime(row[3],'%Y.%m.%d') - datetime.now()).days <ESTIMATED_DATE_LEFT ):
                    #문서 url이 입력되어있지 않고, 첫 메일도 안보내져있지만, 투고 예정일은 정해져있고, 해당 투고 예정일이 한달도 안남았을 때 
                    #-> 관리자에게 필진이 투고할 문서 만들고, 해당 url을 입력하라고 알리는 메일 보냄
                    email = MANAGER_EMAIL
                    print("email:", email)
                    message = makeDocu(row[1], email)
                    send_message(service2, GOOGLE_EMAIL, message)
                    continue
                elif (LastEdit == '' and row[13]=="FALSE" and row[12] != ""):
                    #문서 url은 입력되어있고, 첫 메일이 보내지지 않은 상황 -> 필진에게 투고할 문서 url을 전달하고 투고해달라고 요청하는 메일 보냄
                    document_Link = row[12]
                    email = row[6]
                    print("email:", email)
                    message = firstMail(row[1],document_Link, email)
                    send_message(service2, GOOGLE_EMAIL, message)

                    request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                                    range=SPREADSHEET_NAME+'!N'+str(i+1), # 2
                                                                    valueInputOption='RAW',
                                                                    body={
                                                                        'values' : [['TRUE']]
                                                                    })
                    request.execute()
                elif(LastEdit =='' and row[13] == "TRUE" and row[12] != ""):
                    #투고 알림 첫 메일을 보냈을 때
                    #투고 완료인지 문서의 comment를 통해 확인하기
                    DOCUMENT_ID = row[12]
                    index = DOCUMENT_ID.index('/d/')
                    index_end = DOCUMENT_ID.index('/edit?usp=sharing')
                    DOCUMENT_ID = DOCUMENT_ID[index+3:index_end]
                    result = retrieve_comments(service4, DOCUMENT_ID)
                    b = False #초고 제출 댓글이 있는지 없는지 확인하기 위한 변수
                    for j in range(len(result)): #문서의 comment를 for문을 통해 돌면서 초고 제출 댓글이 체크되었는지 확인
                        if (result[j]['content'] == '초고 제출' and 'resolved' in result[j] and result[j]['resolved'] == True):
                            #초고 제출 댓글이 있다면, b 변수를 True로 바꾸고, 관리자에게 초고 제출 알림 메일을 보냄
                            email = MANAGER_EMAIL
                            DOCUMENT_ID = row[12]
                            message = submittedFirst(row[1], DOCUMENT_ID, email)
                            send_message(service2, GOOGLE_EMAIL, message)
                            b = True 

                    if (b==False and datetime.strptime(row[3],'%Y.%m.%d')>datetime.now() and (datetime.strptime(row[3],'%Y.%m.%d') - datetime.now()).days <ESTIMATED_DATE_LEFT_NOT_SUMBITTED):
                        #매거진 제출 예정일이 일주일 남았는데, 아직 초고가 제출되지 않은 경우, 필진에게 초고 미제출 알림 메일을 보냄
                        document_Link = row[12]
                        email = row[6]
                        print("email:", email)
                        message = notSubmittedFirst(row[1], document_Link, email, (datetime.strptime(row[3],'%Y.%m.%d') - datetime.now()).days)
                        send_message(service2, GOOGLE_EMAIL, message)
                elif(LastEdit != '' and datetime.strptime(LastEdit,'%Y.%m.%d')>=datetime.strptime(RecentChecked,'%Y.%m.%d') and row[13]=="TRUE" and row[14]=="FALSE"):
                    #관리자가 문서 수정 후, 수정 알림을 필진에게 아직 안보냈을 때 -> 필진에게 수정 완료 메일 보냄
                    document_Link = row[12]
                    print(row) 
                    email = row[6]
                    print("email:", email)
                    stageURL = str(row[7])
                    message = finishedModify(row[1], stageURL, document_Link, email)
                    send_message(service2, GOOGLE_EMAIL, message)

                    request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                                        range=SPREADSHEET_NAME+'!O'+str(i+1), # 2
                                                                        valueInputOption='RAW',
                                                                        body={
                                                                            'values' : [['TRUE']]
                                                                        })
                    request.execute()
                elif(LastEdit != '' and datetime.strptime(LastEdit,'%Y.%m.%d')<datetime.strptime(RecentChecked,'%Y.%m.%d') and row[13]=="TRUE" and row[14]=="FALSE"):
                    #관리자가 아직 수정 안한경우
                    DOCUMENT_ID = row[12]
                    index = DOCUMENT_ID.index('/d/')
                    index_end = DOCUMENT_ID.index('/edit?usp=sharing')
                    DOCUMENT_ID = DOCUMENT_ID[index+3:index_end]
                    result = retrieve_comments(service4, DOCUMENT_ID)
                    b = False #문서의 comment중 '수정완료' or '수정할 것이 없다'가 있는지 나타내는 변수
                    content = [result[j]['content'] for j in range(len(result)) if 'resolved' in result[j] and result[j]['resolved'] == True]
                    if '수정할 내용이 없습니다' in content:
                        b = True
                        email = MANAGER_EMAIL
                        DOCUMENT_ID = row[12]
                        message = finalSubmitted(row[1], DOCUMENT_ID, email)
                        send_message(service2,GOOGLE_EMAIL, message)

                    elif '수정 완료' in content:
                        b = True
                        email = MANAGER_EMAIL
                        DOCUMENT_ID = row[12]
                        message = requestModify(row[1], DOCUMENT_ID, email)
                        send_message(service2, GOOGLE_EMAIL, message)

                    if (b == False):
                        #필진이 문서에 대해서 아무 체크도 하지 않은 경우
                        now = datetime.now()
                        last_edit_time = datetime.strptime(LastEdit,'%Y.%m.%d')
                        minus = now - last_edit_time
                        if(minus.days>=NOT_YET_CHECK_FILE):
                            #필진이 메일 보내고 7일동안 아무 체크도 하지 않은 경우 -> 관리자에게 확인해달라는 메일 보냄
                            email = MANAGER_EMAIL
                            DOCUMENT_ID = row[12]
                            message = notReadMail(row[1], DOCUMENT_ID, row[6], email)
                            send_message(service2, GOOGLE_EMAIL, message)
                elif(row[14]=="TRUE" and row[15]=="FALSE"):
                    #수정완료에 대한 메일을 보낸 후
                    #수정을 정말 다 했는지 확인
                    DOCUMENT_ID = row[12]
                    index = DOCUMENT_ID.index('/d/')
                    index_end = DOCUMENT_ID.index('/edit?usp=sharing')
                    DOCUMENT_ID = DOCUMENT_ID[index+3:index_end]
                    result = retrieve_comments(service4, DOCUMENT_ID)
                    b = False
                    content = [result[j]['content'] for j in range(len(result)) if 'resolved' in result[j] and result[j]['resolved'] == True]
                    if '수정할 내용이 없습니다' in content:
                        b = True
                        email = MANAGER_EMAIL
                        DOCUMENT_ID = row[12]
                        message = finalSubmitted(row[1], DOCUMENT_ID, email)
                        send_message(service2,GOOGLE_EMAIL, message)
                        request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                        range=SPREADSHEET_NAME+'!P'+str(i+1), # 2
                                                        valueInputOption='RAW',
                                                        body={
                                                            'values' : [['TRUE']]
                                                        })
                        request.execute()
                    elif '수정 완료' in content:
                        b = True
                        email = MANAGER_EMAIL
                        DOCUMENT_ID = row[12]
                        message = requestModify(row[1], DOCUMENT_ID, email)
                        send_message(service2, GOOGLE_EMAIL, message)
                        request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                range=SPREADSHEET_NAME+'!O'+str(i+1), # 2
                                valueInputOption='RAW',
                                body={
                                    'values' : [['FALSE']]
                                })
                        request.execute()

                    if (b == False):
                        #필진이 메일 보내고 7일동안 아무 체크도 하지 않은 경우 -> 관리자에게 확인해달라는 메일 보냄
                        now = datetime.now()
                        last_edit_time = datetime.strptime(LastEdit,'%Y.%m.%d')
                        minus = now - last_edit_time
                        if(minus.days>=NOT_YET_CHECK_FILE):
                            email = MANAGER_EMAIL
                            DOCUMENT_ID = row[12]
                            message = notReadMail(row[1], DOCUMENT_ID, row[6], email)
                            send_message(service2, GOOGLE_EMAIL, message)


        RecentChecked = values[1][16] #만약을 대비해서, Recent Checked를 다시 한 번 읽어서 넣어줌
        RecentChecked = datetime.strptime(RecentChecked, '%Y.%m.%d')
        last_date_dict = dict()


        for i in range(2, len(values), 1):
            row = values[i]
            if(len(row) >= 15):
                if (row[13] == "TRUE" and row[14]=="TRUE" and row[15]=="TRUE" and row[4] != ""):
                    #이미 모두 완료된 경우는 Post Date last_date_dict에 넣음
                    last_date_dict[row[1]] = [row[4], i]
                elif(row[13] == "FALSE" and row[14]=="FALSE" and row[15]=="FALSE"):
                    #아예 투고조차 시작되지 않은 경우는 Estimated Date를 last_date_dict에 넣음
                    last_date_dict[row[1]] = [row[3], i]
                else:
                    if row[1] in last_date_dict.keys():
                        #현재 매거진 관련 작업이 진행중인 경우는 Estimated Date를 받을 준비를 하지 않도록 빼버림
                        del last_date_dict[row[1]]

        # schedule_mail_date = str(datetime.now().year)+"."+str(datetime.now().month)+"."+str(datetime.now().day) #Estimate Date를 업데이트한 날짜 == 코드 돌린 날짜
        # append_list = [] #GSheet에 넣을 내용 넣는 리스트
        final_submit = RC_for_form #가장 마지막으로 읽은 form의 제출 시간

        for i in sub_date_dict.keys():
            if i in last_date_dict.keys():
                row = values[last_date_dict[i][1]]
                if(row[3] == "FALSE"): #다음 매거진 행을 생성했지만, 아직 제출된 Estimated Date가 입력되지 않음 -> Estimated Date를 입력해줌
                    request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                                            range=SPREADSHEET_NAME+'!D'+str(last_date_dict[i][1]+1), # 2
                                                                            valueInputOption='RAW',
                                                                            body={
                                                                                'values' : [[sub_date_dict[i]]]
                                                                            })
                    request.execute()
                    last_date_dict[i][0] = sub_date_dict[i]  #입력된 Estimated Date로 값을 업데이트 해줌
                    if(submitted_date[i] > final_submit):   #이번에 조회하는 TypeForm의 가장 최근 제출 날짜를 구하기 위함
                        final_submit = submitted_date[i]

        schedule_mail_date = str(datetime.now().year)+"."+str(datetime.now().month)+"."+str(datetime.now().day) #Estimate Date를 업데이트한 날짜 == 코드 돌린 날짜
        append_list = [] #GSheet에 넣을 내용 넣는 리스트

        for i in last_date_dict.keys():
            row = values[last_date_dict[i][1]]
            if (last_date_dict[i][0] != "FALSE" and last_date_dict[i][0] != "" and (datetime.now() - datetime.strptime(last_date_dict[i][0], '%Y.%m.%d')).days >= TOGOFORM_DAY):
                #이전 매거진이 모두 완료가 되었고, 지금 시간이 마지막 매거진의 post date에서 DAY만큼 지날 시, 새로운 매거진을 위한 행을 생성한다
                #그리고 필진에게 투고 더 할 의향이 있는지 물어보는 type form url을 첨부한 메일을 보냄
                row = values[last_date_dict[i][1]]
                Author = row[1]
                Estimated_submission_date = 'FALSE'
                Author_email = row[6]
                a_list = ["", Author, "", Estimated_submission_date,"", "", Author_email,"", schedule_mail_date,"","FALSE","","","FALSE","FALSE","FALSE"]
                append_list.append(a_list)
                message = togoPlanForm(Author, Author_email)
                send_message(service2, GOOGLE_EMAIL, message)

            elif (last_date_dict[i][0] == "FALSE" and (datetime.now() - datetime.strptime(row[8], '%Y.%m.%d')).days >= TOGOFORM_RE_DAY):
                #이미 투고 더 할 의향이 있는지 물어봤는데, 답이 없고, DAY만큼 시간이 지났을 때 -> 다시 물어보는 메일 보냄
                Author = row[1]
                Author_email = row[6]
                message = togoPlanForm(Author, Author_email)
                send_message(service2, GOOGLE_EMAIL, message)

                request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                                    range=SPREADSHEET_NAME+'!I'+str(last_date_dict[i][1]+1), # 2
                                                                    valueInputOption='RAW',
                                                                    body={
                                                                        'values' : [[schedule_mail_date]]
                                                                    })
                request.execute()
            else:
                continue
        #GSheet의 밑부분에 아까 새로운 매거진을 위해 추가하려고 준비한 행들을 업로드 해준다.
        resource = {
          "majorDimension" : "ROWS",
          "values" : append_list
        }
        RANGE = SPREADSHEET_NAME+"!A"+str(len(values)+1)
        service3.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=RANGE,body=resource,insertDataOption = 'INSERT_ROWS', valueInputOption = 'RAW' ).execute()
            
        #해당 코드를 실행시킨 시간을 GSheet에 업데이트하기
        now = datetime.now()
        month = str(now.month)
        day = str(now.day)
        if(len(month) == 1) :
            month = '0'+month
        if(len(day)==1):
            day = '0' + day
        now_date = str(now.year)+"."+month+"."+day
        now_date = str(now_date)
        request = service3.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,
                                                                    range=SPREADSHEET_NAME+'!Q2', # 2
                                                                    valueInputOption='RAW',
                                                                    body={
                                                                        'values' : [[now_date],[final_submit]]
                                                                    })
        request.execute()                    


if __name__ == '__main__':
    main()
