import gspread, os, random, base64, smtplib, time
from jinja2 import Environment, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


env = {}
with open('./info_bellringer.txt','r') as f:
    infofile = f.read()
    infofile = infofile.strip()
    infofile = infofile.split('\n')
    for i in range(0,len(infofile)):
        infofile[i] = infofile[i].split(' = ')
        env[infofile[i][0]] = infofile[i][1]

jsonLocalURL = env['JSON_LOCAL_URL']
sheetURL = env['SHEET_URL']
worksheetName = env['WORKSHEET_NAME']
my_account = env['ACCOUNT']
my_password = env['PASSWORD']
githubJavascriptURL = env['GITHUB_JAVASCRIPT_URL']

problemDirPath = './problems/'
outputDirPath = './output/'
userListPath = './setup/userlist.txt'

sheetColInfo = {'timeStamp':1, 'email':2, 'npp':3, 'page':4, 'period':5, 'coverage':6, 'weight5':7, 'weight4':8, 'weight3':9, 'weight2':10, 'weight1':11,'isDone':12}
fileListInfo = {'timeStamp':1, 'email':2, 'image_url':3, 'answer':4, 'coverage':5, 'weight':6, 'organ':7, 'path':8}

gc = gspread.service_account(filename=jsonLocalURL)
sheet = gc.open_by_url(sheetURL)
worksheet = sheet.worksheet(worksheetName)
probability_table = [1, 2,2, 3,3,3,3,3,3, 4,4,4,4,4,4,4,4, 5,5,5,5,5,5,5,5,5,5]


def initiate(requestRowDic):
    ##작업시작##
    with open(userListPath, 'r') as f:
        userList = f.read()
        userList = userList.strip()
    userList = userList.split('\n')
    if not (requestRowDic['data'][sheetColInfo['email']-1] in userList): #비허가 이메일 요청은 없던일로 처리함
        worksheet.update_cell(requestRowDic['row'], sheetColInfo['isDone'], 'denied')
        return
    
    #선택한 범위 중에서 저장된 문제 있는 것들을 가져옴
    coverage_list = requestRowDic['data'][sheetColInfo['coverage']-1]
    coverage_list = coverage_list.split(', ')
    uploaded_coverage = os.listdir(problemDirPath)
    temp = []
    for i in coverage_list:
        if i in uploaded_coverage:
            temp.append(i)
    coverage_list = temp
    del temp
    
    num_of_problems = int(requestRowDic['data'][sheetColInfo['npp']-1]) * int(requestRowDic['data'][sheetColInfo['page']-1])
    generateProbabilityTable(requestRowDic)
    problemList = selectProblem(coverage_list, num_of_problems)
    html_path = makeHTML(requestRowDic, problemList)
    address = requestRowDic['data'][sheetColInfo['email']-1]
    sendingResult = sendMail(requestRowDic, address, html_path)
    worksheet.update_cell(requestRowDic['row'], sheetColInfo['isDone'], sendingResult) 
    return


def selectProblem(coverageList, numOfProblems):
    problemList = []
    uploadedProblemDict = {} #문제 요약본([범위][난이도] 안에는 정답), 저장된 문제들, 사진 여러장인 것들 중복없이 저장함
    uploadedFileDict = {} #파일 경로 어디있나, 모든 업로드정답파일들
    for i in coverageList:
        uploadedProblemDict[i] = {1:[], 2:[], 3:[], 4:[], 5:[]}
        uploadedFileDict[i] = {1:[], 2:[], 3:[], 4:[], 5:[]}
        with open(problemDirPath+i+'/fileList.txt', 'r') as f:
            fl = f.read()
            fl = fl.strip()
        fl = fl.split('\n')
        for j in fl:
            temp = j.split('\t')
            weight = int(temp[fileListInfo['weight']-1])
            temp = {'answer':temp[fileListInfo['answer']-1], 'organ':temp[fileListInfo['organ']-1]}
            l = uploadedProblemDict[i][weight]
            if not(temp in l):
                uploadedProblemDict[i][weight].append(temp)
        for j in fl:
            temp = j.split('\t')
            weight = int(temp[fileListInfo['weight']-1])
            temp = {'answer':temp[fileListInfo['answer']-1], 'organ':temp[fileListInfo['organ']-1], 'path':temp[fileListInfo['path']-1]}
            uploadedFileDict[i][weight].append(temp)
    coverage_list = coverageList[:]
    random.shuffle(coverage_list)
    selectedProblem = {} #여기에 들어가는건 선택되면 안됨, [범위] 안에 선발된 문제들 list로 넣음(다 썼으면 초기화)
    for i in coverageList:
        selectedProblem[i] = []
    
    #html에 넣을 문제 구성
    for i in range(numOfProblems):
        element = {} #이게 한문제 - 정답, 이미지파일, 가중치
        #coverage_list를 랜덤으로 범위를 나열해서 순서대로 사용
        if len(coverage_list)==0:
            coverage_list = coverageList[:]
            random.shuffle(coverage_list)
        coverage = coverage_list.pop() #범위 선택됨
        #중복문제방지
        tot_Prblms = 0
        for i in uploadedProblemDict[coverage].keys():
            tot_Prblms += len(uploadedProblemDict[coverage][i])
        if len(selectedProblem[coverage]) >= tot_Prblms:
            selectedProblem[coverage] = []
        #문제 선택
        while True: 
            selectedWeight = random.choice(probability_table) # weight 선택됨
            if len(uploadedProblemDict[coverage][selectedWeight]) > 0:
                selectedAnswer = random.choice(uploadedProblemDict[coverage][selectedWeight])
                element = {'weight':selectedWeight, 'answer':selectedAnswer['answer'], 'organ':selectedAnswer['organ']}
                if not(element in selectedProblem[coverage]):
                    selectedProblem[coverage].append(element)
                    # random하게 선택된 weight, answer, organ을 기준으로 이미지파일을 random 선택
                    imageFiles = []
                    for i in uploadedFileDict[coverage][selectedWeight]:
                        if i['answer']==element['answer'] and i['organ']==element['organ']:
                            imageFiles.append(i)
                    imageFile = random.choice(imageFiles)
                    with open(imageFile['path'], "rb") as f:
                        base64_data = base64.b64encode(f.read()).decode("utf-8")
                    element_pl = element.copy()
                    element_pl['problem_image']= f"data:image/{imageFile['path'].split('.')[-1]};base64,{base64_data}"
                    problemList.append(element_pl)
                    break
    
    return problemList 


def generateProbabilityTable(requestRowDic):
    global probability_table
    probability_table = []
    for i in [1,2,3,4,5]: 
        col = 'weight' + str(i)
        weight = int(requestRowDic['data'][sheetColInfo[col]-1])
        for j in range(weight):
            probability_table.append(i)
    return probability_table


def makeHTML(requestRowDic, problemList):
    env = Environment(loader=FileSystemLoader('templates'))
    template = env.get_template('jinja_template.html')
    
    period = str(int(requestRowDic['data'][sheetColInfo['period']-1])*1000)
    
    keyValue = requestRowDic['data'][sheetColInfo['email']-1] + "\n"
    keyValue += requestRowDic['data'][sheetColInfo['timeStamp']-1] + "\n"
    keyValue += requestRowDic['data'][sheetColInfo['npp']-1] + "\n"
    keyValue += requestRowDic['data'][sheetColInfo['page']-1] + "\n"
    keyValue += period + "\n"
    keyValue += requestRowDic['data'][sheetColInfo['coverage']-1]
    encoded_bytes = base64.b64encode(keyValue.encode('utf-8'))
    encoded_string = encoded_bytes.decode('utf-8')
    keyValue = encoded_string
    
    html_filename = '땡시 연습문제_' + str(requestRowDic['row']) + '_' + requestRowDic['data'][sheetColInfo['timeStamp']-1] + '.html'
    rendered_html = template.render(problem_list=problemList, key_value=keyValue, javascriptURL=githubJavascriptURL)
    with open(outputDirPath+html_filename, 'w') as f:
        f.write(rendered_html)
    
    return outputDirPath + html_filename


def sendMail(requestRowDic, address, filepath):
    msg = MIMEMultipart()
    msg["Subject"] = "땡시 연습문제"
    msg["From"] = my_account
    msg["To"] = address

    requestTime = requestRowDic['data'][sheetColInfo['timeStamp']-1]
    requestNPP = requestRowDic['data'][sheetColInfo['npp']-1]
    requestTotalPages = requestRowDic['data'][sheetColInfo['page']-1]
    requestPeriod = requestRowDic['data'][sheetColInfo['period']-1]
    requestCoverage = requestRowDic['data'][sheetColInfo['coverage']-1]
    requestData = f"\n요청시간: {requestTime}\n페이지 당 문제 수: {requestNPP}\n총 페이지 수: {requestTotalPages}\n페이지 당 시간: {requestPeriod}\n범위: {requestCoverage}\n"
    content = f"안녕하세요\n\n요청하신 땡시 연습문제를 보내드립니다.\n\n연습문제를 첨부하였으니 다운받아서 사용하시면 됩니다.\n\n요청하신 정보는 다음과 같습니다.\n\n{requestData}\n\n감사합니다."
    msg.attach(MIMEText(content, "plain"))
    
    with open(filepath, 'rb') as f:
        filename = filepath.split('/')[-1]
        attachment = MIMEApplication(f.read())
        attachment.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(attachment)
    
    attempt = 0
    while attempt<5:
        try:
            smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            smtp.login(my_account, my_password)
            smtp.sendmail(my_account, address, msg.as_string())
        except smtplib.SMTPResponseException as e:
            error_code = str(e.smtp_code)
            if error_code[0]=='4': #일시적 오류
                attempt += 1
                time.sleep(60)
            elif error_code[0]=='5':#
                return error_code
        else:#메일 발송 성공
            attempt = 10
        finally:
            smtp.quit()
        
    if attempt == 10: #메일 발송 성공
        return 'O'
    else: #메일 발송 여러번 시도했으니 실패, 400 오류코드
        return ''


if __name__ == '__main__':
    while True:
        try:
            requestList = worksheet.get_all_values()[1:]
        except:
            sleep(60)
        else:
            unfinishedRow = []  #새로 들어 온 것들, 처리안된 구글폼 요청들
            for i in range(len(requestList)):
                if requestList[i][sheetColInfo['isDone']-1]=='':
                    unfinishedRow.append({'row':i+2, 'data':requestList[i]})
                    
            for i in unfinishedRow:
                initiate(i)
            time.sleep(300)
