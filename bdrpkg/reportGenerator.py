##########################################################################################
 
# Certain source files written and distributed by Leith Jun are subject to the AGPLv3.
# Otherwise you should ask me for the permission to use this package.

# Author : Leith Jun
# Contributor : Jihye Kwon
# Email : leithjun@osbc.co.kr

#    Copyright OSBC Inc. 2019  Leith Jun

#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.

#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.

#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.
 

##########################################################################################



import requests
import logging,sys
import csv
import os
import json
import xlsxwriter
import datetime
import copy
import urllib.parse





class blackduckRPT:

	def __init__(self, bdHttp, bdAccount, bdPassword):
		self.reportName = ""
		self.http = bdHttp
		self.account = bdAccount
		self.password = bdPassword
		self.idColor = ['#dfdfdf','#c05252','#f5dda8','#c2dfa5','#ffffff']
		self.fontName = "맑은 고딕"
		self.setSpace = 10
		self.datasetComponent = []


		self.mySession = requests.Session()
		self.mySession.verify = False
		requests.packages.urllib3.disable_warnings()
		self.CSRF = ''


		logging.debug("http is " + self.http)
		logging.debug("account is " + self.account)
		logging.debug("password is " + self.password)

		self.authenticate(self.http, self.account, self.password)
		

	#인증
	def authenticate(self, bdHttp, bdAccount, bdPassword):

		url = bdHttp + "/j_spring_security_check"
		logging.debug("authentication url is " + url)
		#세션 가져오기
		s = self.mySession
		#재인증은 반드시 무시
		s.verify=False
		#워닝 메시지 없애기
		requests.packages.urllib3.disable_warnings()
		#CSRF 초기화 
		s.CSRF = ''
		response = s.post(url, data = {'j_username':bdAccount, 'j_password':bdPassword})
		#인증의 성공은 204
		logging.debug("response code is " + str(response))
		#check for Success
		if response.ok:
			#인증 성공했을 떄 x-csrf-token을 가져오기 - 인증 할 때 마다 달라진다.
			logging.debug("You are ok with authentication.")
			s.CSRF = response.headers['x-csrf-token']			
		else:
			logging.error("Error in authentication. Please check out your ID and password")
			return sys.exit("try again")


	#프로젝트명이랑 버전 가져오기
	def getProjectsAndVersions(self):
		#현재 유저의 아이디 가져오기 
		currentUser_url = self.http + "/api/current-user"
		logging.debug("current URL is " + currentUser_url)

		s = self.mySession
		abc = s.get(currentUser_url)
		bcd = abc.json()
		cde = bcd['_meta']['href'].split('/')
		lastNum = len(cde)-1
		userID = cde[lastNum]

		logging.debug("user ID is " + userID)
		project_url = self.http + "/api/users/" + userID + "/projects"
		


		#파라미터 초기화

		limit = 0
		offset = 0
		sort = ''
		q = ''

		#프로젝트 총 개수 구하기
		repTemp = s.get(project_url)
		totalNum = repTemp.json()['totalCount']
		limit = totalNum
		#파라미터에 프로젝트 총 개수 대입
		payload = {'limit':limit, 'offset':offset, 'sort':sort, 'q':q}
		rep = s.get(project_url, params= payload)

		
		if rep.ok:
			projectForUser = rep.json()
			totalProjectDic = {}
			totalProjectInfo = []

			for i in range(len(projectForUser['items'])):
				
				#프로젝트 아이디 추출
				projectInfo = projectForUser['items'][i]['project']
				projectRowID = projectInfo.split('/')
				projectID = projectRowID[len(projectRowID)-1]
				logging.debug("project ID is " + projectID)
				#버전 아이디 추출
				url = self.http + "/api/projects/" + projectID + "/versions"
				s = self.mySession
				rep = s.get(url).json()
				count = len(rep['items'])
				#프로젝트 아이디 버전아이디 JSON을 변환 
				for j in range(0,count):
					temp_versionID = rep['items'][j]['_meta']['href'].split('/')
					logging.debug("ProjectName : " + projectForUser['items'][i]['name'] + " ProjecctID : " + projectID + " VersionName : " + rep['items'][j]['versionName'] + " VersionID : " + temp_versionID[len(temp_versionID) - 1])
					temp = {'projectName' : projectForUser['items'][i]['name'] , 'projectID' : projectID , 'versionName' : rep['items'][j]['versionName'] , 'versionID' : temp_versionID[len(temp_versionID) - 1]}
					totalProjectInfo.append(temp)

			totalProjectDic["KEPCO"]= totalProjectInfo	
			totalProjectJson = json.dumps(totalProjectDic)

			return totalProjectJson


					
		else:
			logging.error('Bad response in getProjects')
			return response.json()


		
	def getCSV(self):
		dataset2 = []
		temp = {'projectID':None, 'versionID':None}

		path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
		getPath = os.path.join(path,"resource", "analysis_project.csv")
		if os.path.basename(getPath) =="analysis_project.csv":
			logging.debug("successfully have an access to csv")
			file = open(getPath, newline="")
			reader = csv.reader(file)
			header = next(reader)
			tempData = [row for row in reader]
			
			
			for i in range(len(tempData)):

				temp['projectID'] = tempData[i][0]
				temp['versionID'] = tempData[i][1]

				dataset.append(copy.copy(temp))
				
			return dataset

		else:
			sys.exit("check out the name of csv file")


	def findIdentity(self, bdProject, bdVersion):
		projectName = bdProject
		versionName = bdVersion

		tempProject ={}
		myProject = []
		totalProjectInfo = self.getProjectsAndVersions()
		
		totalProject = json.loads(totalProjectInfo)
		
		#myProject = [i for i in totalProject['KEPCO'] if i['projectName'] == projectName and i['versionName'] == versionName]
		for i in totalProject['KEPCO']:

			if i['projectName'] == projectName and i['versionName'] == versionName:
				myProject = i
				break	
			else:
				logging.debug("I am still looking for it")
		
		if not myProject:
			logging.debug("Sorry I couldn't find it")
			
		
		else:
			logging.debug("I got it mate:")

			return myProject

	def createExcel(self,myProject):
		mp = myProject

		dataset = []
		dataset2 = []
		#리포트 산출 디렉토리 생성
		currentPath = os.path.dirname(__file__)
		logging.debug(os.path.dirname(__file__))
		try:
			logging.debug("projectID : " + mp['projectID'] + " versionID : " + mp['versionID'])
		except Exception as ex:
			logging.debug("Please check out project and version is uploaded in Black Duck.")
		newPath = os.path.abspath(os.path.join(currentPath, "..")) + "\\report"
		os.makedirs(newPath, exist_ok=True)

		fileName = "검증보고서-" + mp['projectName'] + "_" + mp['versionName'] + ".xlsx"

		reportPath = newPath + "\\" + fileName
		self.reportName = reportPath
		

		#엑셀 생성

		#################### 각 워크 시트 생성 및 기본 설정 #################### 

		worksheetList = ["프로젝트 개요","식별 컴포넌트 현황", "라이선스 종합의견", "보안취약점 대체솔루션", "리스크 분류기준", "용어정의", "(부록)보안취약점 현황", "(부록)검증파일 경로"]

		wb = xlsxwriter.Workbook(reportPath)
		wk = []
		riskCal = {}
		for i in range(len(worksheetList)):
			wk.append(wb.add_worksheet(worksheetList[i]))

		

		goToFuncList = [self.wkOverview(wb, wk[0], mp), self.wkComponent(wb, wk[1], mp), self.wkLicenseOpinion(wb, wk[2], mp), \
						self.wkSecuritySolution(wb, wk[3], mp), self.wkRiskCategory(wb, wk[4], mp), self.wkTermDefinition(wb, wk[5], mp), self.wkVulnerabilities(wb, wk[6], mp), self.wkSourcePath(wb, wk[7], mp)]

		for i in range(len(worksheetList)):
			#식별 컴포넌트 현황일 때 필요 내용 반환하여 다른 시트에서 재사용 
			if i == 1:

				riskCal = goToFuncList[i]
				self.writeOverviewRisk(riskCal, wb, wk[0])


			else:
				goToFuncList[i]

		dataset = self.wkSecuritySolutionData( self.reportName, wb, wk[1], mp)
		self.wkSecuritySolution2(dataset, wb, wk[3], mp)
		dataset2 = self.wkSecuritySolutionData3(dataset, wb, wk[3], mp)
		self.wkSecuritySolution3(dataset2, wb, wk[3], mp)
		

		wb.close()
		logging.info("#################### Lovely!! the report was created:) ####################")

	def wkOverview(self,  wkBook, wkSheet, myProject):

		logging.info("#################### wkOverview : Start ####################")

		wb = wkBook
		wk = wkSheet
		mp = myProject

		logging.info("current sheet is "  + wk.get_name())


		# 보고서 라인 그리기
		bg_line = wb.add_format()
		bg_line.set_bg_color("black")
		
		wk.merge_range('B2:I2', None ,bg_line)
		wk.merge_range('B4:I4', None ,bg_line)
		wk.merge_range('B11:I11', None ,bg_line)
		wk.merge_range('B13:I13', None ,bg_line)



		#타이틀 작성
		mainTitle = "한국전력공사" 
		subTitle = "공개SW 라이선스 검증보고서"
		titleFormat1 = {'bold':True, 'font_size':30, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter'}
		titleFormat2 = {'bold':True, 'font_size':24, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter'}

		cell_MainTitle = wb.add_format(titleFormat1)
		cell_SubTitle = wb.add_format(titleFormat2)
		wk.write('B6', mainTitle, cell_MainTitle)
		wk.write('B8', subTitle, cell_SubTitle)
		
		titleFormat3 = {'border':1, 'border_color': self.idColor[0] }
		cell_border = wb.add_format(titleFormat3)
		wk.write('B6:I9',"", cell_border)

		wk.merge_range('B6:I7', mainTitle ,cell_MainTitle)
		wk.merge_range('B8:I9', subTitle ,cell_SubTitle)

		row_line = self.setSpace
		wk.set_row(1,row_line)
		wk.set_row(2,row_line)
		wk.set_row(3,row_line)
		wk.set_row(10,row_line)
		wk.set_row(11,row_line)
		wk.set_row(12,row_line)

		wk.set_row(5,40)

		wk.set_column(0,0,1)
		wk.set_column(9,9,1)
		wk.set_column(1,8,10)

		#보고서 개요

		item1 = "■ 프로젝트 검증개요"
		temp1_item1 = {'bold':True, 'font_size':16, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter'}
		cell1_item1 = wb.add_format(temp1_item1)
		wk.write('B19', item1, cell1_item1)

		item1_content1 = "프로젝트 개요"
		temp1_content1 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color': self.idColor[0] }
		cell1_content1 = wb.add_format(temp1_content1)
		wk.merge_range('B21:I21', item1_content1, cell1_content1)
		wk.set_row(19,row_line)

		item1_content2 = ["서비스명", "서비스 유형", "프로젝트 배포유형", "보고서 생성일자"]
		temp1_item2 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter', 'border':1, 'border_color':'#000000'}
		cell1_item2 = wb.add_format(temp1_item2)
		for i in range(22,26):
			wk.merge_range('B'+ str(i) + ":D" + str(i), item1_content2[i-22], cell1_item2)
			wk.merge_range('E'+ str(i) + ":I" + str(i), None, cell1_item2)
		
		#보고서 생성일
		todayIs = datetime.date.today()
		todayIs = todayIs.isoformat()
		wk.write("E25", todayIs, cell1_item2)

		
		# 보안, 라이선스 위험도 표시
		item2 = ["■ 프로젝트 보안 위험 컴포넌트 현황", "■ 프로젝트 라이선스 위험 컴포넌트 현황"]
		temp_item2 = {'bold':True, 'font_size':16, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter'}
		cell_item2 = wb.add_format(temp_item2)
		wk.write('B31', item2[0], cell_item2)
		wk.set_row(31,row_line)
		wk.write('B38', item2[1], cell_item2)
		wk.set_row(38,row_line)


		item2_content1 = ["High", "Medium", "Low", "None"]
		temp2_content1 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[0] }
		cell2_content1 = wb.add_format(temp2_content1)

		temp2_content2 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[1] }
		temp2_content3 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[2]}
		temp2_content4 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[3] }
		temp2_content5 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }

		cell2_content2 = wb.add_format(temp2_content2)
		cell2_content3 = wb.add_format(temp2_content3)
		cell2_content4 = wb.add_format(temp2_content4)
		cell2_content5 = wb.add_format(temp2_content5)



		for i in range(33, 41, 7):

			wk.merge_range('B'+ str(i) + ":C" + str(i), item2_content1[0], cell2_content1)
			wk.merge_range('D'+ str(i) + ":E" + str(i), item2_content1[1], cell2_content1)
			wk.merge_range('F'+ str(i) + ":G" + str(i), item2_content1[2], cell2_content1)
			wk.merge_range('H'+ str(i) + ":I" + str(i), item2_content1[3], cell2_content1)

			wk.merge_range('B'+ str(i+1) + ":C" + str(i+1), None, cell2_content2)
			wk.merge_range('D'+ str(i+1) + ":E" + str(i+1), None, cell2_content3)
			wk.merge_range('F'+ str(i+1) + ":G" + str(i+1), None, cell2_content4)
			wk.merge_range('H'+ str(i+1) + ":I" + str(i+1), None, cell2_content5)
			
		dataset = self.overviewData(mp)

		distributionType = dataset['distribution']
		logging.info("project distribution type : " + distributionType)
		distributionCategory = {'INTERNAL':'내부사용(Internal Use)', 'SAAS':'네트워크 서비스(Network Service)', 'EXTERNAL':'외부배포(External Distribution)'}
		
		wk.write("E22", mp['projectName'] + " / " + mp['versionName'], cell1_item2)
		wk.write("E24", distributionCategory[distributionType], cell1_item2)


		logging.info("#################### wkOverview : Successfully created ####################")




	def overviewData(self, myProject):
		mp = myProject

		dataset = {}
		if not myProject:
			logging.debug("Sorry I couldn't find it")		
		else:
			url = self.http + "/api/projects/" + mp['projectID'] + "/versions/" + mp['versionID']
			logging.info("Project Name : " + mp['projectID'] + " / Project Version :" + mp['versionID'])
			s = self.mySession
			rep = s.get(url).json()
	
			dataset = {'distribution': rep['distribution']}

			return dataset

	def writeOverviewRisk(self, riskCal, wkBook, wkSheet):

		rc = riskCal
		wb = wkBook
		wk = wkSheet

		temp_content1 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[1] }
		temp_content2 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[2]}
		temp_content3 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[3] }
		temp_content4 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }

		cell_content1 = wb.add_format(temp_content1)
		cell_content2 = wb.add_format(temp_content2)
		cell_content3 = wb.add_format(temp_content3)
		cell_content4 = wb.add_format(temp_content4)

		wk.write('B34', rc["securityHigh"], cell_content1)
		wk.write('D34', rc["securityMedium"], cell_content2)
		wk.write('F34', rc["securityLow"], cell_content3)
		wk.write('H34', rc["securityNone"], cell_content4)			
		
		wk.write('B41', rc["licenseHigh"], cell_content1)
		wk.write('D41', rc["licenseMedium"], cell_content2)
		wk.write('F41', rc["licenseLow"], cell_content3)
		wk.write('H41', rc["licenseNone"], cell_content4)	
		


	
	def wkComponent(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkComponent : Start ####################")

		wb = wkBook
		wk = wkSheet
		mp = myProject

		riskCal = {}

		calSecurityHigh = 0
		calSecurityMedium = 0
		calSecurityLow = 0
		calSecurityNone = 0

		calLicenseHigh = 0
		calLicenseMedium = 0
		calLicenseLow = 0
		calLicenseNone = 0

		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 

		columnItems = ["버전", "라이선스 유형", "라이선스", "결합형태", "보안 위험도", "라이선스 위험도"]
		riskCategory = ["High", "Medium", "Low", "None"]

		wkTitle = {'bold':True, 'font_size':10, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[0] }
		wkCellTitle = wb.add_format(wkTitle)
		wkContent = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent = wb.add_format(wkContent)

		wkRiskHigh = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[1] }
		wkCellRiskHigh = wb.add_format(wkRiskHigh)
		wkRiskMedium = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[2] }
		wkCellRiskMedium = wb.add_format(wkRiskMedium)
		wkRiskLow = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[3] }
		wkCellRiskLow = wb.add_format(wkRiskLow)
		wkRiskNone = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellRiskNone = wb.add_format(wkRiskNone)

		
		#셀 크기 설정
		wk.set_default_row(25)
		wk.set_column(0,0,1)
		wk.set_column(9,9,1)
		wk.set_column(1,1,12)
		wk.set_column(3,3,5)
		wk.set_column(4,6,12)
		wk.set_column(7,7,10)
		wk.set_column(7,8,12)

		

		# 컴포넌트만 병합	
		wk.merge_range('B'+ str(1) + ":C" + str(1), "컴포넌트", wkCellTitle)	

		# 나머지 제목

		for i in range(len(columnItems)):

			wk.write(0, i+3, columnItems[i], wkCellTitle)

		#################### 각 워크 시트 데이터 #################### 

		dataset = self.wkComponentData(mp)



		

		#for i in range(len(dataset)):
		for i in range(len(dataset)):
			
			
			wk.merge_range('B'+ str(i+2) + ":C" + str(i+2), dataset[i]['componentName'], wkCellContent)
			wk.write(i+1,3, dataset[i]['componentVersionName'], wkCellContent)

			#라이선스 유형 
			wk.write(i+1,4, dataset[i]['licenseFamily'], wkCellContent)
			wk.write(i+1,5, dataset[i]['licenseDisplay'], wkCellContent)
			wk.write(i+1,6, dataset[i]['usages'], wkCellContent)

			#보안위험도

			if dataset[i]['securityRisk'] == "UNKNOWN" or dataset[i]['securityRisk'] == "HIGH" or dataset[i]['securityRisk'] == "CRITICAL":
				wk.write(i+1,7, riskCategory[0], wkCellRiskHigh)
				calSecurityHigh += 1
			elif dataset[i]['securityRisk'] == "MEDIUM":
				wk.write(i+1,7, riskCategory[1], wkCellRiskMedium)
				calSecurityMedium += 1
			elif dataset[i]['securityRisk'] == "LOW":
				wk.write(i+1,7, riskCategory[2], wkCellRiskLow)
				calSecurityLow += 1
			elif dataset[i]['securityRisk'] == "NONE":
				wk.write(i+1,7, riskCategory[3], wkCellRiskNone)
				calSecurityNone += 1


			#라이선스 위험도

			if dataset[i]['licenseRisk'] == "UNKNOWN" or dataset[i]['licenseRisk'] == "HIGH" or dataset[i]['licenseRisk'] == "CRITICAL":
				wk.write(i+1,8, riskCategory[0], wkCellRiskHigh)
				calLicenseHigh += 1
			elif dataset[i]['licenseRisk'] == "MEDIUM":
				wk.write(i+1,8, riskCategory[1], wkCellRiskMedium)
				calLicenseMedium += 1
			elif dataset[i]['licenseRisk'] == "LOW":
				wk.write(i+1,8, riskCategory[2], wkCellRiskLow)
				calLicenseLow += 1
			elif dataset[i]['licenseRisk'] == "NONE":
				wk.write(i+1,8, riskCategory[3], wkCellRiskNone)
				calLicenseNone += 1

		logging.info("#################### wkComponent : Successfully created ####################")
		riskCal = {'securityHigh' : calSecurityHigh,'securityMedium' : calSecurityMedium,'securityLow' : calSecurityLow,'securityNone' : calSecurityNone,'licenseHigh' : calLicenseHigh,'licenseMedium' : calLicenseMedium,'licenseLow' : calLicenseLow,'licenseNone' : calLicenseNone}
		
		return riskCal

	def wkComponentData(self, myProject):
		mp = myProject


		dataset =[]


		if not myProject:
			logging.debug("Sorry I couldn't find it")		
		else:
			url = self.http + "/api/projects/" + mp['projectID'] + "/versions/" + mp['versionID'] + "/components"
			

			########################### 리포트 산출값 조정 limit, ###########################
			limit = 0
			offset =0
			sort = ''
			q = ''

			s = self.mySession

			#프로젝트 총 개수 구하기
			repTemp = s.get(url)
			totalNum = repTemp.json()['totalCount']
			
			
			limit = totalNum
			temp = {}
			licenseRisk = []
			securityRisk = []
			licenseFamily = []

			#파라미터에 프로젝트 총 개수 대입
			payload = {'limit':limit, 'offset':offset, 'sort':sort, 'q':q}
			rep = s.get(url, params= payload).json()

			tempSecurityRisk = ""
			securityRiskFlag = 0
			licenseRiskFlag = 0
			securityRiskScore = [ 0, 0, 10, 20, 50, 100]
			licenseRiskScore = [ 0, 0, 10, 20, 50, 100]


			

			# 라이선스 시큐리티 리스크 구하기
			for i in range(limit):
				for j in range(len(rep['items'][i]['licenseRiskProfile']['counts'])):


					#보안 리스크 스코어링
					if rep['items'][i]['securityRiskProfile']['counts'][j]['count'] > 0 :

						if rep['items'][i]['securityRiskProfile']['counts'][j]['countType'] == "UNKNOWN":
							securityRiskFlag += securityRiskScore[0]

						elif rep['items'][i]['securityRiskProfile']['counts'][j]['countType'] == "OK":
							securityRiskFlag += securityRiskScore[1]

						elif rep['items'][i]['securityRiskProfile']['counts'][j]['countType'] == "LOW":
							securityRiskFlag += securityRiskScore[2]

						elif rep['items'][i]['securityRiskProfile']['counts'][j]['countType'] == "MEDIUM":
							securityRiskFlag += securityRiskScore[3]
					
						elif rep['items'][i]['securityRiskProfile']['counts'][j]['countType'] == "HIGH":
							securityRiskFlag += securityRiskScore[4]

						elif rep['items'][i]['securityRiskProfile']['counts'][j]['countType'] == "CRITICAL":
							securityRiskFlag += securityRiskScore[5]
					
					else:
						logging.debug("Error came up!! check it out")
						securityRiskFlag += 0

					# 리스크 레벨 지정 로직 
					if securityRiskFlag == 0 :
						tempSecurityRisk = "NONE"
						
						
					elif securityRiskFlag >= 10 and securityRiskFlag < 20 :
						tempSecurityRisk = "LOW"
						

					elif securityRiskFlag >= 20 and securityRiskFlag <= 30 :
						tempSecurityRisk = "MEDIUM"
						
						
					elif securityRiskFlag >= 50 :
						tempSecurityRisk = "HIGH"
						


					#라이선스 리스크 스코어링
					if rep['items'][i]['licenseRiskProfile']['counts'][j]['count'] > 0 :

						if rep['items'][i]['licenseRiskProfile']['counts'][j]['countType'] == "UNKNOWN":
							licenseRiskFlag += licenseRiskScore[0]
	
						elif rep['items'][i]['licenseRiskProfile']['counts'][j]['countType'] == "OK":
							licenseRiskFlag += licenseRiskScore[1]


						elif rep['items'][i]['licenseRiskProfile']['counts'][j]['countType'] == "LOW":
							licenseRiskFlag += licenseRiskScore[2]


						elif rep['items'][i]['licenseRiskProfile']['counts'][j]['countType'] == "MEDIUM":
							
							licenseRiskFlag += licenseRiskScore[3]

					
						elif rep['items'][i]['licenseRiskProfile']['counts'][j]['countType'] == "HIGH":
							licenseRiskFlag += licenseRiskScore[4]


						elif rep['items'][i]['licenseRiskProfile']['counts'][j]['countType'] == "CRITICAL":
							licenseRiskFlag += licenseRiskScore[5]

					else:
						logging.debug("Error came up!! check it out")


					# 리스크 레벨 지정 로직 - MIT 예외처리 함
					if licenseRiskFlag == 0 :
						tempLicenseRisk = "NONE"
						
					########################### 특정 라이선스 제외 ###########################	
					elif licenseRiskFlag >= 10 and licenseRiskFlag < 20 :
						
						if rep['items'][i]['licenses'][0]['licenseDisplay'] == "MIT License":
							tempLicenseRisk = "NONE"
							
						else:
							tempLicenseRisk = "LOW"
							
					########################### 특정 라이선스 제외 ###########################
					elif licenseRiskFlag >= 20 and licenseRiskFlag <= 30 :
						
						if rep['items'][i]['licenses'][0]['licenseDisplay'] == "MIT License":
							tempLicenseRisk = "NONE"
							
						else:
							tempLicenseRisk = "MEDIUM"
							
						
					elif licenseRiskFlag >= 50 :
						tempLicenseRisk = "HIGH"
						
					

						

				securityRisk.append(tempSecurityRisk)
				logging.debug("Security Risk Score = " + str(securityRiskFlag) + " and Risk = " + tempSecurityRisk)
				#logging.info("Security Risk Score = " + str(securityRiskFlag) + " and Risk = " + tempSecurityRisk)

				
				licenseRisk.append(tempLicenseRisk)
				logging.debug("License Risk Score = " + str(licenseRiskFlag) + " and Risk = " + tempLicenseRisk)
				#logging.info("License Risk Score = " + str(licenseRiskFlag) + " and Risk = " + tempLicenseRisk)

				licenseRiskFlag = 0
				securityRiskFlag = 0

			logging.info("Extract Security Risk : " + json.dumps(securityRisk))
			logging.info("Extract License Risk : " + json.dumps(licenseRisk))


			temp = { 'componentName': None, \
							'componentURL': None, \
							'componentVersionName': None, \
							'componentVersionURL': None, \
							'licenseDisplay': None, \
							'licenseURL': None, \
							'usages': None, \
							'vulnerabilityURL': None, \
							'licenseRisk': None, \
							'securityRisk': None \
							}

			for i in range(limit):
						
				logging.info("############# Component No." + str(i) + " Information #############")
				temp['componentName']=rep['items'][i]['componentName']
				temp['componentURL'] =rep['items'][i]['component']
				
				try:		
					temp['componentVersionName']=rep['items'][i]['componentVersionName']
				except KeyError as ex:
					temp['componentVersionName']:None

				try:		
					temp['componentVersionURL']=rep['items'][i]['componentVersion']
				except KeyError as ex:
					temp['componentVersionURL']:None
			

				try:
					temp['licenseDisplay'] = rep['items'][i]['licenses'][0]['licenseDisplay']
					temp['licenseURL']= rep['items'][i]['licenses'][0]['license']
				except KeyError as ex:
					temp['licenseDisplay'] = rep['items'][i]['licenses'][0]['licenseDisplay'] + " : 듀얼 라이선스 중 선택"
					temp['licenseURL'] = rep['items'][i]['licenses'][0]['licenses'][0]['license']
				licenseFamily = self.wkComponentGetLicenseFamilty(temp['licenseURL'], mp)
				temp['licenseFamily'] = licenseFamily

				temp['usages'] = rep['items'][i]['usages'][0]
				temp['vulnerabilityURL'] = rep['items'][i]['_meta']['links'][3]['href']
				temp['licenseRisk'] = licenseRisk[i]
				temp['securityRisk'] = securityRisk[i]

				logging.info(temp)
			
				dataset.append(copy.copy(temp))
			
			self.datasetComponent = dataset
			
			return dataset

	def wkComponentGetLicenseFamilty(self, licenseURL, myProject):
		mp = myProject
		url = licenseURL
		logging.debug("license URL IS : " + url)
		s = self.mySession

			#프로젝트 총 개수 구하기
		temp = s.get(url)

		licenseFamily = temp.json()
		return licenseFamily['codeSharing']





	def wkLicenseOpinion(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkLicenseOpinion : Start ####################")

		wb = wkBook
		wk = wkSheet
		mp = myProject

		row_line = self.setSpace


		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 



		#셀 크기 설정
		wk.set_default_row(20)
		wk.set_column(0,0,1)
		wk.set_column(9,9,1)
		wk.set_column(1,1,35)
		wk.set_column(2,2,57)
		wk.set_row(1,row_line)
		
		#보고서 개요

		item1 = "■ 공개SW 라이선스 의견"
		temp1_item1 = {'bold':True, 'font_size':16, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter'}
		cell1_item1 = wb.add_format(temp1_item1)
		wk.write('B1', item1, cell1_item1)



		item1_content1 = "종합의견"
		temp1_content1 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color': self.idColor[0] }
		cell1_content1 = wb.add_format(temp1_content1)
		wk.merge_range('B3:C3', item1_content1, cell1_content1)
	
		wkContent = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent = wb.add_format(wkContent)


		# 라이선스 출력
		dataset = self.wkLicenseOpinionData(mp)


		licenseCategory = ['RECIPROCAL_AGPL', 'RECIPROCAL', 'WEAK_RECIPROCAL']
		license_index = 4
		for i in licenseCategory:
			if i in dataset:
				for j in range(len(dataset[i]['values'])):

					########################### 특정 라이선스 제외 ###########################
					if dataset[i]['values'][j]['label'] =="MIT License":
						continue
					else:
						wk.write('B' + str(license_index), dataset[i]['values'][j]['label'],wkCellContent)
						wk.write('C' + str(license_index), None,wkCellContent)
						license_index +=1



		logging.info("#################### wkLicenseOpinion : Successfully created ####################")

	def wkLicenseOpinionData(self, myProject):
		mp = myProject

		dataset = {}
		if not myProject:
			logging.debug("Sorry I couldn't find it")		
		else:
			url = self.http + "/api/projects/" + mp['projectID'] + "/versions/" + mp['versionID'] + "/components-filters?filterKey=bomLicense"
			
			s = self.mySession
			rep = s.get(url).json()
			print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ API의 값이 아니라 실제 BOM으로 작업해야 오류가 없다. 현재 다른 값을 가져오는 경우가 있음 $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
			print(rep)
			print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")

			for i in range(len(rep['values'])):
				
				if rep['values'][i]['label'] == "PERMISSIVE":
					continue
				elif rep['values'][i]['label'] == "UNKNOWN":
					continue
				elif rep['values'][i]['label'] == "RECIPROCAL_AGPL":
					dataset['RECIPROCAL_AGPL'] = rep['values'][i]

					
				elif rep['values'][i]['label'] == "RECIPROCAL":
					dataset['RECIPROCAL'] = rep['values'][i]


				elif rep['values'][i]['label'] == "WEAK_RECIPROCAL":
					dataset['WEAK_RECIPROCAL'] = rep['values'][i]

			
			
			return dataset		




	def wkSecuritySolution(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkSecuritySolution : Start ####################")

		wb = wkBook
		wk = wkSheet
		mp = myProject

		row_line = self.setSpace

		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 
		wk.set_default_row(20)
		wk.set_column(0,0,1)
		wk.set_column(9,9,1)


		#셀 크기 설정
		wk.set_default_row(20)
		wk.set_column(0,0,1)
		wk.set_column(1,1,30)
		wk.set_column(2,5,13)
		wk.set_column(6,6,1)
		wk.set_row(1,row_line)
		
		#보고서 개요

		item1 = "■ 공개SW 보안취약점 대체 솔루션"
		temp1_item1 = {'bold':True, 'font_size':16, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter'}
		cell1_item1 = wb.add_format(temp1_item1)
		wk.write('B1', item1, cell1_item1)


		# 주의사항
		item1_content1 = ["✓ 보안취약점 개수는 전체 보안취약점으로 산출됩니다", \
						  "✓ 보안취약점 레벨은 여러 등급이 함께 존재할 경우 높은 등급을 우선으로 기재합니다", \
						  "✓ 대체솔루션 버전은 보안취약점이 발견된 상위버전 중 가장 가까운 버전으로 선정되며, 리스크가 없거나 적은 다른 대체버전으로도 선택이 가능합니다. " \
						  ]




		temp1_content1 = {'font_size':12, 'font_name':self.fontName, 'align':'left', 'valign':'top'}
		cell1_content1 = wb.add_format(temp1_content1)
		
		for i in range(len(item1_content1)):
			wk.merge_range('B'+ str(i+3) + ":F" + str(i+3), item1_content1[i], cell1_content1 )

	
		# 보고서 항목


		columnItems = ["컴포넌트명", "이슈버전", "보안취약점 개수", "보안취약점 레벨", "대체솔루션 버전"]
		#riskCategory = ["High", "Medium", "Low", "None"]

		wkTitle = {'bold':True, 'font_size':10, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[0] }
		wkCellTitle = wb.add_format(wkTitle)

		for i in range(len(columnItems)):

			wk.write(6,i+1,columnItems[i], wkCellTitle)


		logging.info("#################### wkSecuritySolution : Successfully created ####################")




	def wkSecuritySolutionData(self, reportName, wkBook, wkSheet, myProject):
		
		rn = reportName
		wb = wkBook
		wk = wkSheet
		mp = myProject
		dataset2=[]
		filteredData = []
		temp = {"componentName":None, \
				"componentID":None, \
				"componentVersionName":None, \
				"componentVersionID":None, \
				"vulnerabilityNum":None, \
				"securityRisk":None, \
				"securitySolution":None \
			    }
		
		# BOM 데이터 : 신뢰 가능함 
		myComponent = self.datasetComponent

		for i in range(len(myComponent)):
			if myComponent[i]['securityRisk'] =="NONE":
				continue
			else:
				filteredData.append(copy.copy(myComponent[i]))

		logging.info("######################### Filtered list #########################")
		
		
		for i in range(len(myComponent)):
			for j in range(len(filteredData)):
				if filteredData[j]['componentName'] in myComponent[i]["componentName"] and filteredData[j]['componentVersionName'] in myComponent[i]["componentVersionName"]:
		
					temp["componentName"] = myComponent[i]["componentName"]
					temp["componentID"] = myComponent[i]["componentURL"].split("/")[5]
					temp["componentVersionName"] = myComponent[i]["componentVersionName"]
					temp["componentVersionID"] = myComponent[i]["componentVersionURL"].split("/")[7]
					temp["securityRisk"] = myComponent[i]["securityRisk"]
					
					dataset2.append(copy.copy(temp))
					print(temp)
				else:
					continue


		logging.info("######################### Security Solution Data #########################")
		logging.info(dataset2)


		return dataset2


		
	def wkSecuritySolution2(self, dataset, wkBook, wkSheet, myProject):
		dt = dataset
		wb = wkBook
		wk = wkSheet
		mp = myProject
		riskCount = 0
		dataset2 =[]
		riskVulnerability = ""
		# 최초 보안취약점 대체 솔루션 코멘트로 인해 10줄 필요
		solutionSpace = 10 
		row_line = self.setSpace
		
		
		wkContent = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent = wb.add_format(wkContent)

		wkRiskHigh = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[1] }
		wkCellRiskHigh = wb.add_format(wkRiskHigh)
		wkRiskMedium = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[2] }
		wkCellRiskMedium = wb.add_format(wkRiskMedium)
		wkRiskLow = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[3] }
		wkCellRiskLow = wb.add_format(wkRiskLow)
		wkRiskNone = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellRiskNone = wb.add_format(wkRiskNone)

		temp ={"componentName": None, "componentVersionName": None}


		# 보안취약점 대체 솔루션
		for i in range(len(dt)):

			solutionSpace +=1

			riskCount = self.wkSecuritySolutionData2(dt[i], mp)


			wk.write(i+7,1,dt[i]["componentName"], wkCellContent)
			wk.write(i+7,2,dt[i]["componentVersionName"], wkCellContent)

			#보안취약점 개수
			wk.write(i+7,3,riskCount, wkCellContent)

			

			if dt[i]['securityRisk'] == "UNKNOWN" or dt[i]['securityRisk'] == "HIGH" or dt[i]['securityRisk'] == "CRITICAL":
				wk.write(i+7,4,dt[i]["securityRisk"], wkCellRiskHigh)
				
			elif dataset[i]['securityRisk'] == "MEDIUM":
				wk.write(i+7,4,dt[i]["securityRisk"], wkCellRiskMedium)
				
			elif dataset[i]['securityRisk'] == "LOW":
				wk.write(i+7,4,dt[i]["securityRisk"], wkCellRiskLow)
				
			elif dataset[i]['securityRisk'] == "NONE":
				wk.write(i+7,4,dt[i]["securityRisk"], wkCellRiskNone)

			# 대체 솔루션 템플릿만 먼저 그리고 처리는 다른 곳에서 한다
			wk.write(i+7,5,None, wkCellContent)



			#보안취약점 솔루션

			#if dt[i]['securityRisk'] == "MEDIUM":
			if dt[i]['securityRisk'] == "UNKNOWN" or dt[i]['securityRisk'] == "HIGH" or dt[i]['securityRisk'] == "CRITICAL":
				temp["componentName"] = dt[i]["componentName"]
				temp["componentVersionName"] = dt[i]["componentVersionName"]


				dataset2.append(copy.copy(temp))

		

		for i in range(len(dataset2)):

			riskVulnerability = riskVulnerability + dataset2[i]["componentName"] + " / "+dataset2[i]["componentVersionName"] + " \n"


		
		item1 = "■ 공개SW 보안취약점 의견"
		temp1_item1 = {'bold':True, 'font_size':16, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter'}
		cell1_item1 = wb.add_format(temp1_item1)
		wk.write('B'+str(solutionSpace), item1, cell1_item1)



		wk.set_row(solutionSpace,row_line)


		item1_content1 = "종합의견"
		temp1_content1 = {'bold':True, 'font_size':12, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color': self.idColor[0] }
		cell1_content1 = wb.add_format(temp1_content1)
	
		wk.merge_range('B'+str(solutionSpace+2)+':F'+str(solutionSpace+2), item1_content1, cell1_content1)



		# 결과에 따라 코멘트를 다르게 만든다
		if riskVulnerability =="" and dataset2 ==[]:
			item1_content3 = "위험도가 존재하는 보안취약점이 없습니다."
	

		elif riskVulnerability =="" and dataset2 !=[]:
			item1_content3 = "중위험도(Medium Risk), 저위험도(Low Risk)로 분류된 컴포넌트는 위험도에 따라 중기조치(6개월 이내)," \
						 "이력관리, 반기별 1회 점검 대상 또는 보호관찰, 장기조치 계획, 이력관리, 연1회 점검시행을 권장드립니다."
		
		else:
			item1_content3 = "해당 컴포넌트는 CVE의 기본점수(Base Score)의 7.0 이상의 점수로 판정된 보안취약점이 포함되어 있으며," \
						 "고위험도(High Risk) 분류된 버전의 오픈소스입니다. 보안취약점에 대한 면밀한 검토와 3개월 이내 조치가 필요하며 이력관리," \
						 "분기별 1회 점검 대상으로 분류하여 지속적인 모니터링을 하시길 권장합니다." \
						 "\n" + "\n" \
						 "그 외 중위험도(Medium Risk), 저위험도(Low Risk)로 분류된 컴포넌트는 위험도에 따라 중기조치(6개월 이내)," \
						 "이력관리, 반기별 1회 점검 대상 또는 보호관찰, 장기조치 계획, 이력관리, 연1회 점검시행을 권장드립니다."
		


		temp1_content2 = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		cell1_content2 = wb.add_format(temp1_content2)

		# 숫자가 하나 적은 것은 인덱스로 카운트하기 때문
		wk.set_row(solutionSpace+2,180)

		wk.merge_range('B'+str(solutionSpace+3)+':C'+str(solutionSpace+3), riskVulnerability,cell1_content2)
		wk.merge_range('D'+str(solutionSpace+3)+':F'+str(solutionSpace+3), item1_content3,cell1_content2)
		
		

	
	
	def wkSecuritySolutionData2(self, dataset, myProject):
		
		dt = dataset

		mp = myProject
		riskCount = 0

		if not myProject:
			logging.debug("Sorry I couldn't find it")		
		else:

			url = self.http + "/api/components/" + dt['componentID'] + "/versions/" + dt['componentVersionID'] + "/risk-profile"

			s = self.mySession
			rep = s.get(url).json()

			logging.info(rep['riskData']['counts'])
			for i in range(len(rep['riskData']['counts'])):
				riskCount += rep['riskData']['counts'][i]['count']

			return riskCount

	def wkSecuritySolutionData3(self, dataset, wkBook, wkSheet, myProject):
		rn = self.reportName
		dt = dataset
		mp = myProject
		wb = wkBook
		wk = wkSheet
		versionList = []
		versionListFiltered = []
		temp = {}
		temp2 = {}
		wait4Sec = ""
		dataset2 =[]

		

		for i in range(len(dt)):
		#totalcount 추출
			url = self.http + "/api/components/" + dt[i]['componentID'] + "/versions?limit="+str(1)
			s = self.mySession
			totalNum_rep = s.get(url).json()
	
			#limit에 전체 버전수 추가하여 리스트 추출
			
			limit = int(totalNum_rep['totalCount'])
			offset = 0
			sort = "versionName%20DESC"
		
			url = self.http + "/api/components/" + dt[i]['componentID'] + "/versions?limit="+str(limit)+"&offset="+str(offset)+"&sort="+sort
			

			rep = s.get(url).json()

			#전체 버전 리스트 추출
		
			for j in range(len(rep['items'])):
				versionList.append(copy.copy(rep['items'][j]['versionName']))

			logging.debug("componentName " + dt[i]['componentName'] + " version name : " + dt[i]['componentVersionName'])

			#버전 리스트에서 인덱스만 추출

			try:
				newLimit = versionList.index(dt[i]['componentVersionName'])
			except Exception as ex:
				print("")
				print("########################################")
				print("Error came up!!")
				logging.error("componentName " + dt[i]['componentName'] + " version name : " + dt[i]['componentVersionName'])
				logging.error(ex)
				print("Probably this version " + dt[i]['componentVersionName'] + " is not in the list " + dt[i]['componentName'] + ". have a look!!")
				print("########################################")
				newLimit = 100

			versionList=[]

			#이 인덱스 상위로 리미트 정하기
			#print("componentVersionName : "+ dt[i]['componentVersionName'] +"New Limit " + str(newLimit))
			# 최신버전에 보안취약점이 있을 경우 대체 솔루션이 없다
			if newLimit == 0:
				temp2['componentName'] = dt[i]['componentName']
				temp2['versionName'] = dt[i]['componentVersionName']
				temp2['alternativeSolution'] = "Not yet"
			else:


				url = self.http + "/api/components/" + dt[i]['componentID'] + "/versions?limit="+str(newLimit)+"&offset="+str(offset)+"&sort="+sort

				rep2 = s.get(url).json()
				for j in range(len(rep2['items'])):

					temp['versionName'] = rep2['items'][j]['versionName']
					temp['versionRisk'] = rep2['items'][j]['_meta']['links'][4]['href']

					versionListFiltered.append(copy.copy(temp))

						
				# 보안취약점 찾기
				temp2['componentName'] = dt[i]['componentName']
				temp2['versionName'] = dt[i]['componentVersionName']
				temp2['alternativeSolution'] = self.wkFindAlternative(versionListFiltered)
			

			wait4Sec = "✈"
			sys.stdout.write(wait4Sec)
			wait4Sec = ""

			
			dataset2.append(copy.copy(temp2))
		
		logging.info("#################### Alternative Solution ####################")
		logging.info(dataset2)

		return dataset2




	def wkFindAlternative(self, versionListFiltered):
		# 여기 들어온 버전 자체는 한번 중간지점을 찝어준 버전이라 루프 카운팅이 작아진다.
		vfl = versionListFiltered
		print("vfl is " + str(vfl))
		tempRiskNum = 0
		s = self.mySession
		alternativeSolution = ""
		print("version number is " + str(len(versionListFiltered)))
		#전체 리스트 몇개
		for i in range(len(versionListFiltered)):
			j = int(len(versionListFiltered) - i - 1)
			print("J is " + str(j))
			url = str(versionListFiltered[j]['versionRisk'])
			print("url is " + str(url))
			rep3 = s.get(url).json()
			print("rep3 is " + str(rep3))


			if rep3['totalCount'] !=0:
				# 이버전의 리스크가 있는지 없는지 확인해야 한다.
				for t in range(rep3['totalCount']):
	
					tempRiskNum = tempRiskNum + rep3['totalCount']

			else:
				tempRiskNum = 0

			#리스크 버전이 0이면 그 버전이 바로 솔루션
			if tempRiskNum == 0:

				tempRiskNum = 0	
				if alternativeSolution =="":
					alternativeSolution = versionListFiltered[j]['versionName']
				else:
					continue

			else:

				tempRiskNum = 0	
			if alternativeSolution !="":
				break

		return alternativeSolution

	def wkSecuritySolution3(self, dataset, wkBook, wkSheet, myProject):

		wb = wkBook
		wk = wkSheet
		mp = myProject
		dt = dataset
		# 최초 보안취약점 대체 솔루션 코멘트로 인해 10줄 필요
		solutionSpace = 10 

		
		wkContent = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent = wb.add_format(wkContent)


		# 보안취약점 대체 솔루션
		for i in range(len(dt)):

			wk.write(i+7,5,dt[i]["alternativeSolution"], wkCellContent)


	
	def wkRiskCategory(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkRiskCategory : Start ####################")
		wb = wkBook
		wk = wkSheet
		mp = myProject

		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 


		dataset = self.wkRiskCategoryData(mp)




		logging.info("#################### wkRiskCategory : Successfully created ####################")



	def wkRiskCategoryData(self, myProject):
		mp = myProject



	def wkTermDefinition(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkTermDefinition : Start ####################")
		wb = wkBook
		wk = wkSheet
		mp = myProject

		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 


		dataset = self.wkTermDefinitionData(mp)

		logging.info("#################### wkTermDefinition : Successfully created ####################")


	def wkTermDefinitionData(self, myProject):
		mp = myProject



	def wkVulnerabilities(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkVulnerabilities : Start ####################")
		wb = wkBook
		wk = wkSheet
		mp = myProject

		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 

		columnItems = ["컴포넌트", "버전", "보안취약점 출처", "보안취약점", "베이스 스코어", "상세설명"]
		

		wkTitle = {'bold':True, 'font_size':10, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[0] }
		wkCellTitle = wb.add_format(wkTitle)

		for i in range(len(columnItems)):
			wk.write(0, i, columnItems[i], wkCellTitle)
		

		wk.set_default_row(25)
		wk.set_column(0,0,50)
		wk.set_column(1,1,15)
		wk.set_column(2,2,15)
		wk.set_column(3,3,25)
		wk.set_column(4,4,15)
		wk.set_column(5,5,150)

		dt = self.wkVulnerabilitiesData(wb, wk, mp)



		wkContent1 = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent1 = wb.add_format(wkContent1)

		wkContent2 = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent2 = wb.add_format(wkContent2)


		for i in range(len(dt)):
			wk.write('A'+ str(i+2), dt[i]['componentName'], wkCellContent1)	
			wk.write('B'+ str(i+2), dt[i]['componentVersionName'], wkCellContent1)	
			wk.write('C'+ str(i+2), dt[i]['source'], wkCellContent1)	
			wk.write('D'+ str(i+2), dt[i]['vulnerabilityName'], wkCellContent1)	
			wk.write('E'+ str(i+2), dt[i]['baseScore'], wkCellContent1)	
			wk.write('F'+ str(i+2), dt[i]['description'], wkCellContent2)

		logging.info("#################### wkVulnerabilities : Successfully created ####################")


	def wkVulnerabilitiesData(self, wkBook, wkSheet, myProject):
		wb = wkBook
		wk = wkSheet
		mp = myProject
		myComponent = self.datasetComponent
		dataset = []

		temp = {}
		temp2 = {}

		limit = 0
		offset = 0
		
		

		for i in range(len(myComponent)):
			
			temp["componentName"] = myComponent[i]["componentName"]
			temp["componentID"] = myComponent[i]["componentURL"].split("/")[5]
			temp["componentVersionName"] = myComponent[i]["componentVersionName"]
			temp["componentVersionID"] = myComponent[i]["componentVersionURL"].split("/")[7]


			if not temp['componentID']:
				logging.debug("Sorry I couldn't find it")		
			else:
				
				url = self.http + "/api/components/" + temp["componentID"] + "/versions/" + temp["componentVersionID"] + "/vulnerabilities"
				s = self.mySession
				payload = {'limit':1, 'offset':offset}
				tempRep = s.get(url, params=payload).json()
				
				limit = tempRep['totalCount']
				payload = {'limit':limit, 'offset':offset}
				rep = s.get(url, params = payload).json()
				
			

				if rep['items']:
					for j in range(len(rep['items'])):
					
						temp2['componentName'] = temp['componentName']
						temp2['componentVersionName'] = temp['componentVersionName']
						temp2['source'] = rep['items'][j]['source']
						temp2['vulnerabilityName'] = rep['items'][j]['vulnerabilityName']
						temp2['baseScore'] = rep['items'][j]['baseScore']
						temp2['description'] = rep['items'][j]['description']

						
						dataset.append(copy.copy(temp2))
						logging.info("######################### vulnerability Data No." + str(i) + " #########################")
						logging.info(temp2)
				else:
					continue

		return dataset




	def wkSourcePath(self, wkBook, wkSheet, myProject):
		logging.info("#################### wkSourcePath : Start ####################")
		wb = wkBook
		wk = wkSheet
		mp = myProject

		logging.info("current sheet is "  + wk.get_name())


		#################### 각 워크 시트 생성 및 기본 설정 #################### 

		columnItems = ["컴포넌트", "버전", "결합형태", "매치파일경로"]
		

		wkTitle = {'bold':True, 'font_size':10, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[0] }
		wkCellTitle = wb.add_format(wkTitle)

		for i in range(len(columnItems)):
			wk.write(0, i, columnItems[i], wkCellTitle)
		

		wk.set_default_row(25)
		wk.set_column(0,0,50)
		wk.set_column(1,1,15)
		wk.set_column(2,2,15)
		wk.set_column(3,3,160)
	

		dt = self.wkSourcePathData(wb, wk, mp)



		wkContent1 = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'center', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent1 = wb.add_format(wkContent1)

		wkContent2 = {'bold':False, 'font_size':8, 'font_name':self.fontName, 'align':'left', 'valign':'vcenter', 'border':1, 'border_color':'#000000', 'bg_color':self.idColor[4] }
		wkCellContent2 = wb.add_format(wkContent2)

		if dt:
			for i in range(len(dt)):
				wk.write('A'+ str(i+2), dt[i]['componentName'], wkCellContent1)	
				wk.write('B'+ str(i+2), dt[i]['componentVersionName'], wkCellContent1)	
				wk.write('C'+ str(i+2), dt[i]['matchType'], wkCellContent1)	
				wk.write('D'+ str(i+2), dt[i]['analysisPath'], wkCellContent2)	
			
			

		logging.info("#################### wkSourcePath : Successfully created ####################")


	def wkSourcePathData(self, wkBook, wkSheet, myProject):
		wb = wkBook
		wk = wkSheet
		mp = myProject
		temPath = ""
		tempURL = ""
		tempComponentID = ""
		tempComponentVersionID = ""
		temp={}
		dataset = []
		limit = 0
		url = self.http + "/api/projects/" + mp['projectID'] + "/versions/" + mp['versionID'] + "/matched-files?limit=1"
		s = self.mySession
		tempRep = s.get(url).json()
		limit = tempRep['totalCount']
		
		temp={'componentName':None, 'componentVersionName':None, 'matchType':None, 'usage':None, 'analysisPath':None}

		if limit == 0:
			logging.debug("Sorry I couldn't find it")		
		else:
		
			url = self.http + "/api/projects/" + mp['projectID'] + "/versions/" + mp['versionID'] + "/matched-files?limit=" + str(limit)
		
			rep = s.get(url).json()
			
			for i in range(len(rep['items'])):

				temp['matchType'] = rep['items'][i]['matches'][0]['matchType']
				# 컴포넌트  추출
				tempURL = rep['items'][i]['matches'][0]['component']
				tempComponentID = rep['items'][i]['matches'][0]['component'].split("/")[5]

				url = "https://192.168.0.18/api/components/" + tempComponentID
				tempRep2 = s.get(url).json()
				temp["componentName"] = tempRep2['name']


				# 버전 추출
				if len(rep['items'][i]['matches'][0]['component'].split("/")) <= 6:
					temp["componentVersionName"] = None
	
				else:

					tempComponentVersionID = rep['items'][i]['matches'][0]['component'].split("/")[7]
					url = "https://192.168.0.18/api/components/92a62dae-28ba-467b-a999-60e889d11a58/versions/" + tempComponentVersionID
					tempRep2 = s.get(url).json()
					try:		
						temp['componentVersionName']=tempRep2['versionName']
					except KeyError as ex:
						temp['componentVersionName']:None


				
				# 분석경로 찾기 3번은 무조건 디코딩 해줘야함
				tempPath = rep['items'][i]['uri']



				for i in (range(3)):
					tempPath = urllib.parse.unquote(tempPath)
					if i == 2:
						temp['analysisPath'] = tempPath.split('file:///')[1]
					else:
						continue

				

				logging.info(temp)

				dataset.append(copy.copy(temp))

			return dataset
	

		
logging.basicConfig(level=logging.DEBUG)
logging.disable(logging.DEBUG)
dataset = []

# 엑셀표 참조해서 돌릴 때
#br = blackduckRPT("https://192.168.0.18","kep","changeme")
#dataset = br.getCSV()
#for i in range(len(dataset)):

#	myProject = br.findIdentity(dataset[i]['projectID'],dataset[i]['versionID'])
#	br.createExcel(myProject)


#한개만 돌릴 떄 
br = blackduckRPT("https://192.168.0.18","junsulee","blackduck")
myProject = br.findIdentity("K-MDMS-MDA","Default Detect Version")
br.createExcel(myProject)