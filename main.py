import requests
import json
from openpyxl import load_workbook

def login(req, id: str, pw: str):
	login = req.post('https://reading.gne.go.kr/r/newReading/member/login.jsp', 
		headers={'content-type':'application/x-www-form-urlencoded'},
		data={'msg':'','userid':id,'passwd':pw,'authType':'','authId':'','s_id':id,'s_pwd':pw})
	return login.status_code == '200'

def register(req, bname, bauthor, bpub, byear):
	url = 'https://reading.gne.go.kr/r/newReading/readbook/userBookRegist.jsp'
	headers = {'content-type':'application/x-www-form-urlencoded; charset=UTF-8'}
	data = {'bookNm':bname,'writer1':bauthor,'writer2':'','writer3':'','publisher':bpub,'publicationYy':byear,'maskingYn':'Y'}
	reg = req.post(url, headers=headers, data=data)
	result = json.loads(reg.content.decode())
	del(result['result'])
	return(result)

def readForm(req, bookInfo, title, content, contentLength):
	url = 'https://reading.gne.go.kr/r/newReading/readbook/writeImpression.jsp'
	headers = {'content-type':'application/x-www-form-urlencoded; charset=UTF-8'}
	dateYYMMDD = '20210726'
	data = {'subjectSerial': '',
		'writeSelectBox': 'writeImpression',
		'title': title,
		'content': content,
		'contentLength': contentLength,
		'portfolioFlag': 'N',
		'endFlag': 'Y',
		'bookSerial': bookInfo['userBookSerial'],
		'bookCd': bookInfo['bookCd'],
		'subjectSerial': '' ,
		'valuationDate': dateYYMMDD,
		'editorType': 'readbook',
		'minLength': '10',
		'pasteYn': 'N'}
	read = req.post(url, headers=headers, data=data)
	return read.content.decode()

def loadExcel():
	load_wb = load_workbook("bookList.xlsx", data_only=True)
	load_ws = load_wb['Sheet1']

	lname = []
	lauthor = []
	lpub = []
	lyear = []
	lcontent = []

	for cell in load_ws['A']: 
		lname.append(cell.value)
	for cell in load_ws['B']: 
		lauthor.append(cell.value)
	for cell in load_ws['C']: 
		lpub.append(cell.value)
	for cell in load_ws['D']: 
		lyear.append(cell.value)
	for cell in load_ws['E']: 
		lcontent.append(cell.value)

	ret = []
	for i in range(len(lname)):
		ret.append({'bname':lname[i],'bauthor':lauthor[i], 'bpub': lpub[i], 'byear': lyear[i], 'content':lcontent[i], })
	return ret

def main():
	req = requests.session()
	id = ''
	pw = ''
	login(req, id, pw)

	data = loadExcel()
	print(len(data))

	for book in data:
		bookInfo = register(req, book['bname'], book['bauthor'], book['bpub'], book['byear'])
		title = book['bname'] + '을 읽고'
		resullt = readForm(req, bookInfo, title, book['content'], len(book['content']))
		print(resullt)

if __name__=="__main__":
	main()
