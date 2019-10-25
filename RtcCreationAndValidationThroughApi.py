import requests
import json
import datetime

import xlsxwriter
import time
import xlrd
import xlwt
from xlwt import Workbook

#URls

base_url = "https://amsin.hirepro.in/py/common/user/login_user/"
create_question_url = "https://amsin.hirepro.in/py/assessment/authoring/api/v1/createQuestion/"
getQuestionForld_url = "https://amsin.hirepro.in/py/assessment/authoring/api/v1/getQuestionForId/"
xlsfileloc = ("C:\\Users\\con661\\Desktop\\rtcexelinputdata.xls")
xls_report_1 = ("C:\\Users\\con661\\Desktop\\AuthoringMcqReport1.xls")



class CreateMcqQuestion:

# CRPO login application

    def __init__(self):
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y")
        self.rowsize = 1

        # CRPO LOGIN APPLICATION
        self.header = {"content-type": "application/json"}
        self.data = {"LoginName":"",
                      "Password":"",
                      "TenantAlias":"",
                      "UserName":""}
        response = requests.post(base_url, data=json.dumps(self.data, default=str))
        self.abc = response.json()
        self.headers = {"content-type": "application/json", "X-AUTH-TOKEN": self.abc.get("Token")}

        #print Exel Header

        self.exelHeaderMaker()
        file_path_report1 =("D:\AtAutomation1\PythonScripts\WebScraping\AuthoringQP1\PyScripts1\CreateQuestion\exelReport2.xls")






# Creating question through api reqest

        # To open Workbook
        wb = xlrd.open_workbook(xlsfileloc)
        sheet = wb.sheet_by_index(0)
        # total_rows = sheet.nrows
        #print(sheet.nrows)
        #print(sheet.row_values)
        a = list()
        for b in range(0, (sheet.nrows)):
            excel_col_val = sheet.row_values(b)
            a.append({"questionType": int((excel_col_val[0])),
                      "difficultyLevel": int((excel_col_val[1])),
                      "subCategoryId": int((excel_col_val[2])),
                      "categoryId": int((excel_col_val[3])),
                      "htmlString": (excel_col_val[4]),
                      "questionStr": (excel_col_val[5]),
                      "authorId": int((excel_col_val[6])),
                      "isRevised": False,
                      "notes": (excel_col_val[8]),
                      "statusId": int((excel_col_val[9])),
                      "questionFlag": int((excel_col_val[10])),
                      "childQuestions": [{"questionType": int((excel_col_val[11])),
                                          "difficultyLevel": int((excel_col_val[12])),
                                          "subCategoryId": int((excel_col_val[13])),
                                         "categoryId": int((excel_col_val[14])),
                                          "htmlString": (excel_col_val[15]),
                                          "questionStr": (excel_col_val[16]),
                                          "authorId": int((excel_col_val[17])),
                                          "isRevised": False,
                                          "notes": (excel_col_val[19]),
                                          "statusId": int((excel_col_val[20])),
                                          "questionFlag": int((excel_col_val[21])),
                                          "answers": [{"htmlString": (excel_col_val[22]), "correctAnswer": (excel_col_val[23])}],
                                          "answerChoices": [{"htmlString": (excel_col_val[24]), "choice": (excel_col_val[25])},
                                                            {"htmlString": (excel_col_val[26]), "choice": (excel_col_val[27])},
                                                            {"htmlString": (excel_col_val[28]), "choice": (excel_col_val[29])},
                                                            {"htmlString": (excel_col_val[30]), "choice": (excel_col_val[31])}]},
                                         {"questionType": int((excel_col_val[32])),
                                          "difficultyLevel": int((excel_col_val[33])),
                                          "subCategoryId": int((excel_col_val[34])),
                                          "categoryId": int((excel_col_val[35])),
                                          "htmlString": (excel_col_val[36]),
                                          "questionStr": (excel_col_val[37]),
                                          "authorId": int((excel_col_val[38])),
                                          "isRevised": False,
                                          "notes": (excel_col_val[40]),
                                          "statusId": int((excel_col_val[41])),
                                          "questionFlag": int((excel_col_val[42])),
                                          "answers": [{"htmlString": (excel_col_val[43]), "correctAnswer": (excel_col_val[44])}],
                                          "answerChoices": [{"htmlString": (excel_col_val[45]), "choice": (excel_col_val[46])},
                                                            {"htmlString": (excel_col_val[47]), "choice": (excel_col_val[48])},
                                                            {"htmlString": (excel_col_val[49]), "choice": (excel_col_val[50])},
                                                            {"htmlString": (excel_col_val[51]), "choice": (excel_col_val[52])}]},
                                         {"questionType": int((excel_col_val[53])),
                                          "difficultyLevel": int((excel_col_val[54])),
                                          "subCategoryId": int((excel_col_val[55])),
                                          "categoryId": int((excel_col_val[56])),
                                          "htmlString": (excel_col_val[57]),
                                          "questionStr": (excel_col_val[58]),
                                          "authorId": int((excel_col_val[59])),
                                          "isRevised": False,
                                          "notes": (excel_col_val[61]),
                                          "statusId": int((excel_col_val[62])),
                                          "questionFlag": int((excel_col_val[63])),
                                          "answers": [{"htmlString": (excel_col_val[64]), "correctAnswer": (excel_col_val[65])}],
                                          "answerChoices": [{"htmlString": (excel_col_val[66]), "choice": (excel_col_val[67])},
                                                            {"htmlString": (excel_col_val[68]), "choice": (excel_col_val[69])},
                                                            {"htmlString": (excel_col_val[70]), "choice": (excel_col_val[71])},
                                                            {"htmlString": (excel_col_val[72]), "choice": (excel_col_val[73])}]},
                                         {"questionType": int((excel_col_val[74])),
                                          "difficultyLevel": int((excel_col_val[75])),
                                          "subCategoryId": int((excel_col_val[76])),
                                          "categoryId": int((excel_col_val[77])),
                                          "htmlString": (excel_col_val[78]),
                                          "questionStr": (excel_col_val[79]),
                                          "authorId": int((excel_col_val[80])),
                                          "isRevised": False,
                                          "notes": (excel_col_val[82]),
                                          "statusId": int((excel_col_val[83])),
                                          "questionFlag": int((excel_col_val[84])),
                                          "answers": [{"htmlString": (excel_col_val[85]), "correctAnswer": (excel_col_val[86])}],
                                          "answerChoices": [{"htmlString": (excel_col_val[87]), "choice": (excel_col_val[88])},
                                                            {"htmlString": (excel_col_val[89]), "choice": (excel_col_val[90])},
                                                            {"htmlString": (excel_col_val[91]), "choice": (excel_col_val[92])},
                                                            {"htmlString": (excel_col_val[93]), "choice": (excel_col_val[94])}]}]})


#            print(a[b])


            responseFromCreatequestion = requests.post(create_question_url, headers=self.headers,
                                      data=json.dumps(a[b], default=str))
            self.abcd = responseFromCreatequestion.json()

            print(self.abcd)

            if (self.abcd['status'] == 'OK'):

                varGetQuestionForldInput = {"id": (self.abcd['data']['questionId'])}

                responseFormGetQuestionForldResponse = requests.post(getQuestionForld_url, headers=self.headers,
                                                                     data=json.dumps(varGetQuestionForldInput,
                                                                                     default=str))
                responseFormGetQuestionForldResponse_1 = responseFormGetQuestionForldResponse.json()

                if (responseFormGetQuestionForldResponse_1['statusId'] == 200):
                    self.ws.write(((2 * b) + 2), 1, 'Input', self.style0)

                    # Printing Input value
                    # parent

                    self.ws.write(((2 * b) + 2), 2, int((excel_col_val[0])), self.style5)
                    self.ws.write(((2 * b) + 2), 3, int((excel_col_val[1])), self.style5)
                    self.ws.write(((2 * b) + 2), 4, int((excel_col_val[2])), self.style5)
                    self.ws.write(((2 * b) + 2), 5, int((excel_col_val[3])), self.style5)
                    self.ws.write(((2 * b) + 2), 6, (excel_col_val[4]), self.style5)
                    self.ws.write(((2 * b) + 2), 7, (excel_col_val[5]), self.style5)
                    self.ws.write(((2 * b) + 2), 8, int((excel_col_val[6])), self.style5)
                    self.ws.write(((2 * b) + 2), 9, False, self.style5)
                    self.ws.write(((2 * b) + 2), 10, (excel_col_val[8]), self.style5)
                    self.ws.write(((2 * b) + 2), 11, int((excel_col_val[9])), self.style5)
                    self.ws.write(((2 * b) + 2), 12, int((excel_col_val[10])), self.style5)

                    #child1

                    self.ws.write(((2 * b) + 2), 13, int((excel_col_val[11])), self.style5)
                    self.ws.write(((2 * b) + 2), 14, int((excel_col_val[12])), self.style5)
                    self.ws.write(((2 * b) + 2), 15, int((excel_col_val[13])), self.style5)
                    self.ws.write(((2 * b) + 2), 16, int((excel_col_val[14])), self.style5)
                    self.ws.write(((2 * b) + 2), 17, (excel_col_val[15]), self.style5)
                    self.ws.write(((2 * b) + 2), 18, (excel_col_val[16]), self.style5)
                    self.ws.write(((2 * b) + 2), 19, int((excel_col_val[17])), self.style5)
                    self.ws.write(((2 * b) + 2), 20, False, self.style5)
                    self.ws.write(((2 * b) + 2), 21, (excel_col_val[19]), self.style5)
                    self.ws.write(((2 * b) + 2), 22, int((excel_col_val[20])), self.style5)
                    self.ws.write(((2 * b) + 2), 23, (excel_col_val[21]), self.style5)

                    self.ws.write(((2 * b) + 2), 24, (excel_col_val[22]), self.style5)
                    self.ws.write(((2 * b) + 2), 25, (excel_col_val[23]), self.style5)

                    self.ws.write(((2 * b) + 2), 26, (excel_col_val[24]), self.style5)
                    self.ws.write(((2 * b) + 2), 27, (excel_col_val[25]), self.style5)
                    self.ws.write(((2 * b) + 2), 28, (excel_col_val[26]), self.style5)
                    self.ws.write(((2 * b) + 2), 29, (excel_col_val[27]), self.style5)
                    self.ws.write(((2 * b) + 2), 30, (excel_col_val[28]), self.style5)
                    self.ws.write(((2 * b) + 2), 31, (excel_col_val[29]), self.style5)
                    self.ws.write(((2 * b) + 2), 32, (excel_col_val[30]), self.style5)
                    self.ws.write(((2 * b) + 2), 33, (excel_col_val[31]), self.style5)

                    #child 2

                    self.ws.write(((2 * b) + 2), 34, int((excel_col_val[32])), self.style5)
                    self.ws.write(((2 * b) + 2), 35, int((excel_col_val[33])), self.style5)
                    self.ws.write(((2 * b) + 2), 36, int((excel_col_val[34])), self.style5)
                    self.ws.write(((2 * b) + 2), 37, int((excel_col_val[35])), self.style5)
                    self.ws.write(((2 * b) + 2), 38, (excel_col_val[36]), self.style5)
                    self.ws.write(((2 * b) + 2), 39, (excel_col_val[37]), self.style5)
                    self.ws.write(((2 * b) + 2), 40, int((excel_col_val[38])), self.style5)
                    self.ws.write(((2 * b) + 2), 41, False, self.style5)
                    self.ws.write(((2 * b) + 2), 42, (excel_col_val[40]), self.style5)
                    self.ws.write(((2 * b) + 2), 43, int((excel_col_val[41])), self.style5)
                    self.ws.write(((2 * b) + 2), 44, (excel_col_val[42]), self.style5)

                    self.ws.write(((2 * b) + 2), 45, (excel_col_val[43]), self.style5)
                    self.ws.write(((2 * b) + 2), 46, (excel_col_val[44]), self.style5)

                    self.ws.write(((2 * b) + 2), 47, (excel_col_val[45]), self.style5)
                    self.ws.write(((2 * b) + 2), 48, (excel_col_val[46]), self.style5)
                    self.ws.write(((2 * b) + 2), 49, (excel_col_val[47]), self.style5)
                    self.ws.write(((2 * b) + 2), 50, (excel_col_val[48]), self.style5)
                    self.ws.write(((2 * b) + 2), 51, (excel_col_val[49]), self.style5)
                    self.ws.write(((2 * b) + 2), 52, (excel_col_val[50]), self.style5)
                    self.ws.write(((2 * b) + 2), 53, (excel_col_val[51]), self.style5)
                    self.ws.write(((2 * b) + 2), 54, (excel_col_val[52]), self.style5)

                    #child3

                    self.ws.write(((2 * b) + 2), 55, int((excel_col_val[53])), self.style5)
                    self.ws.write(((2 * b) + 2), 56, int((excel_col_val[54])), self.style5)
                    self.ws.write(((2 * b) + 2), 57, int((excel_col_val[55])), self.style5)
                    self.ws.write(((2 * b) + 2), 58, int((excel_col_val[56])), self.style5)
                    self.ws.write(((2 * b) + 2), 59, (excel_col_val[57]), self.style5)
                    self.ws.write(((2 * b) + 2), 60, (excel_col_val[58]), self.style5)
                    self.ws.write(((2 * b) + 2), 61, int((excel_col_val[59])), self.style5)
                    self.ws.write(((2 * b) + 2), 62, False, self.style5)
                    self.ws.write(((2 * b) + 2), 63, (excel_col_val[61]), self.style5)
                    self.ws.write(((2 * b) + 2), 64, int((excel_col_val[62])), self.style5)
                    self.ws.write(((2 * b) + 2), 65, (excel_col_val[63]), self.style5)

                    self.ws.write(((2 * b) + 2), 66, (excel_col_val[64]), self.style5)
                    self.ws.write(((2 * b) + 2), 67, (excel_col_val[65]), self.style5)

                    self.ws.write(((2 * b) + 2), 68, (excel_col_val[66]), self.style5)
                    self.ws.write(((2 * b) + 2), 69, (excel_col_val[67]), self.style5)
                    self.ws.write(((2 * b) + 2), 70, (excel_col_val[68]), self.style5)
                    self.ws.write(((2 * b) + 2), 71, (excel_col_val[69]), self.style5)
                    self.ws.write(((2 * b) + 2), 72, (excel_col_val[70]), self.style5)
                    self.ws.write(((2 * b) + 2), 73, (excel_col_val[71]), self.style5)
                    self.ws.write(((2 * b) + 2), 74, (excel_col_val[72]), self.style5)
                    self.ws.write(((2 * b) + 2), 75, (excel_col_val[73]), self.style5)

                    #child4

                    self.ws.write(((2 * b) + 2), 76, int((excel_col_val[74])), self.style5)
                    self.ws.write(((2 * b) + 2), 77, int((excel_col_val[75])), self.style5)
                    self.ws.write(((2 * b) + 2), 78, int((excel_col_val[76])), self.style5)
                    self.ws.write(((2 * b) + 2), 79, int((excel_col_val[77])), self.style5)
                    self.ws.write(((2 * b) + 2), 80, (excel_col_val[78]), self.style5)
                    self.ws.write(((2 * b) + 2), 81, (excel_col_val[79]), self.style5)
                    self.ws.write(((2 * b) + 2), 82, int((excel_col_val[80])), self.style5)
                    self.ws.write(((2 * b) + 2), 83, False, self.style5)
                    self.ws.write(((2 * b) + 2), 84, (excel_col_val[82]), self.style5)
                    self.ws.write(((2 * b) + 2), 85, int((excel_col_val[83])), self.style5)
                    self.ws.write(((2 * b) + 2), 86, (excel_col_val[84]), self.style5)

                    self.ws.write(((2 * b) + 2), 87, (excel_col_val[85]), self.style5)
                    self.ws.write(((2 * b) + 2), 88, (excel_col_val[86]), self.style5)

                    self.ws.write(((2 * b) + 2), 89, (excel_col_val[87]), self.style5)
                    self.ws.write(((2 * b) + 2), 90, (excel_col_val[88]), self.style5)
                    self.ws.write(((2 * b) + 2), 91, (excel_col_val[89]), self.style5)
                    self.ws.write(((2 * b) + 2), 92, (excel_col_val[90]), self.style5)
                    self.ws.write(((2 * b) + 2), 93, (excel_col_val[91]), self.style5)
                    self.ws.write(((2 * b) + 2), 94, (excel_col_val[92]), self.style5)
                    self.ws.write(((2 * b) + 2), 95, (excel_col_val[93]), self.style5)
                    self.ws.write(((2 * b) + 2), 96, (excel_col_val[94]), self.style5)

                    # Printing Output value

                    self.ws.write(((2 * b) + 3), 0, 'Pass', self.style7)

                    self.ws.write(((2 * b) + 3), 1, 'Output', self.style1)

                    #Parent question validation

                    if ((int(excel_col_val[0])) == (responseFormGetQuestionForldResponse_1['data']['questionType'])):
                        self.ws.write(((2 * b) + 3), 2, responseFormGetQuestionForldResponse_1['data']['questionType'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 2, responseFormGetQuestionForldResponse_1['data']['questionType'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[1])) == (responseFormGetQuestionForldResponse_1['data']['difficultyLevel'])):
                        self.ws.write(((2 * b) + 3), 3,
                                      responseFormGetQuestionForldResponse_1['data']['difficultyLevel'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 3,
                                      responseFormGetQuestionForldResponse_1['data']['difficultyLevel'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[2])) == (responseFormGetQuestionForldResponse_1['data']['subCategoryId'])):
                        self.ws.write(((2 * b) + 3), 4, responseFormGetQuestionForldResponse_1['data']['subCategoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 4, responseFormGetQuestionForldResponse_1['data']['subCategoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[3])) == (responseFormGetQuestionForldResponse_1['data']['categoryId'])):
                        self.ws.write(((2 * b) + 3), 5, responseFormGetQuestionForldResponse_1['data']['categoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 5, responseFormGetQuestionForldResponse_1['data']['categoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[4]) == (responseFormGetQuestionForldResponse_1['data']['htmlString'])):
                        self.ws.write(((2 * b) + 3), 6, responseFormGetQuestionForldResponse_1['data']['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 6, responseFormGetQuestionForldResponse_1['data']['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[5]) == (responseFormGetQuestionForldResponse_1['data']['questionStr'])):
                        self.ws.write(((2 * b) + 3), 7, responseFormGetQuestionForldResponse_1['data']['questionStr'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 7, responseFormGetQuestionForldResponse_1['data']['questionStr'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[6])) == (responseFormGetQuestionForldResponse_1['data']['authorId'])):
                        self.ws.write(((2 * b) + 3), 8, responseFormGetQuestionForldResponse_1['data']['authorId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 8, responseFormGetQuestionForldResponse_1['data']['authorId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if (False == (responseFormGetQuestionForldResponse_1['data']['isRevised'])):
                        self.ws.write(((2 * b) + 3), 9, responseFormGetQuestionForldResponse_1['data']['isRevised'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 9, responseFormGetQuestionForldResponse_1['data']['isRevised'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[8]) == (responseFormGetQuestionForldResponse_1['data']['notes'])):
                        self.ws.write(((2 * b) + 3), 10, responseFormGetQuestionForldResponse_1['data']['notes'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 10, responseFormGetQuestionForldResponse_1['data']['notes'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)


                    if ((int(excel_col_val[9])) == (responseFormGetQuestionForldResponse_1['data']['statusId'])):
                        self.ws.write(((2 * b) + 3), 11, responseFormGetQuestionForldResponse_1['data']['statusId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 11, responseFormGetQuestionForldResponse_1['data']['statusId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[10])) == (responseFormGetQuestionForldResponse_1['data']['questionFlagId'])):
                        self.ws.write(((2 * b) + 3), 12,
                                      responseFormGetQuestionForldResponse_1['data']['questionFlagId'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 12,
                                      responseFormGetQuestionForldResponse_1['data']['questionFlagId'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)





























                    #Child question 1 validation


                    if ((int(excel_col_val[11])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionType'])):
                        self.ws.write(((2 * b) + 3), 13, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionType'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 13, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionType'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[12])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['difficultyLevel'])):
                        self.ws.write(((2 * b) + 3), 14,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['difficultyLevel'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 14,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['difficultyLevel'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[13])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['subCategoryId'])):
                        self.ws.write(((2 * b) + 3), 15, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['subCategoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 15, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['subCategoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[14])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['categoryId'])):
                        self.ws.write(((2 * b) + 3), 16, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['categoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 16, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['categoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[15]) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 17, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 17, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[16]) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionStr'])):
                        self.ws.write(((2 * b) + 3), 18, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionStr'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 18, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionStr'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[17])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['authorId'])):
                        self.ws.write(((2 * b) + 3), 19, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['authorId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 19, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['authorId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if (False == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['isRevised'])):
                        self.ws.write(((2 * b) + 3), 20, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['isRevised'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 20, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['isRevised'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[19]) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['notes'])):
                        self.ws.write(((2 * b) + 3), 21, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['notes'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 21, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['notes'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)


                    if ((int(excel_col_val[20])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['statusId'])):
                        self.ws.write(((2 * b) + 3), 22, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['statusId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 22, responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['statusId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[21])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionFlagId'])):
                        self.ws.write(((2 * b) + 3), 23,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionFlagId'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 23,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['questionFlagId'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    #child1 correct answer

                    if ((excel_col_val[22]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answers'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 24,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answers'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 24,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answers'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[23]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answers'][0]['correctAnswer'])):
                        self.ws.write(((2 * b) + 3), 25,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answers'][0]['correctAnswer'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 25,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answers'][0]['correctAnswer'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    #child1 options

                    if ((excel_col_val[24]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 26,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 26,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[25]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][0]['choice'])):
                        self.ws.write(((2 * b) + 3), 27,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][0]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 27,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][0]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[26]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][1]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 28,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][1]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 28,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][1]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[27]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][1]['choice'])):
                        self.ws.write(((2 * b) + 3), 29,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][1]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 29,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][1]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[28]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][2]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 30,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][2]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 30,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][2]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[29]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][2]['choice'])):
                        self.ws.write(((2 * b) + 3), 31,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][2]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 31,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][2]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[30]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][3]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 32,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][3]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 32,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][3]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[31]) == (
                    responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][3]['choice'])):
                        self.ws.write(((2 * b) + 3), 33,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][3]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 33,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][0]['answerChoices'][3]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

















                    # Child question 2 validation

                    if ((int(excel_col_val[32])) == (responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['questionType'])):
                        self.ws.write(((2 * b) + 3), 34,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'questionType'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 34,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'questionType'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[33])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                'difficultyLevel'])):
                        self.ws.write(((2 * b) + 3), 35,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'difficultyLevel'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 35,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'difficultyLevel'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[34])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['subCategoryId'])):
                        self.ws.write(((2 * b) + 3), 36,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'subCategoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 36,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'subCategoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[35])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['categoryId'])):
                        self.ws.write(((2 * b) + 3), 37,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'categoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 37,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'categoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[36]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 38,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 38,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[37]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['questionStr'])):
                        self.ws.write(((2 * b) + 3), 39,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'questionStr'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 39,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'questionStr'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[38])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['authorId'])):
                        self.ws.write(((2 * b) + 3), 40,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'authorId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 40,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'authorId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if (False == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['isRevised'])):
                        self.ws.write(((2 * b) + 3), 41,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'isRevised'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 41,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'isRevised'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[40]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['notes'])):
                        self.ws.write(((2 * b) + 3), 42,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['notes'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 42,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['notes'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[41])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['statusId'])):
                        self.ws.write(((2 * b) + 3), 43,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'statusId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 43,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'statusId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[42])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['questionFlagId'])):
                        self.ws.write(((2 * b) + 3), 44,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'questionFlagId'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 44,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'questionFlagId'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    # child2 correct answer

                    if ((excel_col_val[43]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answers'][0][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 45,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answers'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 45,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answers'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[44]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answers'][0][
                                'correctAnswer'])):
                        self.ws.write(((2 * b) + 3), 46,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answers'][0]['correctAnswer'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 46,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answers'][0]['correctAnswer'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    # child1 options

                    if ((excel_col_val[45]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][0][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 47,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 47,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[46]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][0][
                                'choice'])):
                        self.ws.write(((2 * b) + 3), 48,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][0]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 48,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][0]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[47]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][1][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 49,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][1]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 49,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][1]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[48]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][1][
                                'choice'])):
                        self.ws.write(((2 * b) + 3), 50,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][1]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 50,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][1]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[49]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][2][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 51,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][2]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 51,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][2]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[50]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][2][
                                'choice'])):
                        self.ws.write(((2 * b) + 3), 52,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][2]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 52,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][2]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[51]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][3][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 53,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][3]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 53,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][3]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[52]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][1]['answerChoices'][3][
                                'choice'])):
                        self.ws.write(((2 * b) + 3), 54,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][3]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 54,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][1][
                                          'answerChoices'][3]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)












                    # Child question 3 validation

                    if ((int(excel_col_val[53])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'questionType'])):
                        self.ws.write(((2 * b) + 3), 55,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'questionType'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 55,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'questionType'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[54])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'difficultyLevel'])):
                        self.ws.write(((2 * b) + 3), 56,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'difficultyLevel'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 56,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'difficultyLevel'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[55])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'subCategoryId'])):
                        self.ws.write(((2 * b) + 3), 57,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'subCategoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 57,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'subCategoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[56])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['categoryId'])):
                        self.ws.write(((2 * b) + 3), 58,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'categoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 58,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'categoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[57]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 59,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 59,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[58]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'questionStr'])):
                        self.ws.write(((2 * b) + 3), 60,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'questionStr'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 60,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'questionStr'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[59])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['authorId'])):
                        self.ws.write(((2 * b) + 3), 61,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'authorId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 61,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'authorId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if (False == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['isRevised'])):
                        self.ws.write(((2 * b) + 3), 62,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'isRevised'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 62,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'isRevised'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[61]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['notes'])):
                        self.ws.write(((2 * b) + 3), 63,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'notes'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 63,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'notes'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[62])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['statusId'])):
                        self.ws.write(((2 * b) + 3), 64,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'statusId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 64,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'statusId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[63])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'questionFlagId'])):
                        self.ws.write(((2 * b) + 3), 65,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'questionFlagId'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 65,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'questionFlagId'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    # child1 correct answer

                    if ((excel_col_val[64]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['answers'][0][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 66,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answers'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 66,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answers'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[65]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2]['answers'][0][
                                'correctAnswer'])):
                        self.ws.write(((2 * b) + 3), 67,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answers'][0]['correctAnswer'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 67,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answers'][0]['correctAnswer'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    # child3 options

                    if ((excel_col_val[66]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 68,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 68,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[67]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][0]['choice'])):
                        self.ws.write(((2 * b) + 3), 69,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][0]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 69,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][0]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[68]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][1]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 70,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][1]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 70,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][1]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[69]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][1]['choice'])):
                        self.ws.write(((2 * b) + 3), 71,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][1]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 71,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][1]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[70]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][2]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 72,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][2]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 72,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][2]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[71]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][2]['choice'])):
                        self.ws.write(((2 * b) + 3), 73,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][2]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 73,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][2]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[72]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][3]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 74,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][3]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 74,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][3]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[73]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                'answerChoices'][3]['choice'])):
                        self.ws.write(((2 * b) + 3), 75,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][3]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 75,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][2][
                                          'answerChoices'][3]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)


















                    # Child question 4 validation

                    if ((int(excel_col_val[74])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'questionType'])):
                        self.ws.write(((2 * b) + 3), 76,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'questionType'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 76,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'questionType'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[75])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'difficultyLevel'])):
                        self.ws.write(((2 * b) + 3), 77,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'difficultyLevel'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 77,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'difficultyLevel'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[76])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'subCategoryId'])):
                        self.ws.write(((2 * b) + 3), 78,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'subCategoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 78,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'subCategoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[77])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'categoryId'])):
                        self.ws.write(((2 * b) + 3), 79,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'categoryId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 79,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'categoryId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[78]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'htmlString'])):
                        self.ws.write(((2 * b) + 3), 80,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 80,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[79]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'questionStr'])):
                        self.ws.write(((2 * b) + 3), 81,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'questionStr'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 81,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'questionStr'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[80])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'authorId'])):
                        self.ws.write(((2 * b) + 3), 82,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'authorId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 82,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'authorId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if (False == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'isRevised'])):
                        self.ws.write(((2 * b) + 3), 83,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'isRevised'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 83,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'isRevised'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[82]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3]['notes'])):
                        self.ws.write(((2 * b) + 3), 84,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'notes'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 84,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'notes'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[83])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'statusId'])):
                        self.ws.write(((2 * b) + 3), 85,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'statusId'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 85,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'statusId'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((int(excel_col_val[84])) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'questionFlagId'])):
                        self.ws.write(((2 * b) + 3), 86,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'questionFlagId'], self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 86,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'questionFlagId'], self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    # child4 correct answer

                    if ((excel_col_val[85]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3]['answers'][
                                0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 87,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answers'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 87,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answers'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[86]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3]['answers'][
                                0]['correctAnswer'])):
                        self.ws.write(((2 * b) + 3), 88,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answers'][0]['correctAnswer'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 88,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answers'][0]['correctAnswer'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    # child4 options

                    if ((excel_col_val[87]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 89,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 89,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[88]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][0]['choice'])):
                        self.ws.write(((2 * b) + 3), 90,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][0]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 90,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][0]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[89]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][1]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 91,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][1]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 91,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][1]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[90]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][1]['choice'])):
                        self.ws.write(((2 * b) + 3), 92,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][1]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 92,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][1]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[91]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][2]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 93,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][2]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 93,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][2]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[92]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][2]['choice'])):
                        self.ws.write(((2 * b) + 3), 94,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][2]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 94,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][2]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[93]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][3]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 95,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][3]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 95,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][3]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[94]) == (
                            responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                'answerChoices'][3]['choice'])):
                        self.ws.write(((2 * b) + 3), 96,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][3]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 96,
                                      responseFormGetQuestionForldResponse_1['data']['childQuestions'][3][
                                          'answerChoices'][3]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)
























                 #   self.ws.write(((2 * b) + 3), 26, 'Successful', self.style1)

                    self.wb_Result.save("exelReportRTC.xls")

                else:
                    print("Unable to create and validate question. Please Check the server ")

            # If unable to create question because of wrong json input
            elif(self.abcd['status'] == 400):

              #  createQuestionErrorMassage = {(self.abcd['error']['massage'])}


                self.ws.write(((2 * b) + 2), 1, 'Input', self.style0)


                # Printing Input value

                self.ws.write(((2 * b) + 2), 2, int((excel_col_val[0])), self.style5)
                self.ws.write(((2 * b) + 2), 3, int((excel_col_val[1])), self.style5)
                self.ws.write(((2 * b) + 2), 4, int((excel_col_val[2])), self.style5)
                self.ws.write(((2 * b) + 2), 5, int((excel_col_val[3])), self.style5)
                self.ws.write(((2 * b) + 2), 6, (excel_col_val[4]), self.style5)
                self.ws.write(((2 * b) + 2), 7, (excel_col_val[5]), self.style5)
                self.ws.write(((2 * b) + 2), 8, int((excel_col_val[6])), self.style5)
                self.ws.write(((2 * b) + 2), 9, False, self.style5)
                self.ws.write(((2 * b) + 2), 10, (excel_col_val[8]), self.style5)
                self.ws.write(((2 * b) + 2), 11, int((excel_col_val[9])), self.style5)
                self.ws.write(((2 * b) + 2), 12, int((excel_col_val[10])), self.style5)

                # child1

                self.ws.write(((2 * b) + 2), 13, int((excel_col_val[11])), self.style5)
                self.ws.write(((2 * b) + 2), 14, int((excel_col_val[12])), self.style5)
                self.ws.write(((2 * b) + 2), 15, int((excel_col_val[13])), self.style5)
                self.ws.write(((2 * b) + 2), 16, int((excel_col_val[14])), self.style5)
                self.ws.write(((2 * b) + 2), 17, (excel_col_val[15]), self.style5)
                self.ws.write(((2 * b) + 2), 18, (excel_col_val[16]), self.style5)
                self.ws.write(((2 * b) + 2), 19, int((excel_col_val[17])), self.style5)
                self.ws.write(((2 * b) + 2), 20, False, self.style5)
                self.ws.write(((2 * b) + 2), 21, (excel_col_val[19]), self.style5)
                self.ws.write(((2 * b) + 2), 22, int((excel_col_val[20])), self.style5)
                self.ws.write(((2 * b) + 2), 23, (excel_col_val[21]), self.style5)

                self.ws.write(((2 * b) + 2), 24, (excel_col_val[22]), self.style5)
                self.ws.write(((2 * b) + 2), 25, (excel_col_val[23]), self.style5)

                self.ws.write(((2 * b) + 2), 26, (excel_col_val[24]), self.style5)
                self.ws.write(((2 * b) + 2), 27, (excel_col_val[25]), self.style5)
                self.ws.write(((2 * b) + 2), 28, (excel_col_val[26]), self.style5)
                self.ws.write(((2 * b) + 2), 29, (excel_col_val[27]), self.style5)
                self.ws.write(((2 * b) + 2), 30, (excel_col_val[28]), self.style5)
                self.ws.write(((2 * b) + 2), 31, (excel_col_val[29]), self.style5)
                self.ws.write(((2 * b) + 2), 32, (excel_col_val[30]), self.style5)
                self.ws.write(((2 * b) + 2), 33, (excel_col_val[31]), self.style5)

                # child 2

                self.ws.write(((2 * b) + 2), 34, int((excel_col_val[32])), self.style5)
                self.ws.write(((2 * b) + 2), 35, int((excel_col_val[33])), self.style5)
                self.ws.write(((2 * b) + 2), 36, int((excel_col_val[34])), self.style5)
                self.ws.write(((2 * b) + 2), 37, int((excel_col_val[35])), self.style5)
                self.ws.write(((2 * b) + 2), 38, (excel_col_val[36]), self.style5)
                self.ws.write(((2 * b) + 2), 39, (excel_col_val[37]), self.style5)
                self.ws.write(((2 * b) + 2), 40, int((excel_col_val[38])), self.style5)
                self.ws.write(((2 * b) + 2), 41, False, self.style5)
                self.ws.write(((2 * b) + 2), 42, (excel_col_val[40]), self.style5)
                self.ws.write(((2 * b) + 2), 43, int((excel_col_val[41])), self.style5)
                self.ws.write(((2 * b) + 2), 44, (excel_col_val[42]), self.style5)

                self.ws.write(((2 * b) + 2), 45, (excel_col_val[43]), self.style5)
                self.ws.write(((2 * b) + 2), 46, (excel_col_val[44]), self.style5)

                self.ws.write(((2 * b) + 2), 47, (excel_col_val[45]), self.style5)
                self.ws.write(((2 * b) + 2), 48, (excel_col_val[46]), self.style5)
                self.ws.write(((2 * b) + 2), 49, (excel_col_val[47]), self.style5)
                self.ws.write(((2 * b) + 2), 50, (excel_col_val[48]), self.style5)
                self.ws.write(((2 * b) + 2), 51, (excel_col_val[49]), self.style5)
                self.ws.write(((2 * b) + 2), 52, (excel_col_val[50]), self.style5)
                self.ws.write(((2 * b) + 2), 53, (excel_col_val[51]), self.style5)
                self.ws.write(((2 * b) + 2), 54, (excel_col_val[52]), self.style5)

                # child3

                self.ws.write(((2 * b) + 2), 55, int((excel_col_val[53])), self.style5)
                self.ws.write(((2 * b) + 2), 56, int((excel_col_val[54])), self.style5)
                self.ws.write(((2 * b) + 2), 57, int((excel_col_val[55])), self.style5)
                self.ws.write(((2 * b) + 2), 58, int((excel_col_val[56])), self.style5)
                self.ws.write(((2 * b) + 2), 59, (excel_col_val[57]), self.style5)
                self.ws.write(((2 * b) + 2), 60, (excel_col_val[58]), self.style5)
                self.ws.write(((2 * b) + 2), 61, int((excel_col_val[59])), self.style5)
                self.ws.write(((2 * b) + 2), 62, False, self.style5)
                self.ws.write(((2 * b) + 2), 63, (excel_col_val[61]), self.style5)
                self.ws.write(((2 * b) + 2), 64, int((excel_col_val[62])), self.style5)
                self.ws.write(((2 * b) + 2), 65, (excel_col_val[63]), self.style5)

                self.ws.write(((2 * b) + 2), 66, (excel_col_val[64]), self.style5)
                self.ws.write(((2 * b) + 2), 67, (excel_col_val[65]), self.style5)

                self.ws.write(((2 * b) + 2), 68, (excel_col_val[66]), self.style5)
                self.ws.write(((2 * b) + 2), 69, (excel_col_val[67]), self.style5)
                self.ws.write(((2 * b) + 2), 70, (excel_col_val[68]), self.style5)
                self.ws.write(((2 * b) + 2), 71, (excel_col_val[69]), self.style5)
                self.ws.write(((2 * b) + 2), 72, (excel_col_val[70]), self.style5)
                self.ws.write(((2 * b) + 2), 73, (excel_col_val[71]), self.style5)
                self.ws.write(((2 * b) + 2), 74, (excel_col_val[72]), self.style5)
                self.ws.write(((2 * b) + 2), 75, (excel_col_val[73]), self.style5)

                # child4


                self.ws.write(((2 * b) + 2), 76, int((excel_col_val[74])), self.style5)
                self.ws.write(((2 * b) + 2), 77, int((excel_col_val[75])), self.style5)
                self.ws.write(((2 * b) + 2), 78, int((excel_col_val[76])), self.style5)
                self.ws.write(((2 * b) + 2), 79, int((excel_col_val[77])), self.style5)
                self.ws.write(((2 * b) + 2), 80, (excel_col_val[78]), self.style5)
                self.ws.write(((2 * b) + 2), 81, (excel_col_val[79]), self.style5)
                self.ws.write(((2 * b) + 2), 82, int((excel_col_val[80])), self.style5)
                self.ws.write(((2 * b) + 2), 83, False, self.style5)
                self.ws.write(((2 * b) + 2), 84, (excel_col_val[82]), self.style5)
                self.ws.write(((2 * b) + 2), 85, int((excel_col_val[83])), self.style5)
                self.ws.write(((2 * b) + 2), 86, (excel_col_val[84]), self.style5)

                self.ws.write(((2 * b) + 2), 87, (excel_col_val[85]), self.style5)
                self.ws.write(((2 * b) + 2), 88, (excel_col_val[86]), self.style5)

                self.ws.write(((2 * b) + 2), 89, (excel_col_val[87]), self.style5)
                self.ws.write(((2 * b) + 2), 90, (excel_col_val[88]), self.style5)
                self.ws.write(((2 * b) + 2), 91, (excel_col_val[89]), self.style5)
                self.ws.write(((2 * b) + 2), 92, (excel_col_val[90]), self.style5)
                self.ws.write(((2 * b) + 2), 93, (excel_col_val[91]), self.style5)
                self.ws.write(((2 * b) + 2), 94, (excel_col_val[92]), self.style5)
                self.ws.write(((2 * b) + 2), 95, (excel_col_val[93]), self.style5)
                self.ws.write(((2 * b) + 2), 96, (excel_col_val[94]), self.style5)
                self.ws.write_merge(((2 * b) + 2),((2 * b) + 2), 97,98, (excel_col_val[95]), self.style5)

                self.ws.write(((2 * b) + 3), 0, 'Pass', self.style7)

                self.ws.write(((2 * b) + 3), 1, 'Output', self.style1)


                if(excel_col_val[95] == self.abcd['error']['errorDescription']):
                    self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 97, 98, self.abcd['error']['errorDescription'], self.style1)
                else:
                    self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 97, 98, 'Error description is not matching', self.style2)
                    self.ws.write(((2 * b) + 3), 0, 'Fail', self.style2)
                    self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)


                #self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 25, 26, self.abcd['error']['errorDescription'],self.style2)
                self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 99, 100, self.abcd['error']['message'], self.style1)
                #self.ws.write(((2 * b) + 3), 26, 'Unsuccessful', self.style2)


                self.wb_Result.save("exelReportRTC.xls")





    def exelHeaderMaker(self):

        self.style0 = xlwt.easyxf('font: name Arial, color-index black, bold on')
        self.style1 = xlwt.easyxf('font: name Arial, color-index green, bold off')
        self.style2 = xlwt.easyxf('font: name Arial, color-index red, bold on')
        self.style3 = xlwt.easyxf('font: name Arial, color-index black, bold on')
        self.style5 = xlwt.easyxf('font: name Arial, color-index black, bold off')
        self.style6 = xlwt.easyxf('font: name Arial, color-index blue, bold on')
        self.style8 = xlwt.easyxf('font: name Arial, color-index brown, bold on')
        self.style7 = xlwt.easyxf('font: name Arial, color-index green, bold on')

        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('EC_Verification',cell_overwrite_ok=True)

        self.ws.write_merge(0, 0, 0, 1, 'RTC type question Creation and Validation : Rel _ 108', self.style6)
        self.ws.write_merge(0, 0, 2, 3, ('Date : ')+ self.__current_DateTime, self.style8)
        self.ws.write_merge(0, 0, 4, 5, ('No of test Cases : 2 ') , self.style3)
        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Pass '), self.style7)


        #parent question header

        self.ws.write(1, 0, 'Status', self.style0)
        self.ws.write(1, 1, 'Value Type', self.style0)
        self.ws.write(1, 2, 'Question Type', self.style0)
        self.ws.write(1, 3, 'Difficulty Level', self.style0)
        self.ws.write(1, 4, 'Sub Category Id', self.style0)
        self.ws.write(1, 5, 'Category Id', self.style0)
        self.ws.write(1, 6, 'Html String', self.style0)
        self.ws.write(1, 7, 'questionStr', self.style0)
        self.ws.write(1, 8, 'authorId', self.style0)
        self.ws.write(1, 9, 'isRevised', self.style0)
        self.ws.write(1, 10, 'notes', self.style0)
        self.ws.write(1, 11, 'statusId', self.style0)
        self.ws.write(1, 12, 'questionFlag', self.style0)

        #child1 question header


        self.ws.write(1, 13, 'Question Type', self.style0)
        self.ws.write(1, 14, 'Difficulty Level', self.style0)
        self.ws.write(1, 15, 'Sub Category Id', self.style0)
        self.ws.write(1, 16, 'Category Id', self.style0)
        self.ws.write(1, 17, 'Html String', self.style0)
        self.ws.write(1, 18, 'questionStr', self.style0)
        self.ws.write(1, 19, 'authorId', self.style0)
        self.ws.write(1, 20, 'isRevised', self.style0)
        self.ws.write(1, 21, 'notes', self.style0)
        self.ws.write(1, 22, 'statusId', self.style0)
        self.ws.write(1, 23, 'questionFlag', self.style0)

        self.ws.write(1, 24, 'htmlString', self.style0)
        self.ws.write(1, 25, 'correctAnswer', self.style0)


        self.ws.write(1, 26, 'htmlString', self.style0)
        self.ws.write(1, 27, 'choice', self.style0)

        self.ws.write(1, 28, 'htmlString', self.style0)
        self.ws.write(1, 29, 'choice', self.style0)

        self.ws.write(1, 30, 'htmlString', self.style0)
        self.ws.write(1, 31, 'choice', self.style0)

        self.ws.write(1, 32, 'htmlString', self.style0)
        self.ws.write(1, 33, 'choice', self.style0)

        # child2 question header


        self.ws.write(1, 34, 'Question Type', self.style0)
        self.ws.write(1, 35, 'Difficulty Level', self.style0)
        self.ws.write(1, 36, 'Sub Category Id', self.style0)
        self.ws.write(1, 37, 'Category Id', self.style0)
        self.ws.write(1, 38, 'Html String', self.style0)
        self.ws.write(1, 39, 'questionStr', self.style0)
        self.ws.write(1, 40, 'authorId', self.style0)
        self.ws.write(1, 41, 'isRevised', self.style0)
        self.ws.write(1, 42, 'notes', self.style0)
        self.ws.write(1, 43, 'statusId', self.style0)
        self.ws.write(1, 44, 'questionFlag', self.style0)

        self.ws.write(1, 45, 'htmlString', self.style0)
        self.ws.write(1, 46, 'correctAnswer', self.style0)

        self.ws.write(1, 47, 'htmlString', self.style0)
        self.ws.write(1, 48, 'choice', self.style0)

        self.ws.write(1, 49, 'htmlString', self.style0)
        self.ws.write(1, 50, 'choice', self.style0)

        self.ws.write(1, 51, 'htmlString', self.style0)
        self.ws.write(1, 52, 'choice', self.style0)

        self.ws.write(1, 53, 'htmlString', self.style0)
        self.ws.write(1, 54, 'choice', self.style0)

        # child3 question header


        self.ws.write(1, 55, 'Question Type', self.style0)
        self.ws.write(1, 56, 'Difficulty Level', self.style0)
        self.ws.write(1, 57, 'Sub Category Id', self.style0)
        self.ws.write(1, 58, 'Category Id', self.style0)
        self.ws.write(1, 59, 'Html String', self.style0)
        self.ws.write(1, 60, 'questionStr', self.style0)
        self.ws.write(1, 61, 'authorId', self.style0)
        self.ws.write(1, 62, 'isRevised', self.style0)
        self.ws.write(1, 63, 'notes', self.style0)
        self.ws.write(1, 64, 'statusId', self.style0)
        self.ws.write(1, 65, 'questionFlag', self.style0)

        self.ws.write(1, 66, 'htmlString', self.style0)
        self.ws.write(1, 67, 'correctAnswer', self.style0)

        self.ws.write(1, 68, 'htmlString', self.style0)
        self.ws.write(1, 69, 'choice', self.style0)

        self.ws.write(1, 70, 'htmlString', self.style0)
        self.ws.write(1, 71, 'choice', self.style0)

        self.ws.write(1, 72, 'htmlString', self.style0)
        self.ws.write(1, 73, 'choice', self.style0)

        self.ws.write(1, 74, 'htmlString', self.style0)
        self.ws.write(1, 75, 'choice', self.style0)

        # child4 question header


        self.ws.write(1, 76, 'Question Type', self.style0)
        self.ws.write(1, 77, 'Difficulty Level', self.style0)
        self.ws.write(1, 78, 'Sub Category Id', self.style0)
        self.ws.write(1, 79, 'Category Id', self.style0)
        self.ws.write(1, 80, 'Html String', self.style0)
        self.ws.write(1, 81, 'questionStr', self.style0)
        self.ws.write(1, 82, 'authorId', self.style0)
        self.ws.write(1, 83, 'isRevised', self.style0)
        self.ws.write(1, 84, 'notes', self.style0)
        self.ws.write(1, 85, 'statusId', self.style0)
        self.ws.write(1, 86, 'questionFlag', self.style0)

        self.ws.write(1, 87, 'htmlString', self.style0)
        self.ws.write(1, 88, 'correctAnswer', self.style0)

        self.ws.write(1, 89, 'htmlString', self.style0)
        self.ws.write(1, 90, 'choice', self.style0)

        self.ws.write(1, 91, 'htmlString', self.style0)
        self.ws.write(1, 92, 'choice', self.style0)

        self.ws.write(1, 93, 'htmlString', self.style0)
        self.ws.write(1, 94, 'choice', self.style0)

        self.ws.write(1, 95, 'htmlString', self.style0)
        self.ws.write(1, 96, 'choice', self.style0)



        #error massage

        self.ws.write_merge(1, 1, 97, 98, 'Error Description( If any ) ', self.style0)
        self.ws.write_merge(1, 1, 99, 100, 'Error Massage( If any ) ', self.style0)
        #self.ws.write(1, 26, 'Status', self.style0)

        self.wb_Result.save("exelReportRTC.xls")





ob = CreateMcqQuestion()