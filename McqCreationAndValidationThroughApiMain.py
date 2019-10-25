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
xlsfileloc = ("C:\\Users\\con661\\Desktop\\questionjsonstring1xls.xls")
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
                       "answers": [{"htmlString": (excel_col_val[12]), "correctAnswer": (excel_col_val[13])}],
                       "answerChoices": [{"htmlString": (excel_col_val[15]), "choice": (excel_col_val[16])},
                                         {"htmlString": (excel_col_val[17]), "choice": (excel_col_val[18])},
                                         {"htmlString": (excel_col_val[19]), "choice": (excel_col_val[20])},
                                         {"htmlString": (excel_col_val[21]), "choice": (excel_col_val[22])}]})
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
                    self.ws.write(((2 * b) + 2), 12, (excel_col_val[10]), self.style5)

                    self.ws.write(((2 * b) + 2), 13, (excel_col_val[12]), self.style5)
                    self.ws.write(((2 * b) + 2), 14, (excel_col_val[13]), self.style5)

                    self.ws.write(((2 * b) + 2), 15, (excel_col_val[15]), self.style5)
                    self.ws.write(((2 * b) + 2), 16, (excel_col_val[16]), self.style5)
                    self.ws.write(((2 * b) + 2), 17, (excel_col_val[17]), self.style5)
                    self.ws.write(((2 * b) + 2), 18, (excel_col_val[18]), self.style5)
                    self.ws.write(((2 * b) + 2), 19, (excel_col_val[19]), self.style5)
                    self.ws.write(((2 * b) + 2), 20, (excel_col_val[20]), self.style5)
                    self.ws.write(((2 * b) + 2), 21, (excel_col_val[21]), self.style5)
                    self.ws.write(((2 * b) + 2), 22, (excel_col_val[22]), self.style5)

                    # Printing Output value

                    self.ws.write(((2 * b) + 3), 0, 'Pass', self.style7)

                    self.ws.write(((2 * b) + 3), 1, 'Output', self.style1)

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

                    if ((excel_col_val[12]) == (
                    responseFormGetQuestionForldResponse_1['data']['answers'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 13,
                                      responseFormGetQuestionForldResponse_1['data']['answers'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 13,
                                      responseFormGetQuestionForldResponse_1['data']['answers'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[13]) == (
                    responseFormGetQuestionForldResponse_1['data']['answers'][0]['correctAnswer'])):
                        self.ws.write(((2 * b) + 3), 14,
                                      responseFormGetQuestionForldResponse_1['data']['answers'][0]['correctAnswer'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 14,
                                      responseFormGetQuestionForldResponse_1['data']['answers'][0]['correctAnswer'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[15]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][0]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 15,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][0]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 15,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][0]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[16]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][0]['choice'])):
                        self.ws.write(((2 * b) + 3), 16,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][0]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 16,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][0]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[17]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][1]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 17,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][1]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 17,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][1]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[18]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][1]['choice'])):
                        self.ws.write(((2 * b) + 3), 18,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][1]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 18,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][1]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[19]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][2]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 19,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][2]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 19,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][2]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[20]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][2]['choice'])):
                        self.ws.write(((2 * b) + 3), 20,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][2]['choice'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 20,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][2]['choice'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[21]) == (
                    responseFormGetQuestionForldResponse_1['data']['answerChoices'][3]['htmlString'])):
                        self.ws.write(((2 * b) + 3), 21,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][3]['htmlString'],
                                      self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 21,
                                      responseFormGetQuestionForldResponse_1['data']['answerChoices'][3]['htmlString'],
                                      self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                    if ((excel_col_val[22]) == (responseFormGetQuestionForldResponse_1['data']['answerChoices'][3]['choice'])):
                        self.ws.write(((2 * b) + 3), 22,responseFormGetQuestionForldResponse_1['data']['answerChoices'][3]['choice'],self.style1)
                    else:
                        self.ws.write(((2 * b) + 3), 22,responseFormGetQuestionForldResponse_1['data']['answerChoices'][3]['choice'],self.style2)
                        self.ws.write(((2 * b) + 3), 0, 'Fail', self.style7)
                        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)

                 #   self.ws.write(((2 * b) + 3), 26, 'Successful', self.style1)

                    self.wb_Result.save("exelReport2.xls")

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
                self.ws.write(((2 * b) + 2), 12, (excel_col_val[10]), self.style5)

                self.ws.write(((2 * b) + 2), 13, (excel_col_val[12]), self.style5)
                self.ws.write(((2 * b) + 2), 14, (excel_col_val[13]), self.style5)

                self.ws.write(((2 * b) + 2), 15, (excel_col_val[15]), self.style5)
                self.ws.write(((2 * b) + 2), 16, (excel_col_val[16]), self.style5)
                self.ws.write(((2 * b) + 2), 17, (excel_col_val[17]), self.style5)
                self.ws.write(((2 * b) + 2), 18, (excel_col_val[18]), self.style5)
                self.ws.write(((2 * b) + 2), 19, (excel_col_val[19]), self.style5)
                self.ws.write(((2 * b) + 2), 20, (excel_col_val[20]), self.style5)
                self.ws.write(((2 * b) + 2), 21, (excel_col_val[21]), self.style5)
                self.ws.write(((2 * b) + 2), 22, (excel_col_val[22]), self.style5)
                self.ws.write_merge(((2 * b) + 2),((2 * b) + 2), 23,24, (excel_col_val[23]), self.style5)

                self.ws.write(((2 * b) + 3), 0, 'Pass', self.style7)

                self.ws.write(((2 * b) + 3), 1, 'Output', self.style1)


                if(excel_col_val[23] == self.abcd['error']['errorDescription']):
                    self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 23, 24, self.abcd['error']['errorDescription'],self.style1)
                else:
                    self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 23, 24, 'Error description is not matching',self.style2)
                    self.ws.write(((2 * b) + 3), 0, 'Fail', self.style2)
                    self.ws.write_merge(0, 0, 6, 7, ('Overall status : Fail '), self.style2)


                #self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 25, 26, self.abcd['error']['errorDescription'],self.style2)
                self.ws.write_merge(((2 * b) + 3), ((2 * b) + 3), 25, 26, self.abcd['error']['message'],self.style1)
                #self.ws.write(((2 * b) + 3), 26, 'Unsuccessful', self.style2)


                self.wb_Result.save("exelReport2.xls")





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

        self.ws.write_merge(0, 0, 0, 1, 'Mcq type question Creation and Validation : Rel _ 108', self.style6)
        self.ws.write_merge(0, 0, 2, 3, ('Date : ')+ self.__current_DateTime, self.style8)
        self.ws.write_merge(0, 0, 4, 5, ('No of test Cases : 2 ') , self.style3)
        self.ws.write_merge(0, 0, 6, 7, ('Overall status : Pass '), self.style7)

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

      # self.ws.write(1, 13, 'answers', self.style0)
        self.ws.write(1, 13, 'htmlString', self.style0)
        self.ws.write(1, 14, 'correctAnswer', self.style0)

       # self.ws.write(1, 16, 'answerChoices', self.style0)

        self.ws.write(1, 15, 'htmlString', self.style0)
        self.ws.write(1, 16, 'choice', self.style0)

        self.ws.write(1, 17, 'htmlString', self.style0)
        self.ws.write(1, 18, 'choice', self.style0)

        self.ws.write(1, 19, 'htmlString', self.style0)
        self.ws.write(1, 20, 'choice', self.style0)

        self.ws.write(1, 21, 'htmlString', self.style0)
        self.ws.write(1, 22, 'choice', self.style0)
        self.ws.write_merge(1, 1, 23, 24, 'Error Description( If any ) ', self.style0)
        self.ws.write_merge(1, 1, 25, 26, 'Error Massage( If any ) ', self.style0)
        #self.ws.write(1, 26, 'Status', self.style0)

        self.wb_Result.save("exelReport2.xls")





ob = CreateMcqQuestion()