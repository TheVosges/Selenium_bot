import pandas as pd 
import time


#ym = time.strftime("%Y-%m") #today year and month


def start(enter_path_from_user,list_from_GUI):

	message = "Success"

	try:
		def convert(Oracle_path,SN_path,SAP_path,IFS_path,Final_path):


			Oracle = pd.read_excel(Oracle_path,skiprows=14)
			SN = pd.read_excel(SN_path)
			SAP = pd.read_excel(SAP_path)

			SAP = SAP[SAP.BillT != 'F2'] #1.	Exclude all F2 from column  “R” 2.	Exclude CA (Can division) from column “AD
			SAP = SAP[SAP.Division != 'CA']

			SAP.to_clipboard()
			IFS = pd.read_excel(IFS_path)

			Final=SN[['Number', 'Description','Opened by','ERP Transaction number','ERP Transaction Number ( Credit )','ERP Transaction Number (Debit)',
			'Adjustment Value (Pretax Invoice Currency)','Invoice Currency designation']]


			writer = pd.ExcelWriter(Final_path, engine='openpyxl')
			#Oracle.to_excel(writer, sheet_name = 'Oracle', index = False)
			#SN.to_excel(writer, sheet_name = 'SN', index = False)
			#SAP.to_excel(writer, sheet_name = 'SAP', index = False)
			#IFS.to_excel(writer, sheet_name = 'IFS', index = False)

			Final.to_excel(writer, sheet_name = 'Reconciliation', index = False)


			Oracle_To_combine=Oracle[['Transaction Number','Curr','Invoice Net Amount','Type','ServiceNow number']]

			SAP_To_combine=SAP[['DocumentNo','Crcy.1','Net amount in doc currency','BillT']]
			SAP_To_combine.columns= ['Transaction Number','Curr','Invoice Net Amount','BillT']

			IFS_To_combine=IFS[['Invoice No','Invoice Type','Currency','Net Curr Amount','Service Now Number']]
			IFS_To_combine.columns=['Transaction Number','Invoice Type','Curr','Invoice Net Amount','ServiceNow Number']


			end_combine=pd.concat([Oracle_To_combine, SAP_To_combine,IFS_To_combine], axis=0)
			end_combine.to_excel(writer, sheet_name = 'Data to reconcile', index = False)
			print(end_combine)


			writer.save()


		import xlwings as xw
		import os



		#Poland='I:\\ICS\\IC\\Automation\\16. Service Now\\April-2022\\Poland\\'
		Main= enter_path_from_user 



		country_list=["1.Poland", "2.Netherlands", "3.Belgium", "4.Germany", "5.Spain", "6.France", "7.United Kingdom", "8.Italy",
			"9.Portugal", "10.Sweden", "Turkiye"]

		newname_country=["1.Poland - SAP","2.Netherlands - SAP","3.Belgium - SAP",
		"4.Germany - SAP + Oracle +IFS","5.Spain - Oracle","6.France - Oracle","7.UK - Oracle + IFS","8.Italy - Oracle +IFS" ,
		"9.Portugal - Oracle","10.Sweden - Oracle","11.Turkiye - Oracle"]

		old_email_dict=dict(zip(country_list, newname_country))

		#list_from_GUI = ['Poland']
		email_dict = {x:old_email_dict[x] for x in list_from_GUI}
		print(email_dict)









		path = os.path.normpath(Main)
		Month=path.split(os.sep)[-1]
		Year=path.split(os.sep)[-2]

		Month_year=Month+"_"+Year
		#print(Month_year)




		for countryloop in list_from_GUI:
			
			country=Main+countryloop


			
			Final_path = str(country) + str('\\'+email_dict[countryloop]+"_"+Month_year+'.xlsx')
			

			Oracle_path = str(country) + str('\\Oracle.xlsx')
			SN_path = str(country) + str('\\audit.xlsx')
			SAP_path = str(country) + str('\\SAP.xlsx')
			IFS_path = str(country) + str('\\IFS.xlsx')
			convert(Oracle_path,SN_path,SAP_path,IFS_path,Final_path)

			#link='I:\\ICS\\IC\\Service Now Reconciliation\2022\\5. May\\4.Germany - SAP + Oracle +IFS - done\\New folder\\'
			link=country


			SAP_xls = xw.Book(SAP_path)
			SN_xls = xw.Book(SN_path)
			Oracle_xls = xw.Book(Oracle_path)
			IFS_xls = xw.Book(IFS_path)
			Final_xls = xw.Book(Final_path)

			Accrual_xls = xw.Book(link+'\\Accrual.xlsx')
			parameter_xls = xw.Book(link+'\\parameters.xlsx')

			# SN 
			ws1 = SAP_xls.sheets.active
			ws2 = Oracle_xls.sheets.active
			ws3 = IFS_xls.sheets.active
			ws4 = Accrual_xls.sheets.active
			ws5 = SN_xls.sheets.active
			ws6 = parameter_xls.sheets.active

			Final=Final_xls.sheets.active


			ws1.api.Copy(Before=Final.api)
			try:
				Final_xls.sheets['Sheet1'].name='SAP'
			except:
				pass
			ws2.api.Copy(Before=Final.api) #Oracle is the combine one ws2
			ws3.api.Copy(Before=Final.api)
			try:
				Final_xls.sheets['QuickReportResultInvoiceCreditR'].name='IFS'
			except:
				pass
			ws4.api.Copy(Before=Final.api)
			try:
				Final_xls.sheets['Page 1'].name='Accrual'
				Final_xls.sheets['Accrual'].range('A:A').insert()
				Final_xls.sheets['Accrual'].range('A1').value="Comment"
			except:
				pass
			ws5.api.Copy(Before=Final.api)
			try:
				Final_xls.sheets['Page 1'].name='SN'
			except:
				pass
			ws6.api.Copy(Before=Final.api)
			dummylist=['dummy','dummy (2)','dummy (3)']

			for dummy in dummylist:
				try:
					Final_xls.sheets[dummy].delete()

				except:
					pass



			Final_xls.sheets.add()
			Final_xls.sheets.active.pictures.add(link+'\\params_accrual.png', left=Final.range("A1").left, top=Final.range("A2").top,width=1, height=2)
			Final_xls.sheets.active.pictures.add(link+'\\params_audit.png', left=Final.range("A1").left, top=Final.range("A2").top,width=1, height=2)
			Final_xls.sheets.active.pictures.add(link+'\\AccrualParams.png', left=Final.range("A1").left, top=Final.range("A2").top,width=1, height=2)
			Final_xls.sheets.active.pictures.add(link+'\\AuditParams.png', width=1, height=80)
			Final_xls.sheets.active.name="Screenshot_Accrual"



			'''Final_xls.sheets.add()
			Final_xls.sheets.active.pictures.add(link+'\\params_accrual.png', left=Final.range("A1").left, top=Final.range("A2").top,width=1, height=2)
			Final_xls.sheets.active.name="Screenshot_Accrual"
			Final_xls.sheets.add()
			Final_xls.sheets.active.pictures.add(link+'\\params_audit.png', left=Final.range("A1").left, top=Final.range("A2").top,width=1, height=2)
			Final_xls.sheets.active.name="Screenshot_audit"
			Final_xls.sheets.add()
			Final_xls.sheets.active.pictures.add(link+'\\AccrualParams.png', left=Final.range("A1").left, top=Final.range("A2").top,width=1, height=2)
			Final_xls.sheets.active.name="AccrualParams.png"
			Final_xls.sheets.add()
			Final_xls.sheets.active.pictures.add(link+'\\AuditParams.png', width=1, height=80)
			Final_xls.sheets.active.name="AuditParams.png"'''



			Final_xls.save()
			
			Accrual_xls.close()
			Oracle_xls.close()
			SAP_xls.close()
			IFS_xls.close()
			SN_xls.close()
			parameter_xls.close()
			Final_xls.close()
			os.rename(Final_path,str(country) + str('\\'+email_dict[countryloop]+"_"+Month_year+'.xlsx'))
			#os.rename(link, Main+email_dict[countryloop])
			print(countryloop+" is done")

	except Exception as e:
		message = e
		print(e)

	return message


#start('I:\\ICS\\IC\\Service Now Reconciliation\\2022\\test_july\\',["1.Poland"])
#SAP_xls.app.quit()
#Oracle_xls.app.quit()
#IFS_xls.app.quit()
#ws2.api.Copy(Before=ws3.api)

#Final = xw.Book()
#Final.sheets.add()
