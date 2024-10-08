import pandas as pd
import docx
import io
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from docxtpl import DocxTemplate
from openpyxl import Workbook,load_workbook

#readingCHPAddressfileasaListstoreinContent&storedeachlineinresult
content=docx.Document(r'C:\Users\Radar Lab\Documents\certificate\area_code.docx')
result=[p.text for p in content.paragraphs]
wb=load_workbook(r'C:\Users\Radar Lab\Documents\certificate\week_report_2024.xlsx')

unit=input("Enter 1 for Radar , 2 for LIDAR \n")
if unit=='1':
			unittype=input("Enter  DS, DE, AS, ZC, ZM, DC, DSP, DEP \n")
			match unittype:
				case "DS":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_ds.docx")
					type = "Stalker DSR"
				case "DSP":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_ds_passed.docx")
					type = "Stalker DSR"
				case "DE":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_de.docx")
					type = "Stalker DSR"
				case "DEP":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_de_passed.docx")
					type = "Stalker DSR"
				case "AS":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_as.docx")
					type = "Stalker II SDR"
				case "ZC":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_zc.docx")
					type = "Stalker dual SL"
				case "ZM":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_zm.docx")
					type = "Stalker dual SL"
				case "DC":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_dc.docx")
					type = "Stalker dual SL"
			date = "08/12/2024"
			#date=input("Enter Unit Arrival Date in Lab\n")
			lab_number=input("Enter unit lab number for Lab\n")
			serial_number=input("Enter unit serial number \n")
			if len(serial_number) != 6:
				print("Wrong Serial Number")
				exit(0)
			chps_number=input("Enter unit chps number \n")
			if len(chps_number) != 5:
				print("Wrong CHPS Number")
				exit(0)
			address_code=input("Enter chps address code\n")
			fork=input("Enter '1' if forks are present\n")
			if fork == '1':
				fa_number=input("Enter FA number from manual\n")
				if len(fa_number) != 6:
					print("Wrong FA Number")
					exit(0)
				fb_number=input("Enter FB number from manual\n")
				if len(fb_number) != 6:
					print("Wrong FB Number")
					exit(0)
			else:
				fa_number="119214"
				fb_number="221484"
			antenna1_number=input("Enter antenna1 serial number \n")
			if unittype != "AS" and len(antenna1_number) != 6:
				print("Wrong Antenna1 Number")
				exit(0)
			antenna2_number=input("Enter antenna2 serial number \n")
			if unittype != "AS" and len(antenna2_number) != 6:
				print("Wrong Antenna2 Number")
				exit(0)
			
#ifCHPAreanumberisnotincludedinfiletest.docxfirstincludewithaddress
			if(f'CHP ({address_code})')in result:
				ind=result.index(f'CHP ({address_code})')
				address_1=result[ind+2]
				address_2=result[ind+3]
				print(address_1)
				print(address_2)
			else:
				print("Address Code not Found")
				exit(0)
			st_address=address_1
			area_address=address_2
			context={'date':date,'lab_number':lab_number,'chps_number':chps_number,'serial_number':serial_number,'fa_number':fa_number,
'fb_number':fb_number,'antenna1_number':antenna1_number,'antenna2_number':antenna2_number,'address_code':address_code,
'st_address':st_address,'area_address':area_address}	
			#Savingthefile
			doc.render(context)
			doc.save(r"C:\Users\Radar Lab\Documents\certificate\radar\AS24-{}.docx".format(lab_number))
			ws=wb['RADAR']
			wb.font=Font(size=12)
			if unittype != "AS":
				ws.append(['AS24-'+lab_number,unittype+serial_number,'CHPS'+chps_number,address_code,date," "," "," ",type," "," ","S/N "+antenna1_number,"S/N "+antenna2_number])
			else:
				ws.append(['AS24-'+lab_number,unittype+serial_number,'CHPS'+chps_number,address_code,date," "," "," ",type," "," ","N/A","N/A"])
			wb.save(r'C:\Users\Radar Lab\Documents\certificate\week_report_2024.xlsx')
			
				


elif unit=='2':
			unit_type=input("Enter:TS,TJ,UX,LP,UL \n")
			match unit_type:
				case "TS":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_ts.docx")
					type = "TruSpeed S"
				case "TJ":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_tj.docx")
					type = "TruSpeed S"
				case "UX":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_ux.docx")
					type = "20/20 Ultralyte 200 LR"
				case "LP":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_lp.docx")
					type = "PRO-LITE+"
				case "UL":
					doc=DocxTemplate(r"C:\Users\Radar Lab\Documents\certificate\template\template_ul.docx")
					type = "20/20 Ultralyte 200 LR"
			date = "08/12/2024"
			#date=input("Enter Unit Arrival Date in Lab\n")
			lab_number=input("Enter unit labnumber for Lab\n")
			serial_number=input("Enter unit serial number\n")
			if unit_type != "LP" and len(serial_number) != 6:
				print("Wrong Serial Number")
				exit(0)
			chps_number=input("Enter unit chps number\n")
			if len(chps_number) != 5:
				print("Wrong CHPS Number")
				exit(0)
			address_code=input("Enter chps address code\n")
			#ifCHPAreanumberisnotincludedinfiletest.docxfirstincludewithaddress
			if(f'CHP ({address_code})')in result:
				ind=result.index(f'CHP ({address_code})')
				address_1=result[ind+2]
				address_2=result[ind+3]
				print(address_1)
				print(address_2)
			else:
				print("Address Code not Found")
				exit(0)
			st_address=address_1
			area_address=address_2
			context={'date':date,'lab_number':lab_number,'chps_number':chps_number,'serial_number':serial_number,
'address_code':address_code,'st_address':st_address,'area_address':area_address}	
			#Savingthefile
			doc.render(context)
			doc.save(r"C:\Users\Radar Lab\Documents\certificate\lidar\ASL24-{}.docx".format(lab_number))
			ws=wb['LIDAR']
			wb.font=Font(size=12)
			#wb.add_format({'font_name':'Times New Roman', 'font_size':12})
			ws.append(['ASL24-'+lab_number,unit_type+serial_number,'CHPS'+chps_number,address_code,date," "," "," ",type])
			wb.save(r'C:\Users\Radar Lab\Documents\certificate\week_report_2024.xlsx')

else:
			print("Enter number 1 or 2")
