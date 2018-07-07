from gtmetrix import *
from xlwt import Workbook
import xlwt
#Workbook is created
wb = Workbook()

#Assigning Max Column Width for XL file
max_length_url = [0]
max_length_pagespeed_issues = [0]
max_length_yslow_issues = [0]
tracking_rows_xl = [0,0,0,0,0]

def extract_data(url_name,location_name,location_id,email_id,api_key):
  """Here we are creating API data dictionary which consist of all information which we have to insert into XL regarding every URL"""
  gt = GTmetrixInterface(email_id,api_key)
  pagespeed_issue_object = interface.IdentifyingPageSpeedIssues(email_id,api_key)
  yslow_issue_object = interface.IdentifyingYslowIssues(email_id,api_key)
  my_test = gt.start_test(url_name,location_id)
  try:
    api_data = my_test.fetch_results()
  except:
    print('Error in fetching API data')
  api_data['pagespeed_issues'] = pagespeed_issue_object.fetch_results(api_data['pagespeed_url'])
  api_data['yslow_issues'] = yslow_issue_object.fetch_results(api_data['yslow_url'])
  api_data['url'] = url_name
  api_data['location_name'] = location_name
  inserting_data_into_xl(**api_data)
  
def init_xl_file():
  #add_sheet is used to create sheet.
  sheet1 = wb.add_sheet('Dallas')
  sheet2 = wb.add_sheet('London')
  sheet3 = wb.add_sheet('Mumbai')
  sheet4 = wb.add_sheet('Sydney')
  sheet5 = wb.add_sheet('Sao Paulo')
  #Adding Column
  #Sheet1 for Dallas
  sheet1.write(0,0,'URL');sheet1.write(0,1,'Page Speed Score');sheet1.write(0,2,'Yslow Score');sheet1.write(0,3,'Fully Loaded Time');sheet1.write(0,4,'Total Page Size');sheet1.write(0,5,'Request');sheet1.write(0,6,'PageSpeed Issues');sheet1.write(0,7,'Yslow Issues')
  #Sheet2 for London
  sheet2.write(0,0,'URL');sheet2.write(0,1,'Page Speed Score');sheet2.write(0,2,'Yslow Score');sheet2.write(0,3,'Fully Loaded Time');sheet2.write(0,4,'Total Page Size');sheet2.write(0,5,'Request');sheet2.write(0,6,'PageSpeed Issues');sheet2.write(0,7,'Yslow Issues')
  #Sheet for Mumbai
  sheet3.write(0,0,'URL');sheet3.write(0,1,'Page Speed Score');sheet3.write(0,2,'Yslow Score');sheet3.write(0,3,'Fully Loaded Time');sheet3.write(0,4,'Total Page Size');sheet3.write(0,5,'Request');sheet3.write(0,6,'PageSpeed Issues');sheet3.write(0,7,'Yslow Issues')
  #Sheet for Sydney
  sheet4.write(0,0,'URL');sheet4.write(0,1,'Page Speed Score');sheet4.write(0,2,'Yslow Score');sheet4.write(0,3,'Fully Loaded Time');sheet4.write(0,4,'Total Page Size');sheet4.write(0,5,'Request');sheet4.write(0,6,'PageSpeed Issues');sheet4.write(0,7,'Yslow Issues')
  #Sheet for Sao Paulo
  sheet5.write(0,0,'URL');sheet5.write(0,1,'Page Speed Score');sheet5.write(0,2,'Yslow Score');sheet5.write(0,3,'Fully Loaded Time');sheet5.write(0,4,'Total Page Size');sheet5.write(0,5,'Request');sheet5.write(0,6,'PageSpeed Issues');sheet5.write(0,7,'Yslow Issues')
  #Setting Column Width which will not change
  sheet1.col(1).width = len('Page Speed Score')*256;sheet1.col(2).width = len('Yslow Score')*256;sheet1.col(3).width = len('Fully Loaded Time')*256;sheet1.col(4).width = len('Total Page Size')*256;sheet1.col(5).width = (len('Requests') + 1)*256
  sheet2.col(1).width = len('Page Speed Score')*256;sheet2.col(2).width = len('Yslow Score')*256;sheet2.col(3).width = len('Fully Loaded Time')*256;sheet2.col(4).width = len('Total Page Size')*256;sheet2.col(5).width = (len('Requests') + 1)*256
  sheet3.col(1).width = len('Page Speed Score')*256;sheet3.col(2).width = len('Yslow Score')*256;sheet3.col(3).width = len('Fully Loaded Time')*256;sheet3.col(4).width = len('Total Page Size')*256;sheet3.col(5).width = (len('Requests') + 1)*256
  sheet4.col(1).width = len('Page Speed Score')*256;sheet4.col(2).width = len('Yslow Score')*256;sheet4.col(3).width = len('Fully Loaded Time')*256;sheet4.col(4).width = len('Total Page Size')*256;sheet4.col(5).width = (len('Requests') +1 )*256
  sheet5.col(1).width = len('Page Speed Score')*256;sheet5.col(2).width = len('Yslow Score')*256;sheet5.col(3).width = len('Fully Loaded Time')*256;sheet5.col(4).width = len('Total Page Size')*256;sheet5.col(5).width = (len('Requests') + 1)*256

def inserting_data_into_xl(**api_data):
  sheet1 = wb.get_sheet(0);sheet2 = wb.get_sheet(1);sheet3 = wb.get_sheet(2);sheet4 = wb.get_sheet(3);sheet5 = wb.get_sheet(4)
  #Setting Column Width which vary with input
  if(len(api_data['url']) > max(max_length_url)):
    max_length_url.append(len(api_data['url']))
    sheet1.col(0).width = (max(max_length_url))*256;sheet2.col(0).width = (max(max_length_url))*256;sheet3.col(0).width = (max(max_length_url))*256;sheet4.col(0).width = (max(max_length_url))*256;sheet5.col(0).width = (max(max_length_url))*256

  if(len(api_data['pagespeed_issues']) > max(max_length_pagespeed_issues)):
    if((len(api_data['pagespeed_issues'])) > 256):
      max_length_pagespeed_issues.append(255)
    elif((len(api_data['pagespeed_issues'])) < 256):
      max_length_pagespeed_issues.append(len(api_data['pagespeed_issues']))
    sheet1.col(6).width = (max(max_length_pagespeed_issues))*256;sheet2.col(6).width = (max(max_length_pagespeed_issues))*256;sheet3.col(6).width = (max(max_length_pagespeed_issues))*256;sheet4.col(6).width = (max(max_length_pagespeed_issues))*256;sheet5.col(6).width = (max(max_length_pagespeed_issues))*256;

  if(len(api_data['yslow_issues']) > max(max_length_yslow_issues)):
    if((len(api_data['yslow_issues'])) > 256):
      max_length_yslow_issues.append(255)
    elif((len(api_data['yslow_issues'])) < 256):
      max_length_yslow_issues.append(len(api_data['yslow_issues']))
    sheet1.col(7).width = (max(max_length_yslow_issues))*256;sheet2.col(7).width = (max(max_length_yslow_issues))*256;sheet3.col(7).width = (max(max_length_yslow_issues))*256;sheet4.col(7).width = (max(max_length_yslow_issues))*256;sheet5.col(7).width = (max(max_length_yslow_issues))*256;

  #Inserting Data Into Sheetl for Dallas Location
  if(api_data['location_name'] == sheet1.name):
   tracking_rows_xl[0] += 1
   # Specifying style
   style = xlwt.easyxf('align: horiz left')
   #Grading for Pagespeed Score
   if(int(api_data['pagespeed_score']) >= 90):
     pagespeed_grade = 'A'
   elif((int(api_data['pagespeed_score']) >= 80) & (int(api_data['pagespeed_score']) < 90)):
     pagespeed_grade = 'B'
   elif((int(api_data['pagespeed_score']) >= 70) & (int(api_data['pagespeed_score']) < 80)):
     pagespeed_grade = 'C'
   elif((int(api_data['pagespeed_score']) >= 60) & (int(api_data['pagespeed_score']) < 70)):
     pagespeed_grade = 'D'

   #Grrading For Yslow Score
   if(int(api_data['yslow_score']) >= 90):
     yslow_grade = 'A'
   elif((int(api_data['yslow_score']) >= 80) & (int(api_data['yslow_score']) < 90)):
     yslow_grade = 'B'
   elif((int(api_data['yslow_score']) >= 70) & (int(api_data['yslow_score']) < 80)):
     yslow_grade = 'C'
   elif((int(api_data['yslow_score']) >= 60) & (int(api_data['yslow_score']) < 70)):
     yslow_grade = 'D'

   total_page_size = round(api_data['total_page_size']/2**20,2)
   fully_loaded_time = round(api_data['fully_loaded_time']/1000,2)
   #Inserting Data
   
   sheet1.write(tracking_rows_xl[0],0,api_data['url'],style)
   sheet1.write(tracking_rows_xl[0],1,(pagespeed_grade +'('+ str(api_data['pagespeed_score']) + '%' + ')'),style)
   sheet1.write(tracking_rows_xl[0],2,(yslow_grade +'('+ str(api_data['yslow_score']) + '%' + ')'),style)
   sheet1.write(tracking_rows_xl[0],3,str(fully_loaded_time)+'s',style)
   sheet1.write(tracking_rows_xl[0],4,str(total_page_size)+'MB',style)
   sheet1.write(tracking_rows_xl[0],5,api_data['requests'],style)
   sheet1.write(tracking_rows_xl[0],6,api_data['pagespeed_issues'],style)
   sheet1.write(tracking_rows_xl[0],7,api_data['yslow_issues'],style)

  #Inserting Into Sheet2 for London Location
  elif(api_data['location_name'] == sheet2.name):
    tracking_rows_xl[1] += 1
    # Specifying style
    style = xlwt.easyxf('align: horiz left')
    #Grading for Pagespeed Score
    if(int(api_data['pagespeed_score']) >= 90):
      pagespeed_grade = 'A'
    elif((int(api_data['pagespeed_score']) >= 80) & (int(api_data['pagespeed_score']) < 90)):
      pagespeed_grade = 'B'
    elif((int(api_data['pagespeed_score']) >= 70) & (int(api_data['pagespeed_score']) < 80)):
      pagespeed_grade = 'C'
    elif((int(api_data['pagespeed_score']) >= 60) & (int(api_data['pagespeed_score']) < 70)):
      pagespeed_grade = 'D'

    #Grrading For Yslow Score
    if(int(api_data['yslow_score']) >= 90):
      yslow_grade = 'A'
    elif((int(api_data['yslow_score']) >= 80) & (int(api_data['yslow_score']) < 90)):
      yslow_grade = 'B'
    elif((int(api_data['yslow_score']) >= 70) & (int(api_data['yslow_score']) < 80)):
      yslow_grade = 'C'
    elif((int(api_data['yslow_score']) >= 60) & (int(api_data['yslow_score']) < 70)):
      yslow_grade = 'D'

    total_page_size = round(api_data['total_page_size']/2**20,2)
    fully_loaded_time = round(api_data['fully_loaded_time']/1000,2)

    #Inserting Data
    sheet2.write(tracking_rows_xl[1],0,api_data['url'],style)
    sheet2.write(tracking_rows_xl[1],1,(pagespeed_grade +'('+ str(api_data['pagespeed_score']) + '%' + ')'),style)
    sheet2.write(tracking_rows_xl[1],2,(yslow_grade +'('+ str(api_data['yslow_score']) + '%' + ')'),style)
    sheet2.write(tracking_rows_xl[1],3,str(fully_loaded_time)+'s',style)
    sheet2.write(tracking_rows_xl[1],4,str(total_page_size)+'MB',style)
    sheet2.write(tracking_rows_xl[1],5,api_data['requests'],style)
    sheet2.write(tracking_rows_xl[1],6,api_data['pagespeed_issues'],style)
    sheet2.write(tracking_rows_xl[1],7,api_data['yslow_issues'],style)


  #Inserting Into Sheet3 for Mumbai Location
  if(api_data['location_name'] == sheet3.name):
    tracking_rows_xl[2] += 1
    # Specifying style
    style = xlwt.easyxf('align: horiz left')
    #Grading for Pagespeed Score
    if(int(api_data['pagespeed_score']) >= 90):
      pagespeed_grade = 'A'
    elif((int(api_data['pagespeed_score']) >= 80) & (int(api_data['pagespeed_score']) < 90)):
      pagespeed_grade = 'B'
    elif((int(api_data['pagespeed_score']) >= 70) & (int(api_data['pagespeed_score']) < 80)):
      pagespeed_grade = 'C'
    elif((int(api_data['pagespeed_score']) >= 60) & (int(api_data['pagespeed_score']) < 70)):
      pagespeed_grade = 'D'

    #Grrading For Yslow Score
    if(int(api_data['yslow_score']) >= 90):
      yslow_grade = 'A'
    elif((int(api_data['yslow_score']) >= 80) & (int(api_data['yslow_score']) < 90)):
      yslow_grade = 'B'
    elif((int(api_data['yslow_score']) >= 70) & (int(api_data['yslow_score']) < 80)):
      yslow_grade = 'C'
    elif((int(api_data['yslow_score']) >= 60) & (int(api_data['yslow_score']) < 70)):
      yslow_grade = 'D'

    total_page_size = round(api_data['total_page_size']/2**20,2)
    fully_loaded_time = round(api_data['fully_loaded_time']/1000,2)

    #Inserting Data
    sheet3.write(tracking_rows_xl[2],0,api_data['url'],style)
    sheet3.write(tracking_rows_xl[2],1,(pagespeed_grade +'('+ str(api_data['pagespeed_score']) + '%' + ')'),style)
    sheet3.write(tracking_rows_xl[2],2,(yslow_grade +'('+ str(api_data['yslow_score']) + '%' + ')'),style)
    sheet3.write(tracking_rows_xl[2],3,str(fully_loaded_time)+'s',style)
    sheet3.write(tracking_rows_xl[2],4,str(total_page_size)+'MB',style)
    sheet3.write(tracking_rows_xl[2],5,api_data['requests'],style)
    sheet3.write(tracking_rows_xl[2],6,api_data['pagespeed_issues'],style)
    sheet3.write(tracking_rows_xl[2],7,api_data['yslow_issues'],style)

  #Inserting Into Sheet4 for Sydney Location
  if(api_data['location_name'] == sheet4.name):
    tracking_rows_xl[3] += 1
    # Specifying style
    style = xlwt.easyxf('align: horiz left')
    #Grading for Pagespeed Score
    if(int(api_data['pagespeed_score']) >= 90):
      pagespeed_grade = 'A'
    elif((int(api_data['pagespeed_score']) >= 80) & (int(api_data['pagespeed_score']) < 90)):
      pagespeed_grade = 'B'
    elif((int(api_data['pagespeed_score']) >= 70) & (int(api_data['pagespeed_score']) < 80)):
      pagespeed_grade = 'C'
    elif((int(api_data['pagespeed_score']) >= 60) & (int(api_data['pagespeed_score']) < 70)):
      pagespeed_grade = 'D'

    #Grrading For Yslow Score
    if(int(api_data['yslow_score']) >= 90):
      yslow_grade = 'A'
    elif((int(api_data['yslow_score']) >= 80) & (int(api_data['yslow_score']) < 90)):
      yslow_grade = 'B'
    elif((int(api_data['yslow_score']) >= 70) & (int(api_data['yslow_score']) < 80)):
      yslow_grade = 'C'
    elif((int(api_data['yslow_score']) >= 60) & (int(api_data['yslow_score']) < 70)):
      yslow_grade = 'D'

    total_page_size = round(api_data['total_page_size']/2**20,2)
    fully_loaded_time = round(api_data['fully_loaded_time']/1000,2)

    #Inserting Data
    sheet4.write(tracking_rows_xl[3],0,api_data['url'],style)
    sheet4.write(tracking_rows_xl[3],1,(pagespeed_grade +'('+ str(api_data['pagespeed_score']) + '%' + ')'),style)
    sheet4.write(tracking_rows_xl[3],2,(yslow_grade +'('+ str(api_data['yslow_score']) + '%' + ')'),style)
    sheet4.write(tracking_rows_xl[3],3,str(fully_loaded_time)+'s',style)
    sheet4.write(tracking_rows_xl[3],4,str(total_page_size)+'MB',style)
    sheet4.write(tracking_rows_xl[3],5,api_data['requests'],style)
    sheet4.write(tracking_rows_xl[3],6,api_data['pagespeed_issues'],style)
    sheet4.write(tracking_rows_xl[3],7,api_data['yslow_issues'],style)



  #Inserting Into Sheet5 for Sao Paulo Location
  if(api_data['location_name'] == sheet5.name):
    tracking_rows_xl[4] += 1
    # Specifying style
    style = xlwt.easyxf('align: horiz left')
    #Grading for Pagespeed Score
    if(int(api_data['pagespeed_score']) >= 90):
      pagespeed_grade = 'A'
    elif((int(api_data['pagespeed_score']) >= 80) & (int(api_data['pagespeed_score']) < 90)):
      pagespeed_grade = 'B'
    elif((int(api_data['pagespeed_score']) >= 70) & (int(api_data['pagespeed_score']) < 80)):
      pagespeed_grade = 'C'
    elif((int(api_data['pagespeed_score']) >= 60) & (int(api_data['pagespeed_score']) < 70)):
      pagespeed_grade = 'D'

    #Grrading For Yslow Score
    if(int(api_data['yslow_score']) >= 90):
      yslow_grade = 'A'
    elif((int(api_data['yslow_score']) >= 80) & (int(api_data['yslow_score']) < 90)):
      yslow_grade = 'B'
    elif((int(api_data['yslow_score']) >= 70) & (int(api_data['yslow_score']) < 80)):
      yslow_grade = 'C'
    elif((int(api_data['yslow_score']) >= 60) & (int(api_data['yslow_score']) < 70)):
      yslow_grade = 'D'

    total_page_size = round(api_data['total_page_size']/2**20,2)
    fully_loaded_time = round(api_data['fully_loaded_time']/1000,2)

    #Inserting Data
    sheet5.write(tracking_rows_xl[4],0,api_data['url'],style)
    sheet5.write(tracking_rows_xl[4],1,(pagespeed_grade +'('+ str(api_data['pagespeed_score']) + '%' + ')'),style)
    sheet5.write(tracking_rows_xl[4],2,(yslow_grade +'('+ str(api_data['yslow_score']) + '%' + ')'),style)
    sheet5.write(tracking_rows_xl[4],3,str(fully_loaded_time)+'s',style)
    sheet5.write(tracking_rows_xl[4],4,str(total_page_size)+'MB',style)
    sheet5.write(tracking_rows_xl[4],5,api_data['requests'],style)
    sheet5.write(tracking_rows_xl[4],6,api_data['pagespeed_issues'],style)
    sheet5.write(tracking_rows_xl[4],7,api_data['yslow_issues'],style)
  
  
  try:
    wb.save('Final_Output.xls')
  except PermissionError:
    print('Kindly close the open file and then rum it')
  


if __name__ == '__main__':
  email_id = 'anuragkrsingh02@outlook.com'
  api_key = '00f429eb6791afdeb1338d2cc27f75bc'
  url_list = ['https://www.datamintelligence.com/research-report/active-pharmaceutical-ingredients-market/',
              'https://www.datamintelligence.com/research-report/asia-pacific-adhesives-sealants-market/',
              'https://www.datamintelligence.com/research-report/aqua-feed-market/',
              'https://www.datamintelligence.com/research-report/asia-compound-feed-market/',
              'https://www.datamintelligence.com/research-report/ancient-grains-market/',
              'https://www.datamintelligence.com/research-report/alcoholic-beverages-market/']
  location_with_id = {'Dallas':'4','London':'2','Mumbai':'5','Sydney':3,'Sao Paulo':6}
  #Initializing XL File
  init_xl_file()
  for location_name,location_id in location_with_id.items():
    for i in range(6):
      extract_data(url_list[i],location_name,location_id,email_id,api_key)
  
  

