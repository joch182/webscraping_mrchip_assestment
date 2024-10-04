from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
import pandas as pd

if __name__ == "__main__":
  print('Initiating...')
  base_url = "https://www.procyclingstats.com/"
  url = Request('https://www.procyclingstats.com/race/tour-de-france/2024/stage-1', headers={'User-Agent': 'Mozilla/5.0'})
  html_page = urlopen(url).read()
  soup = BeautifulSoup(html_page, 'html.parser')
  ul_headers = soup.find(class_ = "restabs")
  data=[]
  for child in ul_headers.children:
    result = dict()
    result['data-id'] = child.a['data-id']
    result['title'] = child.a.text
    data.append(result)
    
  table_number = 1
  tables_html = dict()
  for idx, item in enumerate(data):
    header_div = soup.find('div', class_ = "result-cont", attrs={'data-id': item['data-id']})
    subtabs_divs = header_div.find_all('div', class_ = "subTabs")
    data[idx]['tables'] = []
    for table_div in subtabs_divs:
      tables = table_div.find_all("table" , recursive=False)
      h3 = table_div.find_all('h3')
      for idx2, table in enumerate(tables):
        if len(h3) != 0:
          h3Text = h3[idx2].text
        else:
          h3Text = ''
        
        table_headers = table_div.find('thead').find('tr').find_all('th')
        table_headers = [th.text for th in table_headers]
        table_data = []
        table_data_tr = table_div.find('tbody').find_all('tr')
        for row in table_data_tr:
          row_data = []
          row_cells = row.find_all('td')
          for cell in row_cells:
            anchor = cell.find_all('a')
            if len(anchor) != 0:
              # Hyperlink
              row_data.append('=HYPERLINK("'+base_url+anchor[0]['href']+'","'+cell.text+'")')
            else:
              row_data.append(cell.text)
          table_data.append(row_data)
          
        data[idx]['tables'].append({
          'h3': h3Text,
          'table_number': table_number,
          # 'table_headers': table_headers,
          # 'table_data': table_data,
          'table': pd.DataFrame(table_data, columns=table_headers)
        })
        table_number = table_number + 1

  writer = pd.ExcelWriter('results.xlsx', engine='openpyxl')
  start_row = 0
  for idx, item in enumerate(data):
    for idx2, table in enumerate(item['tables']):
      title = '## Table ' + str(table['table_number']) + " ## " + item['title']
      if table['h3'] != '':
        title = title + ' ## Today ' + table['h3']
      data = {'Title': [title]}
      df_title = pd.DataFrame(data)
      df_title.to_excel(writer, sheet_name='Sheet1', header=False, index=False, startcol=0,startrow=start_row) 
      start_row = start_row + 1
      table['table'].to_excel(writer, sheet_name='Sheet1', header=True, index=False, startcol=0,startrow=start_row) 
      start_row = start_row + len(table['table'].index) + 1
      start_row = start_row + 1

  writer.close()
  print('Finished...')