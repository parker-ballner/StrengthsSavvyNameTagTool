import sys
import os.path
from pandas import read_excel
from pathlib import Path
from pdfkit import from_file, configuration

columns = ['First Name', 'Last Name', 'Theme 1', 
            'Theme 2', 'Theme 3', 'Theme 4', 'Theme 5']
strengthTypes = {
  'executing': ['Achiever', 'Arranger', 'Belief', 'Consistency', 'Deliberative', 'Discipline', 'Focus', 'Responsibility', 'Restorative'],
  'influencing': ['Activator', 'Command', 'Communication', 'Competition', 'Maximizer', 'Self-Assurance', 'Significance', 'Woo'],
  'relationshipBuilding': ['Adaptability', 'Connectedness', 'Developer', 'Empathy', 'Harmony', 'Includer', 'Individualization', 'Positivity', 'Relator'],
  'strategicThinking': ['Analytical', 'Context', 'Futuristic', 'Ideation', 'Input', 'Intellection', 'Learner', 'Strategic']
}

templateMeta_html = './resources/TemplateMeta.html'
pageTemplate_html = './resources/PageTemplate.html'
nameTagTemplate_html = './resources/NameTagTemplate.html'

inputDir = './input/'
outputDir = './output/'

logoLoc = ''
dataLoc = ''
title = ''
outputLoc = {}

pdf_options = {
    'page-size': 'Letter',
    'orientation': 'Landscape',
    'margin-top': '0',
    'margin-right': '0',
    'margin-bottom': '0',
    'margin-left': '0',
    'no-outline': None,
    'disable-smart-shrinking': ''
}
def buildHtml():
  #aquire the data
  with open(templateMeta_html, 'r') as tm, \
      open(pageTemplate_html, 'r') as pt, \
      open(nameTagTemplate_html, 'r') as nt:
    html = tm.read()
    pageTemplate = pt.read()
    nameTagTemplate = nt.read()
  

  data = getExcelData()

  #do stuff with data
  html = html.replace("{{title}}", title)
  for entry in data:
    if '{{nametag}}' not in html:
      #add page
      html = html.replace('{{pages}}', pageTemplate + '\n{{pages}}', 1)

    #populate nametag
    nameTag = nameTagTemplate[:]
    nameTag = nameTag.replace('{{logo}}', '../' + logoLoc)
    nameTag = nameTag.replace('{{name}}', entry[0] + ' ' + entry[1])
    for strength in entry[2:]:
      #determine strength type
      strengthType = [sType for sType in strengthTypes if strength in strengthTypes[sType]][0]
      nameTag = nameTag.replace('{{strengthType}}', strengthType, 1)
      nameTag = nameTag.replace('{{strength}}', strength.upper(), 1)
    
    #add nameTag to page
    html = html.replace('{{nametag}}', nameTag, 1)

  #remove {{pages}} insert
  html = html.replace('{{pages}}', '')
  return html

def getExcelData():
  df = read_excel(dataLoc, sheet_name = 0, header = 2, usecols = columns)
  filtered_df = df.dropna()
  if not filtered_df.empty:
    return filtered_df.values.tolist()
  else:
    return None

def main(args):
  global logoLoc
  global dataLoc
  global title
  global outputLoc

  dataFileName = args[1]
  logoFileName = args[2]
  dataLoc = inputDir + dataFileName
  logoLoc = inputDir + logoFileName

  if not os.path.isfile(dataLoc) or not os.path.isfile(logoLoc):
    print('''Error: invalid filenames for excel sheet and logo.
    Excel sheet filename should be first argument,
    logo filename should be second argument.
    \nExample: `python ss_nametag_tool.py excelsheet.xslx logo.png`''')
    return

  title = dataFileName.split(' ', 1)[0]
  outputLoc = {
    'pdf': outputDir + title + '.pdf',
    'html': outputDir + title + '.html'
  }

  html_str = buildHtml()
  with open(outputLoc['html'], 'w') as html:
    html.write(html_str)
  print('Name Tag HTML created: ' + outputLoc['html'])

  with open('./wkhtmltopdf_path.txt', 'r') as wkPath:
    wkStr = wkPath.read()
    if not os.path.isfile(wkStr):
      print('''Error: invalid wkhtmltopdf path provided.
      Update the `wkhtmltopdf_path.txt` file with the correct
      location of the wkhtmltopdf binary/executable on your system''')
    else:
      config = configuration(wkhtmltopdf=wkStr)
      from_file(outputLoc['html'], outputLoc['pdf'], options=pdf_options, configuration=config)
  

#expects 2 args:
# Name of Excel Sheet File
# Name of Company Logo File
main(sys.argv)