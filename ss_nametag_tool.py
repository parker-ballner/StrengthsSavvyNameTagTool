import sys
import os.path
import os
from pandas import read_excel
from pathlib import Path
import pdfkit
import imgkit

columns = ['Domain', 'First Name', 'Last Name', 'Theme 1', 
            'Theme 2', 'Theme 3', 'Theme 4', 'Theme 5']
strengthTypes = {
  'executing': ['Achiever', 'Arranger', 'Belief', 'Consistency', 'Deliberative', 'Discipline', 'Focus', 'Responsibility', 'Restorative'],
  'influencing': ['Activator', 'Command', 'Communication', 'Competition', 'Maximizer', 'Self-Assurance', 'Significance', 'Woo'],
  'relationshipBuilding': ['Adaptability', 'Connectedness', 'Developer', 'Empathy', 'Harmony', 'Includer', 'Individualization', 'Positivity', 'Relator'],
  'strategicThinking': ['Analytical', 'Context', 'Futuristic', 'Ideation', 'Input', 'Intellection', 'Learner', 'Strategic']
}

cwd = os.getcwd()

digitalTemplate = './resources/digital/Template.html'

templateMeta_html = './resources/print/TemplateMeta.html'
pageTemplate_html = './resources/print/PageTemplate.html'
nameTagTemplate_html = './resources/print/NameTagTemplate.html'

inputDir = cwd + '/input/'
outputDir = cwd + '/output/'

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
    'disable-smart-shrinking': '',
    'enable-local-file-access': '',
}

png_options = {
  'format': 'png',
  'enable-local-file-access': '',
}

def getExcelData():
  df = read_excel(dataLoc, sheet_name = 0, header = 2, usecols = columns)
  filtered_df = df.dropna()
  if not filtered_df.empty:
    return filtered_df.values.tolist()
  else:
    return None

def printNameTags(dataFileName):
  title = dataFileName.split(' ', 1)[0]
  outputLoc = {
    'pdf': outputDir + title + '.pdf',
    'html': outputDir + title + '.html'
  }

  with open('./wkhtmltopdf_path.txt', 'r') as wkPath:
    wkStr = wkPath.read()
    if not os.path.isfile(wkStr):
      print('''Error: invalid wkhtmltopdf path provided.
      Update the `wkhtmltopdf_path.txt` file with the correct
      location of the wkhtmltopdf binary/executable on your system''')
      return
    else:
      config = pdfkit.configuration(wkhtmltopdf=wkStr)

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
    nameTag = nameTag.replace('{{logo}}', logoLoc)
    nameTag = nameTag.replace('{{name}}', entry[1] + ' ' + entry[2])
    for strength in entry[3:]:
      #determine strength type
      strengthTypeArr = [sType for sType in strengthTypes if strength in strengthTypes[sType]]
      if len(strengthTypeArr) == 0:
        raise Exception('Invalid Strength Found: ', strength)
      
      nameTag = nameTag.replace('{{strengthType}}', strengthTypeArr[0], 1)
      nameTag = nameTag.replace('{{strength}}', strength.upper(), 1)
    
    #add nameTag to page
    html = html.replace('{{nametag}}', nameTag, 1)

  #remove {{pages}} insert
  html = html.replace('{{pages}}', '')

  with open(outputLoc['html'], 'w') as htmlFile:
    htmlFile.write(html)
  print('Name Tag HTML created: ' + outputLoc['html'])
  
  pdfkit.from_file(outputLoc['html'], outputLoc['pdf'], options=pdf_options, configuration=config)

def digitalNameTags():
  with open('./wkhtmltoimage_path.txt', 'r') as wkPath:
    wkStr = wkPath.read()
    if not os.path.isfile(wkStr):
      print('''Error: invalid wkhtmltoimage path provided.
      Update the `wkhtmltoimage_path.txt` file with the correct
      location of the wkhtmltoimage binary/executable on your system''')
    else:
      config = imgkit.config(wkhtmltoimage=wkStr)

    os.mkdir(outputDir + 'digital')

    data = getExcelData()

    for entry in data:
      title = f"{entry[1]}_{entry[2]}_background"
      print(title)
      outputLoc = {
        'html': outputDir + 'digital/' + title + '.html',
        'png': outputDir + 'digital/' + title + '.png'
      }

      with open(digitalTemplate, 'r') as tm:
        html = tm.read()
        html = html.replace("{{title}}", title)
        html = html.replace('{{logo}}', logoLoc)
        html = html.replace('{{name}}', entry[1] + ' ' + entry[2])

        topStrength = [sType for sType in strengthTypes if sType[0] == entry[0].lower()]
        if len(topStrength) == 0:
            raise Exception('Invalid Strength Found: ', entry[0])
        html = html.replace("{{topstrength}}", topStrength[0])

        for strength in entry[3:]:
          #determine strength type
          strengthTypeArr = [sType for sType in strengthTypes if strength in strengthTypes[sType]]
          if len(strengthTypeArr) == 0:
            raise Exception('Invalid Strength Found: ', strength)

          html = html.replace('{{strengthType}}', strengthTypeArr[0], 1)
          html = html.replace('{{strength}}', strength.upper(), 1)
          # strengthTypeCount.update({
          #   strengthTypeArr[0]: (strengthTypeCount.get(strengthTypeArr[0]) or 0) + 1
          # })
        # topStrength = max(strengthTypeCount, key=strengthTypeCount.get)

        # with open(outputLoc['html'], 'w') as htmlFile:
        #   htmlFile.write(html)
        #   print('Name Tag HTML created: ' + outputLoc['html'])
        # imgkit.from_file(outputLoc['html'], outputLoc['png'], options=png_options, config=config)

        imgkit.from_string(html, outputLoc['png'], options=png_options, config=config)

def main(args):
  global logoLoc
  global dataLoc
  global title
  global outputLoc

  if len(args) < 3:
    print('''Error: please provide filenames for excel sheet and logo.
    \nExample: `python ss_nametag_tool.py excelsheet.xslx logo.png`''')
    return

  dataFileName = args[1]
  logoFileName = args[2]
  dataLoc = inputDir + dataFileName
  logoLoc = inputDir + logoFileName

  digitalTags = args[3] if len(args) > 3 else None

  if not os.path.isfile(dataLoc) or not os.path.isfile(logoLoc):
    print('''Error: invalid filenames for excel sheet and logo.
    Excel sheet filename should be first argument,
    logo filename should be second argument.
    \nExample: `python ss_nametag_tool.py excelsheet.xslx logo.png`''')
    return
  
  if digitalTags:
    if os.path.isdir(outputDir + 'digital'):
      print('''Error: ./output/digital directory already exists, 
      please remove previous output before running program''')
      return

    digitalNameTags()
  else: 
    printNameTags(dataFileName)
  

#expects 2 args:
# Name of Excel Sheet File
# Name of Company Logo File
main(sys.argv)