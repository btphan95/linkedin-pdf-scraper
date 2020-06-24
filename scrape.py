# importing all the required modules
import argparse
import os
import pdfplumber
import re
import pandas as pd


def main(argv):
  input_file = os.path.abspath(argv.input)
  output_file = os.path.abspath(argv.output)

  #extracting pdf
  print('extracting pdf')
  pdf = pdfplumber.open(input_file)
  pages = pdf.pages
  profiles = [""]
  i = 0
  for page in pages:
    if 'Profile Notes and Activity' not in page.extract_text():
      profiles[i] += page.extract_text() + '\n'
    else:
      i += 1
      profiles.append('')
  profiles = profiles[:-1]
  profiles = [profile.split('\n') for profile in profiles]
  profiles = [profile[:-1] for profile in profiles]

  # compile profiles dataframe
  first_names = [x[0].split(' ')[0] for x in profiles]
  last_names = [x[0].split(' ')[1] for x in profiles]
  cities = [x[1].split(',')[0] for x in profiles]
  states = [x[1].split(',')[1] if len(x[1].split(',')) > 1 else "-" for x in profiles]
  summaries = []
  for profile in profiles:
    summary = ""
    if 'Summary' in profile:
      i = profile.index('Summary') + 1  
      while (profile[i] != 'Experience') and profile[i] != 'Education' and i < len(profile):
        summary += profile[i]
        i += 1
    else:
        summary = '--'
    summaries.append(summary)
  df_profiles = pd.DataFrame({
    'First name' : first_names,
    'Last name' : last_names,
    'City' : cities,
    'State' : states,
    'Summary' : summaries
  })
  writer = pd.ExcelWriter(output_file, engine='xlsxwriter')  
  workbook = writer.book 
  worksheet=workbook.add_worksheet('Profiles')
  writer.sheets['Profiles'] = worksheet
  df_profiles.to_excel(writer, sheet_name='Profiles', startrow=1 , startcol=0)
  worksheet.write('A1', 'Profile')
  print('compiled profiles')

  # compile experience dataframe
  names = []
  positions = []
  start_positions = []
  end_positions = []
  years = []
  months = []
  descriptions = []

  for profile in profiles:
    name = profile[0]
    if 'Experience' in profile:
      i = profile.index('Experience')
      if i + 3 > len(profile):
        if re.search(r'.+ at .+', profile[i+1]):
          position = profile[i+1]
          positions.append(position)
          names.append(name)
          start_positions.append('--')
          end_positions.append('--')
          years.append('--')
          months.append('--')
          descriptions.append('--')
      #iterate through all positions
      while(i + 2 < len(profile) and 
            profile[i+1] != 'Education'):
        #initially check the formatting of each position
        #if position title present, add it
        position = '--'
        if re.search(r'.+ at .+', profile[i+1]):
          position = profile[i+1]

        #if dates present, add them

        start, end, year, month = '--', '--', '--', '--'
        try:
          dates = re.search(r'(.+)\xa0-\xa0(.+)', profile[i+2])
          if dates:
            start = dates.group(1)
            end_re = re.search(r'(.+) \(', dates.group(2))
            if end_re:
              end = end_re.group(1)
            year_re = re.search(r'.*?(\d+) (year)+.*', dates.group(2))
            if year_re:
              year = year_re.group(1)
            month_re = re.search(r'.*?(\d+) (month)+.*', dates.group(2))
            if month_re:
              month = month_re.group(1)
        except:
          print('error in finding date')

        #grab the description: this is a little hacky because I had to consider cases like having no descriptions,
        #or reaching the end of the Experience section
        try:
          des_start = i + 3
          i_des = des_start
          des_end = des_start
          
          #check if there is no description. if none, then skip to next position
          if (re.search(r'\w+ at \w+', profile[i_des]) and
            re.search(r'(.+)\xa0-\xa0(.+)', profile[i_des+1])):
            pass
          else:
            while(i_des < len(profile) and 
                  profile[i_des] != 'Education' and
                  # not re.search(r'\w+ at \w+', profile[i_des]) and
                  not re.search(r'(.+)\xa0-\xa0(.+)', profile[i_des+1])):

              # try:
              #   if (re.search(r'\w+ at \w+', profile[i_des+1]) and
              #      re.search(r'(.+)\xa0-\xa0(.+)', profile[i_des+2])):
              #      break
              # except:
              #   print('something wrong')
              i_des += 1
        except:
          # print('error in finding description')
          pass

        des_end = i_des - 1
        description = ("\n".join(profile[des_start:des_end + 1]).replace('•', '').replace('●','').replace('\nEducation', '') if
          (des_end + 1) - des_start > 1 else '--')
        description = description.replace('\n', '\r\n')
        names.append(name)
        positions.append(position)
        start_positions.append(start)
        end_positions.append(end)
        years.append(year)
        months.append(month)
        descriptions.append(description)
        #get to next position
        i = des_end
        
  df_experience = pd.DataFrame({
    'Name' : names,
    'Position' : positions,
    'Start' : start_positions,
    'End' : end_positions,
    'Years' : years,
    'Months' : months,
    'Descriptions' : descriptions
  })
  df_experience.set_index('Name', inplace=True)
  df_experience.to_excel(writer,sheet_name='Profiles',startrow=1 , startcol=8)
  worksheet.write('I1', 'Experience')
  print('compiled experience')

  # compile education dataframe
  names = []
  schools = []
  degrees = []
  studies = []
  starts = []
  ends = []
  school_keywords = ['University', 'Univ', 'Institu', 'College', 'School', 'Academy']
  degree_keywords = ['Bachelor', 'Master', 'Associate', 'B.S.', 'A.A.',
                      'M.S.', 'BS', 'AA', 'MS']
  school_keywords = '|'.join(['(' + x + ')+' for x in school_keywords])
  degree_keywords = '|'.join(['(' + x + ')+.*' for x in degree_keywords])

  for profile in profiles:
    name = profile[0]
    if 'Education' in profile:
      i = profile.index('Education')
          #iterate through all educations
      while(i + 1< len(profile)):
        #if school present, add it
        school = '--'
        if re.compile(r'%s' % (school_keywords), re.IGNORECASE).search(profile[i+1]):
          school = profile[i+1]
        line = '--'
        
        #grab the degree
        if i + 3 < len(profile):
          if re.compile(r'%s' % (school_keywords), re.IGNORECASE).search(profile[i+3]):
            line = profile[i+2]
            next_i = 2
            
          else:
            line = profile[i+2] + profile[i+3]
            next_i = len(profile)
        else:
          line = profile[i+2]
          next_i = 3
        degree = '--'
        if re.compile(r'%s' % (degree_keywords), re.IGNORECASE).search(line):
          degree = re.search(r'^.*?,', line).group()
          line = line.replace(degree, '')
        
        # grab area of study, and dates
        study = '--'
        start = '--'
        end = '--'
        dates = re.findall(r'\d\d\d\d', line)
        if len(dates) == 2:
          start = dates[0]
          end = dates[1]
        elif len(dates) == 1:
          end = dates[0]
        dates_ = re.search(r',*?\s*\d+.*$', line)
        if dates_:
          line = line.replace(dates_.group(), '')
        study = line
        names.append(name)
        schools.append(school)
        degrees.append(degree)
        studies.append(study)
        starts.append(start)
        ends.append(end)
        i += next_i
        if i < len(profile):
          pass
  df_education = pd.DataFrame({
    'Name' : names,
    'School' : schools,
    'Degree' : degrees,
    'Study' : studies,
    'Start' : starts,
    'End' : ends
  })
  df_education.set_index('Name', inplace=True)
  df_education.to_excel(writer,sheet_name='Profiles',startrow=1 , startcol=20)
  worksheet.write('U1', 'Education')
  print('compiled education')
  writer.save()
  print('wrote to profiles.xlsx')

if __name__ == "__main__":
  parser = argparse.ArgumentParser()
  parser.add_argument('-i', '--input', required = True, help = "input pdf file")
  parser.add_argument('-o', '--output', required = True, help = 'output xlsx file')
  args = parser.parse_args()
  if not os.path.exists(args.input):
      exit("Please specify an existing direcory using the -i parameter.")
  main(args)