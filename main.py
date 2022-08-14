import pdfplumber
import datetime
import xlsxwriter
import os

sections = ('skills', 'education', 'work', 'projects', 'achievements', 'certifications', 'courses', 'responsibility')

# To extract pdf from their current directory
def get_all_resume(dir_path):
    if not os.path.exists(dir_path): # Just to check whether the path exists or not
        return []
    files = os.listdir(dir_path)
    # print(files)
    pdf_files = []
    for file in files:
        if os.path.isfile(os.path.join(dir_path, file)) and file.endswith('.pdf'):
            pdf_files.append(file)

    return pdf_files # All pdf/anything that is file are now extracted from the current directory

def console_out():

    print("CHOOSE:")
    print("0 : Links")
    for i in range(len(sections)):
        print(i + 1, ':', sections[i])
    print("INPUT CHOICES SEPERATED BY SPACES!")
    choices = list(map(int, input().split()))
    return choices

def doc_parser(resume):

    # pdf_content = ""
    words = list()
    with pdfplumber.open(resume) as pdf:
        pages = pdf.pages
        for page in pages:
            # pdf_content += page.extract_text()
            for e in page.extract_words():
                # print(e)
                words.append(e['text'])

    index = dict()
    for section in sections:
        for word in words:
            if section.casefold() == word.casefold() or section.casefold()+':' == word.casefold():
                index[section] = words.index(word)

    

    index_list = sorted(index.items(), key=lambda x: x[1])
    
    
    index = dict(index_list)
    result = dict()
    headings = list(index.keys())
    print(headings)
    list_index = list(index.values())
    
    l = len(headings)-1
    for i in range(l):
        content = ' '.join(words[list_index[i]+1: list_index[i+1]])
        result[headings[i]] = content
    
    result[headings[-1]] = ' '.join(words[list_index[-1]+1: ])

    links = list()
    for word in words:
        if word.__contains__('@') or word.__contains__('.com') or word.__contains__('http'):
            links.append(word)

    # for k in result:
    #     print(k, result[k], sep=':')

    return result, links


def resume_reader(result, links, worksheet, choices, row):

    col = 0 
    for choice in choices:
        if choice < 0 or choice > len(sections):
            output = "NULL"
        elif choice == 0:
            output = ' , '.join(links)
        else:
            key = sections[choice-1]
            output = result.get(key, "NULL") #?#

        # Write text to Excel File
        worksheet.write(row, col, output)
        col += 1
        # print(key, ":", output)


if __name__ == "__main__":

    # GETTING ALL THE PDF FILES FROM THE GIVEN DIRECTORY
    dir_path = "ResumeC"

    # RETURN ALL .PDF FILES
    resumes = get_all_resume(dir_path)

    # GETTING ALL THE SECTIONS TO BE EXTRACTED FROM THE RESUME
    # choices = console_out()
    choices = [i for i in range(0, len(sections)+1)]

    # CREATE A NEW EXCEL WORKSHEET
    ct = datetime.datetime.now().timestamp()
    session_filename = "resume_" + str(ct) + ".xlsx"

    workbook = xlsxwriter.Workbook(session_filename)
    worksheet = workbook.add_worksheet()
    # worksheet.set_column('A:A', 20)

    # WRITING THE HEADINGS IN THE FIRST ROW HERE
    #choice should be in range 0 to len(sections)
    heading = ""
    for i in range(len(choices)):
        if choices[i] < 0 or choices[i] > len(sections):
            heading = "NULL"
        else:
            heading = sections[choices[i] - 1] if choices[i] else "links"
        worksheet.write(0, i, heading)

    # ITERATING OVER EACH RESUME, AND WRITING THE RESULT TO EXCEL FILE
    row = 1
    for resume in resumes:
        result, links = doc_parser(os.path.join(dir_path, resume))
        resume_reader(result, links, worksheet, choices, row)
        row += 1

    # CLOSE AND SAVE THE FINAL EXCEL SHEET
    workbook.close()


