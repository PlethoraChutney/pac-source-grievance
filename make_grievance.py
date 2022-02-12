import os
import glob
import shutil
import sys
from datetime import date
from docxtpl import DocxTemplate

doc = DocxTemplate(os.path.abspath('python_insurance_griev_template.docx'))

def make_grievance(
    image_file: str, name: str, mentor: str,
    email: str, dept: str, rep: str):
    
    global doc

    if not os.path.exists(name):
        os.mkdir(name)
    shutil.move(image_file, os.path.join(name, image_file))
    os.chdir(name)

    context = {
        'today': str(date.today()),
        'gr_name': name.replace('-', ' '),
        'mentor': mentor.replace('-', ' '),
        'gr_email': email + '@ohsu.edu',
        'gr_department': dept,
        'union_rep': rep
    }

    doc.render(context)
    doc.save(f'{str(date.today())}_PS-Grievance_{name}.docx')

    os.chdir('..')


def main(rep):
    image_files = []
    for file_type in ['*.png', '*.jpg']:
        image_files.extend(glob.glob(file_type))

    for image in image_files:
        name, mentor, email, dept = image[:-4].split('_')
        make_grievance(
            image,
            name,
            mentor,
            email,
            dept,
            rep
        )

        print(f'''
Hello,

Please find attached a grievance concerning the loss of health \
insurance for {name.replace("-", " ")} in {mentor.replace("-", " ")}\'s lab. \
This loss of coverage is in violation of the CBA, including but not limited \
to Article 9.

Thank you, 
{rep.replace("-", " ")}
GRU Steward''')

if __name__ == '__main__':
    if len(sys.argv) == 1:
        print('Give your name as an argument, in quotes.')
        sys.exit(1)

    main(sys.argv[1])
    