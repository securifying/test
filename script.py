import zipfile
from docx2pdf import convert
import csv
import os

page_terms = {
        "sr_no":"field0",
        "name_of_student":"field1",
        "name_of_school":"field2",
        "standard":"field3",
        "rank":"field4",
        "seat_number":"field5",
        "school_code":"field6",
        "final_marks":"field7",
        "mega_final_marks":"field8",
        "merit_list_number":"fieldm",
        "award":"fielda",
    }

def cert_replace(old_file,new_file,rep):
    zin = zipfile.ZipFile (old_file, 'r')
    zout = zipfile.ZipFile (new_file, 'w')
    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if (item.filename == 'word/document.xml'):
            res = buffer.decode("utf-8")
            for r in rep:
                if r == "merit_list_number":
                    res = res.replace(page_terms[r],"{} / {}".format(rep["rank"], rep["merit_list_number"]))
                    continue                    
                res = res.replace(page_terms[r],rep[r])
            buffer = res.encode("utf-8")
        zout.writestr(item, buffer)
    zout.close()
    zin.close()

with open('data.csv', mode='r') as csv_file:
    csv_reader = csv.DictReader(csv_file)
    for row in csv_reader:
        cert_replace("one_student_doc_editable.docx","{}.docx".format(row["seat_number"]),row)
        # convert("{}.docx".format(row["seat_number"]), "{}_white.pdf".format(row["seat_number"]))
        # os.system("pdftk {}_white.pdf background background.pdf output {}.pdf".format(row["seat_number"],row["seat_number"]))

