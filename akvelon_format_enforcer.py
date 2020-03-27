import info_extractor
import docx
import os
import datetime

from lxml import etree as ET
import copy
import sys
import shutil
import zipfile

import utils

DOCUMENTS = [item for item in os.listdir("./inputs/") if not item.endswith("gitkeep")]

def format_resume(DOCUMENT_PATH):
    
    FILE_NAME = DOCUMENT_PATH.split("/")[-1].split(".")[0]

    # create copy of template to be populated with resume info
    STATIC_TEMPLATE_ZIP_PATH = "./template.zip"
    RESUME_ZIP_PATH = "./tmp/" + FILE_NAME + ".zip"
    shutil.copyfile(STATIC_TEMPLATE_ZIP_PATH, RESUME_ZIP_PATH)

    # Extract to-be-altered template .zip file for manipulation
    XML_FOLDER = "./tmp/" + FILE_NAME + "/"
    with zipfile.ZipFile(RESUME_ZIP_PATH, 'r') as zip_ref:
        zip_ref.extractall(XML_FOLDER)
    os.remove(RESUME_ZIP_PATH)

    # Extract resume information from source
    doc = docx.Document(DOCUMENT_PATH)
    resumeObject = info_extractor.create_resume(doc)
    
    # Extract Before Document to xml folder
    def unzip_docx(DOCUMENT_PATH, RESULT_DEST = "./", PREFIX = ""):
        FILE_NAME = PREFIX + DOCUMENT_PATH.split("/")[-1].split(".")[0]
        ZIP_PATH = DOCUMENT_PATH.split(".docx")[0] + ".zip"
        # create copy of template to be populated with resume info
        shutil.copyfile(DOCUMENT_PATH, ZIP_PATH)

        XML_FOLDER2 = RESULT_DEST + FILE_NAME
        with zipfile.ZipFile(ZIP_PATH, 'r') as zip_ref:
            zip_ref.extractall(XML_FOLDER2)

        os.remove(ZIP_PATH)

        return XML_FOLDER2
    
    XML_FOLDER_B4 = unzip_docx(DOCUMENT_PATH, "./tmp/", PREFIX = "b4_")
    

    # Modify/populate XML files with appropriate info from source

    # Technical issue which needs to be done
    def register_all_namespaces(filename):
        namespaces = dict([node for _, node in ET.iterparse(filename, events=['start-ns'])])
        for ns in namespaces:
            ET.register_namespace(ns, namespaces[ns])

    # To be altered template
    xml = ET.parse(XML_FOLDER + "/word/document.xml")
    register_all_namespaces(XML_FOLDER + "/word/document.xml")
    
    # Before document
    xml_b4 = ET.parse(XML_FOLDER_B4 + "/word/document.xml")
    register_all_namespaces(XML_FOLDER_B4 + "/word/document.xml")
    
    body = xml.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")[0]
    body_b4 = xml_b4.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body")[0]
    
    tables = body.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")

    for row in tables[0].findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr")[1:]:
        tables[0].remove(row)

    # Write Name, Title and Professional Summary
    for i in body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
        if i.text:
            if i.text == 'Name, Title':
                i.text = resumeObject.nameTitle
                #print(i.text)
            if i.text.startswith("Put text"):
                i.text = resumeObject.proSum
                #print(i.text)

    # Populate Technical Summary
    row_idxs = [ key[0] for key in resumeObject.techSum.keys() ]
    col_idxs = [ key[1] for key in resumeObject.techSum.keys() ]

    rows, cols = max(row_idxs) + 1, max(col_idxs) + 1
    

    first_row = tables[0].findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr")[0]

    for i in range(rows-1):
        dupe = copy.deepcopy(first_row)
        tables[0].append(dupe)

    for row_num, table_row in enumerate(tables[0].findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr")[:]):
        for col_num, table_col in enumerate(table_row):
            for child2 in table_col.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
                for child3 in child2.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"):
                    for child4 in child3.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):  
                        child4.text = resumeObject.techSum[row_num, col_num]

    # Populate Key Technical Skills
    num_tech_skills = len(resumeObject.keyTechSkills)

    key_bullet_element = copy.deepcopy(tables[1].findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr")[0])

    
    def addBulletPoints(key_bullet_element, bullet_points):
        tc_element = key_bullet_element.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc")[0]

        bullet_point_element = tc_element[-2]

        final_element = tc_element[-1]

        tc_element.remove(final_element)

        for bullet in bullet_points:
            bullet_point_element_dupe = copy.deepcopy(bullet_point_element)
            [x for x in list(bullet_point_element_dupe.iter()) if x.text == "Bullet point"][0].text = bullet
            tc_element.append(bullet_point_element_dupe)
            
        tc_element.append(final_element)
        tc_element.remove(bullet_point_element)

        return None
        

    for key_skill, bullet_points in resumeObject.keyTechSkills.items():
        dupe = copy.deepcopy(key_bullet_element)
        [x for x in list(dupe.iter()) if x.text == "Key Skill"][0].text = key_skill
        
        num_bullet_points = len(bullet_points)
        addBulletPoints(dupe, bullet_points)
        tables[1].append(dupe)
        
    tables[1].remove(tables[1].findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr")[0])
    
    
    b4_tables = body_b4.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")
    
    ###################### Populate rest of template file ########################
    
    # Helper function
    def find_text(elem, text):
        flag = False
        for idx, t in enumerate(elem.iter()):
            if str(t.text).lower() == text.lower():
                flag = True
        return flag
              
    for i, t in enumerate(body):
        if find_text(t, "PROFESSIONAL EXPERIENCE"):
            main_idx = i
            break
                
    for i, t in enumerate(body_b4):
        if find_text(t, "PROFESSIONAL EXPERIENCE"):
            b4_idx = i
            break
    while len(body) > (main_idx + 1):
        body.remove(body[-2])

    tmp = copy.deepcopy(body[-1])
    body.remove(body[-1])
    for t in body_b4[b4_idx:-1]:
         body.append(t)
    body.append(tmp)
        
    NAMESPACE_FIX = 'w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14"'
        
    def fix_namespaces(xml):
        
        xml_str = str(ET.tostring(xml))

        less = xml_str.index("<")
        greater = xml_str.index(">")

        xml_str_front = xml_str[:less + 1]
        xml_str_back = xml_str[greater:]

        xml_str = xml_str_front + NAMESPACE_FIX + xml_str_back

        return str.encode(xml_str[2:-1])
        
    xml = ET.ElementTree(ET.fromstring(fix_namespaces(xml)))

    # Write all modifications to file
    xml.write(XML_FOLDER + "/word/document.xml", xml_declaration = True, encoding='UTF-8')


    # Repackage altered xml files to create a word document

    # zip xml folder
    shutil.make_archive("./tmp/FORMAT_" + FILE_NAME, "zip", XML_FOLDER)

    # save result as a .docx file
    shutil.copyfile("./tmp/FORMAT_" + FILE_NAME + ".zip", "./results/" + FILE_NAME + "_FORMAT.docx")

    # Delete unneeded files
    shutil.rmtree(XML_FOLDER)
    shutil.rmtree(XML_FOLDER_B4)
    os.remove("./tmp/FORMAT_" + FILE_NAME + ".zip")
    
    return None



### Create log file

now = datetime.datetime.now()
now_str1 = now.strftime("%Y-%m-%d_%H-%M")
log = open("./logs/format-log-{}.txt".format(now_str1), 'w+')

now_str2 = now.strftime("%I:%M %p %A, %B %m %Y")
mesg = "Log created: {} \n".format(now_str2)
        
log.write(mesg)
log.write("///////////////////////////////// \n")


for DOCUMENT in DOCUMENTS:
    
    DOCUMENT_PATH = utils.convert_to_docx_and_give_path(DOCUMENT)
    
    try:
        format_resume(DOCUMENT_PATH)
        print("Successfully Formatted {}".format(DOCUMENT))
        log.write("Successfully Formatted {} \n \n".format(DOCUMENT))
    except:
        print("Issue with {}".format(DOCUMENT))
        
        log.write("Issue Formatting {} \n \n".format(DOCUMENT))
        shutil.copyfile("./inputs/" + DOCUMENT, "./issues/" + DOCUMENT)
        
        # delete tmp files possibly created:
        FILE_NAME = DOCUMENT.split("/")[-1].split(".")[0]
        XML_FOLDER = "./tmp/" + FILE_NAME + "/"
        XML_FOLDER_B4 = "./tmp/" + "b4_" + FILE_NAME + "/"
        
        try:
            shutil.rmtree(XML_FOLDER)
        except:
            continue
        try:
            shutil.rmtree(XML_FOLDER_B4)
        except:
            continue
        
        
log.close()

shutil.copyfile("./logs/format-log-{}.txt".format(now_str1), "latest_log.txt")