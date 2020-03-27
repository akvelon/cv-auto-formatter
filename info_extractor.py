# Various helper functions to grab specific sections of a resume


class AkvelonResume(object):
    def __init__(self, nameTitle = "Name, Title", 
                 professionalSummary = "Professional Summary Here", 
                 technologySummary = "Technology Summary Here", 
                 keyTechnicalSkills = "Key Technical Skills Here", 
                 professionalExperience = "Professional Experience Here", 
                 education = "Education Here"):
        self.nameTitle = nameTitle # Will be a string
        self.proSum = professionalSummary # witll be a string
        self.techSum = technologySummary # will be a table
        self.keyTechSkills = keyTechnicalSkills # will be a table
        self.proExp = professionalExperience # long complicated string spanning several lines with certain structure
        self.edu = education # several lines of strings with certain structure
        
                
def get_nameTitle(doc):
    for para in doc.paragraphs:
        if para.text:
            nameTitle = para.text
            break
    return nameTitle

def get_proSum(doc):
    proSum_idx = None
    for idx, para in enumerate(doc.paragraphs):
        if len(para.text) > 150:
            proSum_idx = idx
            break
    return doc.paragraphs[proSum_idx].text

def get_techSum(doc):
    tbl = doc.tables[0]

    rows, cols = len(tbl.rows), len(tbl.columns)

    tbl_matrix = {}
    for row in range(rows):
        for col in range(cols):
            tbl_matrix[row, col] = tbl.cell(row, col).text
    return tbl_matrix

def get_keyTechSkills(doc):
    tables = doc.tables
    
    def bold_flag(para):
        flag = False
        flag1 = any([run.bold for run in para.runs])
        flag2 = para.style.font.bold
        if flag1 == True:
            flag = True
        if flag2 == True:
            flag = True
        return flag
    
    para_infos = []
    
    try: # first try approach where tech skills are in a table
        tbl = tables[1]
        rows, cols = len(tbl.rows), len(tbl.columns)

        for row in range(rows):
            for col in range(cols):
                try:
                    for para in [ p for p in tbl.cell(row,col).paragraphs if p.text]:
                        para_text = "".join([ run.text for run in para.runs])
                        para_bold = bold_flag(para)
                        para_infos.append((para_text, para_bold))
                except:
                    continue
    except: # If key skills are not in table but in the document itself
        for idx, p in enumerate(doc.paragraphs):
            if 'technical skil' in p.text.lower():
                tech_skills_start = idx
            if 'professional exper' in p.text.lower():
                tech_skills_end = idx
        
        for p in doc.paragraphs[tech_skills_start+1: tech_skills_end]:
            if p.text:
                para_infos.append((p.text, bold_flag(p)))
        
                
    def create_keySkill_data_structure(infos):
        keySkills = [info[0] for info in infos if info[1] == True and info[0].lower().startswith('skil') == False]
        keySkillsText = { keySkill : [] for keySkill in keySkills}

        currKeySkill = None

        # remove "skill" as first column header if its there
        infos = [info for info in infos if not info[0].lower().startswith('skil')]

        for info in infos:
            if info[1] == True:
                currKeySkill = info[0]
            elif info[1] == False:
                keySkillsText[currKeySkill].append(info[0])

        return keySkillsText
    
    return create_keySkill_data_structure(para_infos)

def get_proExp(doc):
    return None

def get_edu(doc):
    return None

def create_resume(doc):
    return AkvelonResume(
    nameTitle = get_nameTitle(doc),
    professionalSummary = get_proSum(doc),
    technologySummary = get_techSum(doc),
    keyTechnicalSkills = get_keyTechSkills(doc),
    professionalExperience = get_proExp(doc),
    education = get_edu(doc))
