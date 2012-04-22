import xlrd
from collections import OrderedDict, defaultdict
import re
from schooldata import SchoolData

def load_menu(filename):
    # load the menus for primary and secondary coooking and non-cooking schools
    xls = xlrd.open_workbook(filename)
    menus = {}
    for i in range(4):
        sheet = xls.sheets()[i]

        # skip first two columns
        headers = sheet.row_values(1)[2:]
        items = {} 
        for row_num in range(2, sheet.nrows):
            row = sheet.row_values(row_num)
            ingredient = row[0]

            datum = dict(zip(headers, row[2:]))
            items[ingredient] = datum
        menus[sheet.name] = items
    return menus
        
def load_schooltypes(filename):
    # each school is either primary or secondary
    xls = xlrd.open_workbook(filename)
    sheet = xls.sheets()[0]
    school_map = {}

    headers = sheet.row_values(0)
    
    for row_num in range(1, sheet.nrows):
        row = sheet.row_values(row_num)
        data = dict(zip(headers, row))
        school_num = int(data["schoolnumber"])
        data["school_type"] = "Primary" if int(data["primary_school"]) == 1 else "Secondary"
        data["score_card"] = int(data["score_card"])
        school_map[school_num] = data
    return school_map

re_qnum = re.compile("^(Timestamp|Verification Sch Number|[A-Z]\d+[a-z]?.).*")
def extract_header(header):
    match = re_qnum.match(header)
    if match:
        h = match.group(1)
        if h.endswith("."):
            h = h[0:-1]
        return h
    return None

def load_data(filename, visit, calc_year, context):
    xls = xlrd.open_workbook(filename)
    sheet = xls.sheets()[0]
    data = defaultdict(list, {})
    school_map = context["school_map"]

    headers = sheet.row_values(0)
    headers = [extract_header(header) for header in headers]
    
    for row_num in range(1, sheet.nrows):
        row = sheet.row_values(row_num)
        d = OrderedDict(zip(headers, row))
        data_row = d.values()

        current_visit = d["A7"]

        school_data = SchoolData(context, data_row, headers, xls.datemode)
        # Only collect results from this year
        if school_data.visit_date.year != calc_year: continue
        if not school_data.school_number in school_map: continue

        if current_visit == visit:
            data["current_visit"].append(school_data)
        data[current_visit].append(school_data)
        data["current_year"].append(school_data)
        data[school_data.school_number].append(school_data)
    return data

