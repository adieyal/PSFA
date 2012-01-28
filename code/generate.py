import sys
import datetime
import re
import os
import tablib
import xlrd
from mx.DateTime import Time, strptime
from ordereddict import OrderedDict

school_map = {}
menus = {}

q_map = {
    "D1" : "D1. Check the menu for the day. Was this the meal that was served in the school today?",
}

class ScorecardDataset(object):
    def __init__(self, *args, **kwargs):
        self.dataset = tablib.Dataset(*args, **kwargs)
        self.headers = self.dataset.headers
        
        self.dataset.append_col(self.meal_delivery_points, header="Meal Delivery Points")
        self.dataset.append_col(self.hygiene_points, header="Hygiene Points")
        self.dataset.append_col(self.stock_points, header="Stock Points")
        self.dataset.append_col(self.staff_points, header="Staff Points")
        print self.dataset["Meal Delivery Points"]

    def _get_col_value(self, prefix, row):
        data = dict(zip(self.headers, row))
        return data[prefix]

    def yes_is_1(self, val):
        return 1 if val.lower().strip() == "yes" else 0

    def no_is_1(self, val):
        return 1 if val.lower().strip() == "no" else 0

    def is_cooking_school(self, row):
        return "cooking" if self._get_col_value("B7", row) == "Cooking" else "non-cooking"
        
    def meal_delivery_points(self, row):
        pt_D1 = self.yes_is_1(self._get_col_value("D1", row))
        pt_D3 = self.yes_is_1(self._get_col_value("D3", row))
        pt_D4 = self.yes_is_1(self._get_col_value("D4", row))
        pt_D5 = self.yes_is_1(self._get_col_value("D5", row))
        pt_D6 = self.no_is_1(self._get_col_value("D6", row))

        time_1030 = Time(10, 30)
        pt_D7a = 2 if self._get_col_value("D7a", row) < time_1030 else 0
        diff = self._get_col_value("D7b", row) - self._get_col_value("D7a", row)
        pt_D7b = 1 if diff.minutes < 30 else 0
        val_cooking = self.is_cooking_school(row)
        val_school_num = int(self._get_col_value("A3", row))
        val_school_type = school_map[val_school_num]
        val_total_fed = int(self._get_col_value("B3", row))
        val_date_of_visit = strptime(self._get_col_value("A5", row), "%d.%m.%Y")
        dow = {
            0 : "Monday",
            1 : "Tuesday",
            2 : "Wednesday",
            3 : "Thursday",
            4 : "Friday",
            5 : "Saturday",
            6 : "Sunday",
        }
        val_day_of_visit = dow[val_date_of_visit.weekday()]
        menu_name = "%s %s schools" % (val_school_type, val_cooking)
        menu = menus[menu_name]

        ingredient_qtys = {
            "Pilchards" : 1.88,
            "Rice" : 10,
            "Savoury Mince" : 10,
            "Curry Mince" : 10,
            "Samp" : 10,
            "Sugar beans" : 5,
            "Brown lentils" : 500,
            "Breyani spice" : 1,
            "Cabbage" : 3,
            "Carrots" : 5,
            "Butternut" : 5,
            "Fruits" : 10,
            "Peanut butter" : 5,
            "Jam" : 3.75,
            "Baked beans" : 3,
            "Jungle oats" : 1,
            "Salt" : 1,
            "Oil" : 4,
            "Sugar" : 10,
        }

        ranges = [] 
        for field_idx in range(8, 14):
            item_type_field = "D%da" % field_idx
            item_qty_field = "D%db" % field_idx

            item = self._get_col_value(item_type_field, row)
            qty = ingredient_qtys[item]
            amount_needed = val_total_fed * menu[item][val_day_of_visit]
            amount_served = float(self._get_col_value(item_qty_field, row)) * qty
            if amount_needed == 0:
                range_perc = 0
            else:
                range_perc = abs(amount_served - amount_needed) / amount_needed
            ranges.append(range_perc)
            
        ranges_below_10 = [range_perc for range_perc in ranges if range_perc < 0.1]
        if len(ranges_below_10) == len(ranges):
            pt_ranges = 2
        elif len(ranges_below_10) >= 3:
            pt_ranges = 1
        else:
            pt_ranges = 0
        return pt_D1 + pt_D3 + pt_D4 + pt_D5 + pt_D6 + pt_D7a + pt_D7b + pt_ranges

    def hygiene_points(self, row):
        pt_E2 = self.yes_is_1(self._get_col_value("E2", row))
        pt_E3 = self.yes_is_1(self._get_col_value("E3", row))
        pt_E4 = self.yes_is_1(self._get_col_value("E4", row))
        pt_E5 = self.yes_is_1(self._get_col_value("E5", row))
        pt_E6 = self.yes_is_1(self._get_col_value("E6", row))
        pt_E10 = self.yes_is_1(self._get_col_value("E10", row)) * 0.5
        pt_E11 = self.yes_is_1(self._get_col_value("E11", row)) * 0.5

        # 1 point if E5=NO, but E6=YES
        if pt_E5 == 0 and pt_E6 == 1:
            pt_E5 = 1

        return pt_E2 + pt_E3 + pt_E4 + pt_E5 + pt_E10 + pt_E11

    def stock_points(self, row):
        pt_C1 = self.yes_is_1(self._get_col_value("C1", row))
        pt_C2 = self.yes_is_1(self._get_col_value("C2", row))
        pt_C3 = self.yes_is_1(self._get_col_value("C3", row))
        pt_C4 = self.yes_is_1(self._get_col_value("C4", row))
        pt_C5 = self.yes_is_1(self._get_col_value("C5", row)) * 0.5
        pt_C6 = self.yes_is_1(self._get_col_value("C6", row)) * 0.5
        pt_C7 = self.yes_is_1(self._get_col_value("C7", row)) * 0.5
        pt_C8 = self.yes_is_1(self._get_col_value("C8", row)) * 0.5
        pt_C9 = self.yes_is_1(self._get_col_value("C9", row))
        pt_C10 = self.no_is_1(self._get_col_value("C10", row)) * 0.5
        pt_C11 = self.no_is_1(self._get_col_value("C11", row)) * 0.5
        #C12 - C38 # TODO need clarification

        val_cooking = self.is_cooking_school(row)
        dtl_pilchards = self._get_col_value("C12b", row)

        return pt_C1 + pt_C2 + pt_C3 + pt_C4 + pt_C5 + pt_C6 + pt_C7 + pt_C8 + pt_C9 + pt_C10 + pt_C11

    def staff_points(self, row):
        def score_rating(x):
            x = int(x)
            if x >= 5: return 1
            if x >= 3: return 0.5
            return 0
            
        pt_F1 = self.yes_is_1(self._get_col_value("F1", row))
        pt_F3 = self.yes_is_1(self._get_col_value("F3", row))
        pt_F4 = self.yes_is_1(self._get_col_value("F4", row))
        pt_F5 = self.yes_is_1(self._get_col_value("F5", row))
        pt_F6 = self.yes_is_1(self._get_col_value("F6", row))
        pt_F7 = self.yes_is_1(self._get_col_value("F7", row)) * 0.5
        pt_F8 = self.yes_is_1(self._get_col_value("F8", row)) * 0.5
        pt_G1 = score_rating(self._get_col_value("G1", row))
        pt_G2 = score_rating(self._get_col_value("G2", row))
        pt_G3 = score_rating(self._get_col_value("G3", row))
        pt_G4 = score_rating(self._get_col_value("G4", row))

        return sum([
            pt_F1, pt_F3, pt_F4, pt_F5, pt_F6, pt_F7, pt_F8, 
            pt_G1, pt_G2, pt_G3, pt_G4
        ])

re_qnum = re.compile("^(Timestamp|Verification Sch Number|[A-Z]\d+[a-z]?.).*")
def extract_header(header):
    match = re_qnum.match(header)
    if match:
        h = match.group(1)
        if h.endswith("."):
            h = h[0:-1]
        return h
    print header, " IS NONE!!!"
    return None

def parse_willa_time(val, datemode):
    if val == "9:00 or earlier":
        return Time(9, 0, 0)
    elif val == "12:00 or later":
        return Time(12, 0, 0)
    date_tuple = xlrd.xldate_as_tuple(val, datemode)[3:]
    return Time(*date_tuple)
            

def load_data(filename):
    xls = xlrd.open_workbook(filename)
    sheet = xls.sheets()[0]
    data = []

    headers = sheet.row_values(0)
    headers = [extract_header(header) for header in headers]
    
    for row_num in range(1, sheet.nrows):
        row = sheet.row_values(row_num)
        d = OrderedDict(zip(headers, row))
        d["D7a"] = parse_willa_time(d["D7a"], xls.datemode)
        d["D7b"] = parse_willa_time(d["D7b"], xls.datemode)
        data.append(d.values())
    return ScorecardDataset(*data, headers=headers)

def load_schooltypes(filename):
    xls = xlrd.open_workbook(filename)
    sheet = xls.sheets()[0]

    headers = sheet.row_values(0)
    
    for row_num in range(1, sheet.nrows):
        school_num, school_type = sheet.row_values(row_num)
        school_map[int(school_num)] = "Primary" if int(school_type) == 1 else "Secondary"

def load_menu(filename):
    xls = xlrd.open_workbook(filename)
    for i in range(4):
        sheet = xls.sheets()[i]

        headers = sheet.row_values(1)
        items = {} 
        for row_num in range(2, sheet.nrows):
            row = sheet.row_values(row_num)
            datum = dict(zip(headers, row))
            items[datum["Item"]] = datum
        menus[sheet.name] = items
        

def main(args):
    filename = args[1]
    load_schooltypes("../resources/school_type.xls")
    load_menu("../resources/menu.xls")
    dataset = load_data(filename)

    

if __name__ == "__main__":
    main(sys.argv)
