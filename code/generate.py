import math
import sys
import itertools
import sys
from xml.dom import minidom
import datetime
import re
import os
import xlrd
from mx.DateTime import Time, strptime
from collections import OrderedDict, defaultdict

import processutils
import stats
from processutils import xmlutils

school_map = {}
menus = {}

q_map = {
    "D1" : "D1. Check the menu for the day. Was this the meal that was served in the school today?",
}

class AvgSchoolData(object):
    def __init__(self, school_number):
        self._school_number = school_number
        self._school_datas = []

    def add_observation(self, school_data):
        if school_data.school_number != self._school_number:
            raise ValueError("Excepted school number to be %s not %s" % (self._school_number, school_data.school_number))
        self._school_datas.append(school_data)

    @property
    def total_score(self):
        if len(self._school_datas) == 0:
            return 0
        vals = sum(s.total_score for s in self._school_datas)
        return vals / len(self._school_datas)

    @property
    def school_number(self):
        return self._school_number

class SchoolData(object):
    def __init__(self, data_row, headers):
        self.data_row = data_row
        self.headers = headers
        
    def _get_col_value(self, prefix):
        data = dict(zip(self.headers, self.data_row))
        return data[prefix]

    def yes_is_1(self, val):
        return 1 if val.lower().strip() == "yes" else 0

    def no_is_1(self, val):
        return 1 if val.lower().strip() == "no" else 0

    @property
    def is_cooking_school(self):
        return "cooking" if self._get_col_value("B7") == "Cooking" else "non-cooking"

    @property
    def name(self):
        return self._get_col_value("A1")

    @property
    def school_number(self):
        return int(self._get_col_value("A3"))

    @property
    def visit_date(self):
        val_date_of_visit = strptime(self._get_col_value("A5"), "%d.%m.%Y")
        return val_date_of_visit

    @property
    def total_score(self):
        delivery = self.meal_delivery_score
        safety = self.hygiene_score
        stock = self.stock_score
        staff = self.staff_score

        return (delivery + safety + stock + staff) / 40.
        
    @property
    def meal_delivery_score(self):
        pt_D1 = self.yes_is_1(self._get_col_value("D1"))
        pt_D3 = self.yes_is_1(self._get_col_value("D3"))
        pt_D4 = self.yes_is_1(self._get_col_value("D4"))
        pt_D5 = self.yes_is_1(self._get_col_value("D5"))
        pt_D6 = self.no_is_1(self._get_col_value("D6"))

        time_1030 = datetime.time(10, 30)
        pt_D7a = 2 if self._get_col_value("D7a") < time_1030 else 0
        diff = self._get_col_value("D7b") - self._get_col_value("D7a")
        pt_D7b = 1 if diff.minutes < 30 else 0
        val_cooking = self.is_cooking_school
        val_school_num = int(self._get_col_value("A3"))
        val_school_type = school_map[val_school_num]
        val_total_fed = int(self._get_col_value("B3"))
        val_date_of_visit = self.visit_date
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

            item = self._get_col_value(item_type_field)
            qty = ingredient_qtys[item]
            amount_needed = val_total_fed * menu[item][val_day_of_visit]
            amount_served = float(self._get_col_value(item_qty_field)) * qty
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

        return float(sum([
            pt_D1, pt_D3, pt_D4, pt_D5, pt_D6, pt_D7a, pt_D7b, pt_ranges
        ]))

    @property
    def hygiene_score(self):
        pt_E2 = self.yes_is_1(self._get_col_value("E2"))
        pt_E3 = self.yes_is_1(self._get_col_value("E3"))
        pt_E4 = self.yes_is_1(self._get_col_value("E4"))
        pt_E5 = self.yes_is_1(self._get_col_value("E5"))
        pt_E6 = self.yes_is_1(self._get_col_value("E6"))
        pt_E10 = self.yes_is_1(self._get_col_value("E10")) * 0.5
        pt_E11 = self.yes_is_1(self._get_col_value("E11")) * 0.5

        # 1 point if E5=NO, but E6=YES
        if pt_E5 == 0 and pt_E6 == 1:
            pt_E5 = 1

        return pt_E2 + pt_E3 + pt_E4 + pt_E5 + pt_E10 + pt_E11

    @property
    def stock_score(self):
        pt_C1 = self.yes_is_1(self._get_col_value("C1"))
        pt_C2 = self.yes_is_1(self._get_col_value("C2"))
        pt_C3 = self.yes_is_1(self._get_col_value("C3"))
        pt_C4 = self.yes_is_1(self._get_col_value("C4"))
        pt_C5 = self.yes_is_1(self._get_col_value("C5")) * 0.5
        pt_C6 = self.yes_is_1(self._get_col_value("C6")) * 0.5
        pt_C7 = self.yes_is_1(self._get_col_value("C7")) * 0.5
        pt_C8 = self.yes_is_1(self._get_col_value("C8")) * 0.5
        pt_C9 = self.yes_is_1(self._get_col_value("C9"))
        pt_C10 = self.no_is_1(self._get_col_value("C10")) * 0.5
        pt_C11 = self.no_is_1(self._get_col_value("C11")) * 0.5
        #C12 - C38 # TODO need clarification

        val_cooking = self.is_cooking_school
        dtl_pilchards = self._get_col_value("C12b")

        return sum([
            pt_C1, pt_C2, pt_C3, pt_C4, pt_C5, 
            pt_C6, pt_C7, pt_C8, pt_C9, pt_C10, 
            pt_C11
        ])

    @property
    def staff_score(self):
        def score_rating(x):
            x = int(x)
            if x >= 5: return 1
            if x >= 3: return 0.5
            return 0
            
        pt_F1 = self.yes_is_1(self._get_col_value("F1"))
        pt_F3 = self.yes_is_1(self._get_col_value("F3"))
        pt_F4 = self.yes_is_1(self._get_col_value("F4"))
        pt_F5 = self.yes_is_1(self._get_col_value("F5"))
        pt_F6 = self.yes_is_1(self._get_col_value("F6"))
        pt_F7 = self.yes_is_1(self._get_col_value("F7"))
        pt_G1 = score_rating(self._get_col_value("G1"))
        pt_G2 = score_rating(self._get_col_value("G2"))
        pt_G3 = score_rating(self._get_col_value("G3"))
        pt_G4 = score_rating(self._get_col_value("G4"))

        return sum([
            pt_F1, pt_F3, pt_F4, pt_F5, pt_F6, pt_F7, 
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
            

def load_data(filename, visit, calc_year, calc_month):
    xls = xlrd.open_workbook(filename)
    sheet = xls.sheets()[0]
    data = defaultdict(list, {})

    headers = sheet.row_values(0)
    headers = [extract_header(header) for header in headers]
    
    for row_num in range(1, sheet.nrows):
        row = sheet.row_values(row_num)
        d = OrderedDict(zip(headers, row))
        d["D7a"] = parse_willa_time(d["D7a"], xls.datemode)
        d["D7b"] = parse_willa_time(d["D7b"], xls.datemode)
        data_row = d.values()

        current_visit = d["A7"]

        school_data = SchoolData(data_row, headers)
        # Only collect results from this year and from january until the current month
        if school_data.visit_date.year != calc_year: continue
        if school_data.visit_date.month > calc_month: continue

        if current_visit == visit:
            data["current_visit"].append(school_data)
        if school_data.visit_date.month == calc_month:
            data["current_month"].append(school_data)
        data[current_visit].append(school_data)
        data["current_year"].append(school_data)
        data[school_data.school_number].append(school_data)
    return data

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

def generate_graph(xml, prefix, val, width=19.407):
    # ensure that ceil of 5 is 6
    valr = int(math.ceil(val + 0.00001))
    frac = 1. - (valr - val)
    if frac > 0:
        box_width = frac * width
        n = xmlutils.get_el_by_id(xml, "rect", "%s%d" % (prefix, valr))
        n.setAttribute("width", "%f" % box_width)
        valr += 1
        
    for i in range(valr, 11):
        n = xmlutils.get_el_by_id(xml, "rect", "%s%d" % (prefix, i))
        n.parentNode.removeChild(n)

def generate_arrow_graph(xml, element_name, val, lbound, ubound):
    n = xmlutils.get_el_by_id(xml, "g", element_name)
    span = ubound - lbound
    adjustment = val * span
    pos = lbound + adjustment
    n.setAttribute("transform", "translate(%s,0)" % pos)

def render_scorecard(all_data, school, template_xml): 

    def mean_percentile_rank(val, vals):
        """
        Expect that vals does not contain val
        """
        all_vals = [val] + vals
        mean_val = round(stats.mean(all_vals), 1)
        val_rank = stats.rank(all_vals)[0]
        perc = calc_smiley_face(val, vals)

        return mean_val, perc, val_rank

    def calc_smiley_face(val, vals):
        sorted_vals = sorted([val] + vals)
        lbound = stats.percentile(sorted_vals, 0.33333)
        ubound = stats.percentile(sorted_vals, 0.66666)

        if val >= ubound:
            perc = "green"
        elif val >= lbound:
            perc = "yellow"
        else:
            perc = "red"
        return perc

    def show_smiley_face(xml, suffix, perc):
        for colour in ["red", "yellow", "green"]:
            if perc != colour:
                n = xmlutils.get_el_by_id(xml, "g", "%s%s" % (colour, suffix))
                n.parentNode.removeChild(n)

    def stringify_rank(school_rank):
        s_school_rank = str(school_rank)
        last_digit = s_school_rank[-1]
        if last_digit == "1":
            return "%sst" % last_digit
        elif last_digit == "2":
            return "%snd" % last_digit
        elif last_digit == "3":
            return "%srd" % last_digit
        else:
            return "%sth" % last_digit
                
    visit_data = all_data["current_visit"]
    not_my_school = [s for s in visit_data if s != school]
    not_my_total_scores = [s.total_score for s in not_my_school]
    avg_delivery, perc_delivery, _ = mean_percentile_rank(school.meal_delivery_score, [s.meal_delivery_score for s in not_my_school])
    avg_safety, perc_safety, _ = mean_percentile_rank(school.hygiene_score, [s.hygiene_score for s in not_my_school])
    avg_stock, perc_stock, _ = mean_percentile_rank(school.stock_score, [s.stock_score for s in not_my_school])
    avg_staff, perc_staff, _ = mean_percentile_rank(school.staff_score, [s.staff_score for s in not_my_school])
    avg_total, perc_total, rank_total = mean_percentile_rank(school.total_score, not_my_total_scores)

    # this month
    # TODO - this code currently assumes that there is at most 1 visit per month
    not_my_school_month = [s for s in all_data["current_month"] if s != school]
    _, _, month_rank_total = mean_percentile_rank(school.total_score, [s.total_score for s in not_my_school_month])

    
    avg_data = {}
    for school_data in all_data["current_year"]:
        school_number = school_data.school_number
        if not school_number in avg_data:
            avg_data[school_number] = AvgSchoolData(school_number)
        avg_data[school_number].add_observation(school_data)

    my_yearly_avg = avg_data[school.school_number]
    not_my_school_year = [s for s in avg_data.values() if s.school_number != school.school_number]
    _, _, year_rank_total = mean_percentile_rank(my_yearly_avg.total_score, [s.total_score for s in not_my_school_year])
        
    context = {
        "s_name" : school.name,
        "v_date" : school.visit_date.strftime("%d %B %Y"),
        "pt_ydelivery" : str(school.meal_delivery_score),
        "pt_odelivery" : str(avg_delivery),
        "pt_ysafety" : str(school.hygiene_score),
        "pt_osafety" : str(avg_safety),
        "pt_ystock" : str(school.stock_score),
        "pt_ostock" : str(avg_stock),
        "pt_ystaff" : str(school.staff_score),
        "pt_ostaff" : str(avg_staff),
        "pt_total" : str(round(school.total_score * 100, 2)),
        "pt_avg_total" : str(round(avg_total * 100, 2)),
        "rank_total" : str(int(rank_total)),
        "month_rank_total" : str(int(month_rank_total)),
        "month_rank_total_str" : stringify_rank(int(month_rank_total)),
        "year_rank_total" : str(int(year_rank_total)),
        "year_rank_total_str" : stringify_rank(int(year_rank_total)),
    }
    template_xml = processutils.process_svg_template(context, template_xml)
    xml = minidom.parseString(template_xml.encode("utf-8"))
    generate_graph(xml, "del_y", school.meal_delivery_score)
    generate_graph(xml, "del_o", avg_delivery)
    generate_graph(xml, "saf_y", school.hygiene_score)
    generate_graph(xml, "saf_o", avg_safety)
    generate_graph(xml, "stock_y", school.stock_score)
    generate_graph(xml, "stock_o", avg_stock)
    generate_graph(xml, "staff_y", school.staff_score)
    generate_graph(xml, "staff_o", avg_staff)

    generate_arrow_graph(xml, "pt_total", school.total_score, -91.6875, 29.063798)
    generate_arrow_graph(xml, "pt_avg_total", avg_total, -36.8125, 83.9375)

    show_smiley_face(xml, "_delivery", perc_delivery)
    show_smiley_face(xml, "_safety", perc_safety)
    show_smiley_face(xml, "_stock", perc_stock)
    show_smiley_face(xml, "_staff", perc_staff)
    template_xml = xml.toxml()
    return template_xml

def main(args):
    if len(args) not in [3, 5]:
        sys.stderr.write("Usage: %s <data file> <visit number> [year] [month]\n" % args[0])
        sys.exit(1)

    filename = args[1]
    visit = args[2]
    if len(args) == 5:
        calc_year = int(args[3])
        calc_month = int(args[4])
    else:
        now = datetime.datetime.now()
        calc_year = now.year
        calc_month = now.month

    load_schooltypes("../resources/school_type.xls")
    load_menu("../resources/menu.xls")
    template_xml = open("../resources/scorecard.svg").read().decode("utf-8")

    all_data = load_data(filename, visit, calc_year, calc_month)
    all_visit_data = all_data["current_visit"]
    for i, school in enumerate(all_visit_data):
        school_xml = template_xml
        school_xml = render_scorecard(all_data, school, school_xml)

        f = open("output/%d.svg" % i, "w")
        f.write(school_xml.encode("utf-8"))
        f.close()

if __name__ == "__main__":
    main(sys.argv)
