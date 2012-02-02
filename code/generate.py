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

deviation_map = {
    "D1" : "Adhere to the menu guidelines",
    "D3" : "Make additional contributions to the set menu",
    "D4" : "Follow the preparation instructions provided by PSFA",
    "D5" : "Use correct serving and eating utensils",
    "D6" : "Reduce the amount of food waste",
    "meal_served_on_time" : "Serve meal at the correct time",
    "meal_served_efficiently" : "Serve meal efficiently so not to waste class time",
    "D8" : "Serve correct amounts of menu items",
    "E3" : "Ensure that surface used to prepare food is clean",
    "E4" : "Ensure that floor in kitchen is clean",
    "E5" : "Ensure that cleaning materials are available",
    "E6" : "Encourage volunteers to contribute to cleaning supplies",
    "E7" : "Have soap always available for hand washing",
    "E8" : "Keep a serviced fire extinguisher nearby",
    "E9" : "Store gas cylinders in an appropriate way",
    "E10" : "Have all volunteers cover their hair",
    "E11" : "Encourage volunteers to practice appropriate personal hygiene",
    "E12" : "Clean all kitchen equipment thoroughly",
    "E13" : "Clean all equipment used by learners thoroughly",
    "C1" : "Provide appropriate facilities for food storage",
    "C2" : "Keep the store room well organised",
    "C3" : "Keep all food off the ground in the storage area",
    "C4" : "Keep the storage room well secured",
    "food_in_good_condition" : "Keep all the food stuffs in good condition",
    "C9" : "Rotate the stock to avoid wastage",
    "C10" : "Notify PSFA in a timely manner when gas is needed",
    "C11" : "Notify PSFA in timely manner if there is a shortage of food",
    #"E12" : "Ration your supplies appropriately",
}


def memoize(fn):
    def _fn(self, *args, **kwargs):
        #if not hasattr(self, "_cache"):
        #    self._cache = {}
        #if fn in self._cache:
        #    return self._cache[fn]
        res = fn(self, *args, **kwargs)
        #self._cache[fn] = res
        return res
    return _fn

def invalidates_cache(fn):
    def _fn(self, *args, **kwargs):
        self._cache = {}
        return fn(self, *args, **kwargs)
    return _fn

class AvgSchoolData(object):
    def __init__(self, school_datas=None):
        self._school_datas = school_datas or []

    #@invalidates_cache
    def add_observation(self, school_data):
        self._school_datas.append(school_data)

    def __getattr__(self, key):
        if not key in self.__dict__:
            return stats.mean([getattr(s, key) for s in self._school_datas])
        else:
            return self.__dict__[key]
        

class SchoolData(object):
    field_types = {
        "yes_is_1" : [
            "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9",
            "D1", "D3", "D4", "D5", 
            "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13",
            "F1", "F3", "F4", "F5", "F6", "F7", 
        ],
        "no_is_1" : [
            "C10", "C11",
            "D6",
        ],
        "is_date" : [
            "A5",
        ],
        "is_time" : [
            "D7a", "D7b"
        ]
    }

    def __init__(self, data_row, headers, xls_datemode):
        self.data_row = data_row
        # TODO I don't like passing the datemode into the class but I'm not sure it makes sense
        # to pre-parse the data outside of the class either
        self.xls_datemode = xls_datemode
        self.headers = headers

    def __getattr__(self, key):
        if not key in self.__dict__:
            val = self._get_col_value(key)
            if key in SchoolData.field_types["yes_is_1"]:
                return self.yes_is_1(val)
            elif key in SchoolData.field_types["no_is_1"]:
                return self.no_is_1(val)
            elif key in SchoolData.field_types["is_date"]:
                return strptime(val, "%d.%m.%Y")
            elif key in SchoolData.field_types["is_time"]:
                return parse_willa_time(val, self.xls_datemode)
            return val
        else:
            return self.__dict__[key]
        
    def _get_col_value(self, prefix):
        data = dict(zip(self.headers, self.data_row))
        return data[prefix]

    def yes_is_1(self, val):
        return 1 if val.lower().strip() == "yes" else 0

    def no_is_1(self, val):
        return 1 if val.lower().strip() == "no" else 0

    @property
    def is_cooking_school(self):
        return "cooking" if self.B7 == "Cooking" else "non-cooking"

    @property
    def name(self):
        return self.A1

    @property
    def school_number(self):
        return int(self.A3)

    @property
    def visit_date(self):
        return self.A5

    @property
    @memoize
    def total_score(self):
        delivery = self.meal_delivery_score
        safety = self.hygiene_score
        stock = self.stock_score
        staff = self.staff_score

        return (delivery + safety + stock + staff) / 40.

    @property
    def meal_served_on_time(self):
        time_1030 = datetime.time(10, 30)
        return self.D7a < time_1030

    @property
    def meal_served_efficiently(self):
        diff = self.D7b - self.D7a
        return diff.minutes < 30
        
    @property
    @memoize
    def meal_delivery_score(self):
        pt_D1 = self.D1
        pt_D3 = self.D3
        pt_D4 = self.D4
        pt_D5 = self.D5
        pt_D6 = self.D6

        pt_D7a = 2 if self.meal_served_on_time else 0
        pt_D7b = 1 if self.meal_served_efficiently else 0
        val_cooking = self.is_cooking_school
        val_school_num = int(self.A3)
        val_school_type = school_map[val_school_num]
        val_total_fed = int(self.B3)
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

            item = getattr(self, item_type_field)
            qty = ingredient_qtys[item]
            amount_needed = val_total_fed * menu[item][val_day_of_visit]
            amount_served = float(hasattr(self, item_qty_field)) * qty
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
    @memoize
    def hygiene_score(self):
        pt_E2 = self.E2
        pt_E3 = self.E3
        pt_E4 = self.E4
        pt_E5 = self.E5
        pt_E6 = self.E6
        pt_E10 = self.E10 * 0.5
        pt_E11 = self.E11 * 0.5

        # 1 point if E5=NO, but E6=YES
        if pt_E5 == 0 and pt_E6 == 1:
            pt_E5 = 1

        return pt_E2 + pt_E3 + pt_E4 + pt_E5 + pt_E10 + pt_E11

    @property
    @memoize
    def stock_score(self):
        pt_C1 = self.C1
        pt_C2 = self.C2
        pt_C3 = self.C3
        pt_C4 = self.C4
        pt_C5 = self.C5 * 0.5
        pt_C6 = self.C6 * 0.5
        pt_C7 = self.C7 * 0.5
        pt_C8 = self.C8 * 0.5
        pt_C9 = self.C9
        pt_C10 = self.C10 * 0.5
        pt_C11 = self.C11 * 0.5
        #C12 - C38 # TODO need clarification

        val_cooking = self.is_cooking_school
        dtl_pilchards = self.C12b

        return sum([
            pt_C1, pt_C2, pt_C3, pt_C4, pt_C5, 
            pt_C6, pt_C7, pt_C8, pt_C9, pt_C10, 
            pt_C11
        ])

    @property
    @memoize
    def staff_score(self):
        def score_rating(x):
            x = int(x)
            if x >= 5: return 1
            if x >= 3: return 0.5
            return 0
            
        pt_F1 = self.F1
        pt_F3 = self.F3
        pt_F4 = self.F4
        pt_F5 = self.F5
        pt_F6 = self.F6
        pt_F7 = self.F7
        pt_G1 = score_rating(self.G1)
        pt_G2 = score_rating(self.G2)
        pt_G3 = score_rating(self.G3)
        pt_G4 = score_rating(self.G4)

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
        data_row = d.values()

        current_visit = d["A7"]

        school_data = SchoolData(data_row, headers, xls.datemode)
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

    visit_average = AvgSchoolData(school_datas=visit_data)
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
            avg_data[school_number] = AvgSchoolData()
        avg_data[school_number].add_observation(school_data)

    my_yearly_avg = avg_data[school.school_number]
    not_my_school_year = [s for s in avg_data.values() if s.school_number != school.school_number]
    _, _, year_rank_total = mean_percentile_rank(my_yearly_avg.total_score, [s.total_score for s in not_my_school_year])

    def calc_deviation(key):
        return getattr(visit_average, key) - getattr(school, key)
    avg_c1 = calc_deviation("C1")

    service_deviations = sorted([
        ("D1", calc_deviation("D1")),
        ("D3", calc_deviation("D3")),
        ("D4", calc_deviation("D4")),
        ("D5", calc_deviation("D5")),
        ("D6", calc_deviation("D6")),
        ("meal_served_on_time", calc_deviation("meal_served_on_time")),
        ("meal_served_efficiently", calc_deviation("meal_served_efficiently")),
    ], key=lambda x: x[1]) 
    print "Don't forget to fix the service_deviations"

    safety_deviations = sorted([
        ("E3", calc_deviation("E3")),
        ("E4", calc_deviation("E4")),
        ("E5", calc_deviation("E5")),
        ("E6", calc_deviation("E6")),
        ("E7", calc_deviation("E7")),
        ("E8", calc_deviation("E8")),
        ("E9", calc_deviation("E9")),
        ("E10", calc_deviation("E10")),
        ("E11", calc_deviation("E11")),
        ("E12", calc_deviation("E12")),
        ("E13", calc_deviation("E13")),
    ], key=lambda x: x[1]) 
    print deviation_map[service_deviations[-1][0]]
    print deviation_map[safety_deviations[-1][0]]

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
        "service_tip" : deviation_map[service_deviations[-1][0]],
        "safety_tip" : deviation_map[service_deviations[-1][0]],
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

    # each school is either primary or secondary
    load_schooltypes("../resources/school_type.xls")
    # load the menus for primary and secondary coooking and non-cooking schools
    load_menu("../resources/menu.xls")

    template_xml = open("../resources/scorecard.svg").read().decode("utf-8")

    # load all visit data
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
