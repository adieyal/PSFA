import math
import csv
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
from specialint import SpecialInt
import constants
from statswriter import StatsWriter
from schooldata import AvgSchoolData, SchoolData

re_qnum = re.compile("^(Timestamp|Verification Sch Number|[A-Z]\d+[a-z]?.).*")
def extract_header(header):
    match = re_qnum.match(header)
    if match:
        h = match.group(1)
        if h.endswith("."):
            h = h[0:-1]
        return h
    #print header, " IS NONE!!!"
    return None

def load_data(filename, visit, calc_year, calc_month, context):
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
        # Only collect results from this year and from january until the current month
        if school_data.visit_date.year != calc_year: continue
        if school_data.visit_date.month > calc_month: continue
        if not school_data.school_number in school_map: continue

        if current_visit == visit:
            data["current_visit"].append(school_data)
        if school_data.visit_date.month == calc_month:
            data["current_month"].append(school_data)
        data[current_visit].append(school_data)
        data["current_year"].append(school_data)
        data[school_data.school_number].append(school_data)
    return data

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

def remove_disclaimer(xml):
    n = xmlutils.get_el_by_id(xml, "flowRoot", "disclaimer")
    n.parentNode.removeChild(n)

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
            return "%sst" % school_rank
        elif last_digit == "2":
            return "%snd" % school_rank
        elif last_digit == "3":
            return "%srd" % school_rank
        else:
            return "%sth" % school_rank
                
    stats_data = []
    stats_data.append(school.school_number)
    stats_data.append(school.visit)
    stats_data.append(school.stock_score)
    stats_data.append(school.meal_delivery_score)
    stats_data.append(school.hygiene_score)
    stats_data.append(school.staff_score)
    stats_data.append(school.total_score)
    stats_writer.write_stats(stats_data)
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

    def calc_deviation(key, norm_value=1):
        deviation = getattr(visit_average, key) - getattr(school, key)
        return deviation / norm_value

    service_deviations = sorted([
        ("D1", calc_deviation("D1")),
        ("D3", calc_deviation("D3")),
        ("D4", calc_deviation("D4")),
        ("D5", calc_deviation("D5")),
        ("D6", calc_deviation("D6")),
        ("meal_served_on_time", calc_deviation("ind_meal_served_on_time"), 2),
        ("meal_served_efficiently", calc_deviation("ind_meal_served_efficiently")),
        ("food_deviations", calc_deviation("ind_food_deviations"), 2),
    ], key=lambda x: x[1]) 

    safety_deviations = sorted([
        ("E3", calc_deviation("E3")),
        ("E4", calc_deviation("E4")),
        ("E5", calc_deviation("E5")),
        ("E6", calc_deviation("E6")),
        ("E7", calc_deviation("E7")),
        ("E8", calc_deviation("E8")),
        ("E9", calc_deviation("E9")),
        ("E10", calc_deviation("E10"), 0.5),
        ("E11", calc_deviation("E11"), 0.5),
        ("E12", calc_deviation("E12")),
        ("E13", calc_deviation("E13")),
    ], key=lambda x: x[1]) 

    stock_deviations = sorted([
        ("C1", calc_deviation("C3")),
        ("C2", calc_deviation("C4")),
        ("C3", calc_deviation("C5")),
        ("C4", calc_deviation("C6")),
        ("food_in_good_condition", calc_deviation("food_in_good_condition"), 2),
        ("C9", calc_deviation("C9")),
        ("C10", calc_deviation("C10"), 0.5),
        ("C11", calc_deviation("C11"), 0.5),
        ("staples_in_stock", calc_deviation("ind_staples_in_stock"), 2),
    ], key=lambda x: x[1]) 

    staff_deviations = sorted([
        ("F1", calc_deviation("F1")),
        ("F3", calc_deviation("F3")),
        ("F4", calc_deviation("F4")),
        ("F5", calc_deviation("F5")),
        ("F6", calc_deviation("F6")),
        ("F7", calc_deviation("F7")),
        ("G1", calc_deviation("ind_G1")),
        ("G2", calc_deviation("ind_G2")),
        ("G3", calc_deviation("ind_G3")),
        ("G4", calc_deviation("ind_G4")),
    ], key=lambda x: x[1]) 

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
        "service_tip" : constants.deviation_map[service_deviations[-1][0]],
        "safety_tip" : constants.deviation_map[safety_deviations[-1][0]],
        "stock_tip" : constants.deviation_map[stock_deviations[-1][0]],
        "staff_tip" : constants.deviation_map[staff_deviations[-1][0]],
        "num_schools" : str(len(avg_data) + 1),
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

    if not school.voluteers_interviewed: 
        remove_disclaimer(xml)
    template_xml = xml.toxml()
    return template_xml

def main(args):
    if len(args) not in [3, 5]:
        sys.stderr.write("Usage: %s <data file> <visit number> [year] [month]\n" % args[0])
        sys.exit(1)
    code_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)))
    project_root = os.path.join(code_dir, os.path.pardir)
    resource_dir = os.path.join(project_root, "resources")
    output_dir = os.path.join(project_root, "output")

    filename = args[1]
    visit = args[2]
    if len(args) == 5:
        calc_year = int(args[3])
        calc_month = int(args[4])
    else:
        now = datetime.datetime.now()
        calc_year = now.year
        calc_month = now.month

    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
        os.mkdir(os.path.join(output_dir, "scorecard"))
        os.mkdir(os.path.join(output_dir, "noscorecard"))

    context = {}
    context["school_map"] = load_schooltypes(os.path.join(resource_dir, "school_type.xls"))
    context["menu"] = load_menu(os.path.join(resource_dir, "menu.xls"))

    template_xml = open(os.path.join(resource_dir, "scorecard2.svg")).read().decode("utf-8")

    # load all visit data
    all_data = load_data(filename, visit, calc_year, calc_month, context)

    all_visit_data = all_data["current_visit"]
    for i, school in enumerate(all_visit_data):
        print "Processing school: %s" % school.school_number

        school_xml = template_xml
        school_xml = render_scorecard(all_data, school, school_xml)

        output_path = os.path.join(output_dir, "%s" % ("scorecard" if context["school_map"][school.school_number]["score_card"] == 1 else "noscorecard"))
        output_file = "%d.svg" % school.school_number
        f = open(os.path.join(output_path, output_file), "w")
        f.write(school_xml.encode("utf-8"))
        f.close()

if __name__ == "__main__":
    stats_writer = StatsWriter(open("stats.csv", "w"))
    main(sys.argv)
