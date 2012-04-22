import math
import stats
import constants
from xml.dom import minidom
import processutils
from processutils import xmlutils
from schooldata import AvgSchoolData, SchoolData

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
                
    visit_data = all_data["current_visit"]

    visit_average = AvgSchoolData(school_datas=visit_data)
    not_my_school = [s for s in visit_data if s != school]
    not_my_total_scores = [s.total_score for s in not_my_school]
    avg_delivery, perc_delivery, _ = mean_percentile_rank(school.meal_delivery_score, [s.meal_delivery_score for s in not_my_school])
    avg_safety, perc_safety, _ = mean_percentile_rank(school.hygiene_score, [s.hygiene_score for s in not_my_school])
    avg_stock, perc_stock, _ = mean_percentile_rank(school.stock_score, [s.stock_score for s in not_my_school])
    avg_staff, perc_staff, _ = mean_percentile_rank(school.staff_score, [s.staff_score for s in not_my_school])
    avg_total, perc_total, rank_total = mean_percentile_rank(school.total_score, not_my_total_scores)

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

