import sys
import datetime
import os

from statswriter import StatsWriter
import dataloader
import render

# SVG template to use
svg_template = "scorecard2.svg"

def main(args):
    if len(args) not in [3, 4]:
        sys.stderr.write("Usage: %s <data file> <visit number> [year]\n" % args[0])
        sys.exit(1)
    code_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)))
    project_root = os.path.join(code_dir, os.path.pardir)
    resource_dir = os.path.join(project_root, "resources")
    output_dir = os.path.join(project_root, "output")

    filename = args[1]
    visit = args[2]
    if len(args) == 5:
        calc_year = int(args[3])
    else:
        now = datetime.datetime.now()
        calc_year = now.year

    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
        os.mkdir(os.path.join(output_dir, "scorecard"))
        os.mkdir(os.path.join(output_dir, "noscorecard"))
    stats_writer = StatsWriter(open("stats.csv", "w"))

    context = {}
    context["school_map"] = dataloader.load_schooltypes(os.path.join(resource_dir, "school_type.xls"))
    context["menu"] = dataloader.load_menu(os.path.join(resource_dir, "menu.xls"))

    template_xml = open(os.path.join(resource_dir, svg_template)).read().decode("utf-8")

    # load all visit data
    all_data = dataloader.load_data(filename, visit, calc_year, context)

    all_visit_data = all_data["current_visit"]
    for i, school in enumerate(all_visit_data):
        print "Processing school: %s" % school.school_number
        stats_data = []
        stats_data.append(school.school_number)
        stats_data.append(school.visit)
        stats_data.append(school.stock_score)
        stats_data.append(school.meal_delivery_score)
        stats_data.append(school.hygiene_score)
        stats_data.append(school.staff_score)
        stats_data.append(school.total_score)

        school_xml = template_xml
        indicators = render.calculate_indicators(all_data, school) 
        stats_data.append(indicators["rank_total"])
        stats_data.append(indicators["year_rank_total"])
        school_xml = render.render_scorecard(indicators, school_xml)

        output_path = os.path.join(output_dir, "%s" % ("scorecard" if context["school_map"][school.school_number]["score_card"] == 1 else "noscorecard"))
        output_file = "%d.svg" % school.school_number
        f = open(os.path.join(output_path, output_file), "w")
        f.write(school_xml.encode("utf-8"))
        f.close()

        stats_writer.write_stats(stats_data)

if __name__ == "__main__":
    main(sys.argv)
