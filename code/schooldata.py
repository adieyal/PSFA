import sys
import datetime
import stats
import xlrd
from mx.DateTime import Time, strptime
from specialint import SpecialInt
import constants

def parse_willa_time(val, datemode):
    if val == "9:00 or earlier":
        return Time(9, 0, 0)
    elif val == "12:00 or later":
        return Time(12, 0, 0)
    date_tuple = xlrd.xldate_as_tuple(val, datemode)[3:]
    return Time(*date_tuple)

class AvgSchoolData(object):
    """
    Class that takes a list of SchoolDatas and computes a mean for any indicator requested
    """
    def __init__(self, school_datas=None):
        self._school_datas = school_datas or []

    def add_observation(self, school_data):
        self._school_datas.append(school_data)

    def __getattr__(self, key):
        if not key in self.__dict__:
            return stats.mean([getattr(s, key) for s in self._school_datas])
        else:
            return self.__dict__[key]

class SchoolData(object):
    """
    A wrapper around a submission
    """

    field_types = {
        "yes_is_1" : [
            "C1", "C2", "C4", "C5", "C6", "C7", "C8", "C9",
            "D1", "D3", "D4", "D5", 
            "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", "E13",
            "F1", "F3", "F4", "F5", "F6", "F7", "F8",
        ],
        "no_is_1" : [
            "C3", "C10", "C11",
            "D6",
        ],
        "is_date" : [
            "A5",
        ],
        "is_time" : [
            "D7a", "D7b"
        ],
        "is_int" : [
            "B3", "B6",
            "G1", "G2", "G3", "G4",
        ],
        "is_float" : [
            "D8b", "D9b", "D10b", "D11b", "D12b", "D13b",
        ],
    }

    def __init__(self, context, data_row, headers, xls_datemode):
        self.data_row = data_row
        # TODO I don't like passing the datemode into the class but I'm not sure it makes sense
        # to pre-parse the data outside of the class either
        self.xls_datemode = xls_datemode
        self.headers = headers
        self._data = dict(zip(self.headers, self.data_row))
        self.school_map = context["school_map"]
        self.menus = context["menu"]

    def __getattr__(self, key):
        try:
            val = self._get_col_value(key)
            if key in SchoolData.field_types["yes_is_1"]:
                return SpecialInt(self.yes_is_1(val))
            elif key in SchoolData.field_types["no_is_1"]:
                return SpecialInt(self.no_is_1(val))
            elif key in SchoolData.field_types["is_date"]:
                return strptime(val.strip(), "%d.%m.%Y")
            elif key in SchoolData.field_types["is_time"]:
                return self.parse_time(val)
            elif key in SchoolData.field_types["is_int"]:
                return SpecialInt(self.parse_int(val))
            elif key in SchoolData.field_types["is_float"]:
                return float(val)
            return val
        except Exception, e:
            import traceback; traceback.print_exc()
            print "Could not understand the value in col: %s for %s" % (key, self.name)
            sys.exit(1)
            #import traceback
            #traceback.print_exc()
        
    def _get_col_value(self, prefix):
        return self._data[prefix]

    def yes_is_1(self, val):
        if self.is_no_response(val):
            if self.voluteers_interviewed:
                return self.yes_is_1("no")
            return None
        return 1 if val.lower().strip() == "yes" else 0

    def no_is_1(self, val):
        if self.is_no_response(val):
            if self.voluteers_interviewed:
                return self.no_is_1("no")
            return None
        return 1 if val.lower().strip() == "no" else 0

    def parse_int(self, val):
        if self.is_no_response(val):
            return None
        return int(val)

    def parse_time(self, val):
        if self.is_no_response(val):
            return SpecialInt(None)
        return parse_willa_time(val, self.xls_datemode)

    def is_no_response(self, val):
        return val in ["Didn't answer", "", "None"]

    def score_rating(self, x):
        if x == None: return x

        if x >= 5: return 1
        return 0

    @property
    def visit(self):
        return self.A7

    @property
    def is_cooking_school(self):
        return self.B7 == "Cooking"

    @property
    def voluteers_interviewed(self):
        return "Volunteers (VOL)" in self.A10

    @property
    def name(self):
        return self.school_map[self.school_number]["schoolname"]

    @property
    def school_number(self):
        return int(self.A3)

    @property
    def school_type(self):
        return self.school_map[self.school_number]["school_type"]

    @property
    def visit_date(self):
        return self.A5

    @property
    def all_stock(self):
        if self.is_cooking_school:
            all_stock = [
                ("Pilchards", "C12"), 
                ("Rice", "C13"), 
                ("Curry Mince", "C14"), 
                ("Savoury Mince", "C15"),
                ("Samp", "C16"), 
                ("Sugar beans", "C17"), 
                ("Brown lentils", "C18")]
        else:
            all_stock = [
                ("Jam", "C24"), 
                ("Peanut butter", "C25")
            ]
        return all_stock

    @property
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
    def ind_meal_served_on_time(self):
        if self.D7a == None:
            return SpecialInt(None)
        return 2 if self.meal_served_on_time else 0

    @property
    def meal_served_efficiently(self):
        if self.D7a == None or self.D7b == None:
            return SpecialInt(None)
        diff = self.D7b - self.D7a
        return diff.minutes < 30

    @property
    def ind_meal_served_efficiently(self):
        return 1 if self.meal_served_efficiently else 0

    @property
    def food_in_good_condition(self):
        return (self.C5 + self.C6 + self.C7 + self.C8) * 0.5

    @property
    def menu(self):
        menu_name = "%s %s schools" % (self.school_type, "cooking" if self.is_cooking_school else "non-cooking")
        return self.menus[menu_name]

    @property
    def food_deviations(self):
        val_total_fed = self.B3
        val_date_of_visit = self.visit_date
        val_day_of_visit = constants.day_of_week[val_date_of_visit.weekday()]
        if val_day_of_visit == "Saturday": val_day_of_visit = "Friday"
        if val_day_of_visit == "Sunday": val_day_of_visit = "Monday"

        ranges = [] 
        for field_idx in range(8, 14):
            ingredient_type_field = "D%da" % field_idx
            ingredient_qty_field = "D%db" % field_idx

            ingredient = getattr(self, ingredient_type_field)

            # Don't look at ingredients that are not on the menu
            if ingredient in self.menu:
                ingredient_pack_qty = constants.ingredient_qtys[ingredient]
                amount_needed = val_total_fed * self.menu[ingredient][val_day_of_visit]
                amount_served = getattr(self, ingredient_qty_field) * ingredient_pack_qty
                if amount_needed == 0:
                    range_perc = 0
                else:
                    range_perc = abs(amount_served - amount_needed) / amount_needed

                # Don't include deviations where a quantity served is 0
                if amount_served > 0:
                    ranges.append((ingredient, range_perc))
        return ranges

    @property
    def ind_food_deviations(self):
        ranges = self.food_deviations 
        ranges_below_10 = [range_perc for range_perc in ranges if range_perc < 0.1]
        if len(ranges_below_10) == len(ranges):
            pt_ranges = 2
        elif len(ranges_below_10) >= 3:
            pt_ranges = 1
        else:
            pt_ranges = 0
        return pt_ranges
        
    @property
    def meal_delivery_score(self):
        return float(sum([
            self.D1, self.D3, self.D4, self.D5, self.D6, 
            self.ind_meal_served_on_time, self.ind_meal_served_efficiently, self.ind_food_deviations
        ]))

    @property
    def hygiene_score(self):
        return sum([
            self.E2, self.E3, self.E4, self.E5,
            self.E7, self.E8, self.E9, 
            self.E12, self.E13, 
            self.E10 * 0.5, self.E11 * 0.5,
            1 if (self.E5 == 0 and self.E6 == 1) else 0
        ])

    @property
    def staples_in_stock(self):
        all_stock = self.all_stock
        days_left = self.B6
        weeks_left = days_left / 5.0

        stock_surpluses = []
        for (ingredient, indicator) in self.all_stock:
            available_stock_days = getattr(self, "%sb" % indicator)
            feeding_days_per_week = sum([1 for x in self.menu[ingredient].values() if x > 0])
            required_stock_days = feeding_days_per_week * weeks_left
            surplus_stock_days = available_stock_days - required_stock_days
            
            stock_surpluses.append((ingredient, surplus_stock_days >= 0))
        return all([x[1] for x in stock_surpluses])

    @property
    def ind_staples_in_stock(self):
        return 2 if self.staples_in_stock else 0

    @property
    def ind_G1(self):
        return self.score_rating(self.G1)

    @property
    def ind_G2(self):
        return self.score_rating(self.G1)

    @property
    def ind_G3(self):
        return self.score_rating(self.G1)

    @property
    def ind_G4(self):
        return self.score_rating(self.G1)

    @property
    def stock_score(self):
        return sum([
            self.C1, self.C2, self.C3, self.C4, self.C9,
            self.C5 * 0.5, self.C6 * 0.5, self.C7 * 0.5, self.C8 * 0.5,
            self.C10 * 0.5, self.C11 * 0.5,
            self.ind_staples_in_stock
        ])
        return ind1 + ind2


    @property
    def staff_score(self):
        return sum([
            self.F1, self.F3, self.F4, self.F5, self. F6, self.F7,
            self.score_rating(self.G1), self.score_rating(self.G2), 
            self.score_rating(self.G3), self.score_rating(self.G4), 
        ])

