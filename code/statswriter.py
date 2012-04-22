import csv

class StatsWriter(object):
    def __init__(self, fp):
        self.w = csv.writer(fp)
        self.fp = fp
        self.headers = [
            "School number", 
            "Visit number",
            "Handling of stock score (%)",
            "Delivery of school meal score (%)",
            "Safety and hygiene score (%)", 
            "Staff score (%)", 
            "Score for your school (%)",
        ]

        self.w.writerow(self.headers)

    def write_stats(self, row):
        row[2] *= 10
        row[3] *= 10
        row[4] *= 10
        row[5] *= 10
        row[6] *= 100
        self.w.writerow(row)
        self.fp.flush()

