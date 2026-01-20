import pandas as pd

class Logger:
    def __init__(self):
        self.records = []
        self.failures = []

    def log(self, **kwargs):
        self.records.append(kwargs)

    def fail(self, **kwargs):
        self.failures.append(kwargs)

    def save(self):
        if self.records:
            pd.DataFrame(self.records).to_excel('output/archive_list.xlsx', index=False)
        if self.failures:
            pd.DataFrame(self.failures).to_csv('output/classify_log.csv', index=False)