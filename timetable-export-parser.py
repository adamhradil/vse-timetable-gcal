import datetime
import sys
import pandas as pd
from pathlib import Path


def main(argv):
    if len(argv) != 1:
        print("Wrong amount of arguments passed.")
        return 1
    else:
        file_path = argv[0]

    insis_export = pd.read_excel(file_path)

    result = []

    for row in insis_export.to_numpy():
        new_row = {}
        new_row['Subject'] = row[3] + ' - ' + row[4]
        new_row['Start Date'] = datetime.datetime.strptime(row[0].split(' ')[1], '%d.%m.%Y').strftime('%m/%d/%y')
        new_row['End Date'] = datetime.datetime.strptime(row[0].split(' ')[1], '%d.%m.%Y').strftime('%m/%d/%y')
        new_row['Start Time'] = row[1]
        new_row['End Time'] = row[2]
        new_row['Location'] = row[5]
        new_row['Description'] = 'Vyučující: ' + row[6]
        result.append(new_row)

    pd.DataFrame(result).to_csv(Path(file_path).stem + '.csv', index=False)


if __name__ == "__main__":
    main(sys.argv[1:])
