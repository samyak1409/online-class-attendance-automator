# MADE REQUIRED CHANGES HERE:


# Must:

path_to_csv_dir = 'Attendance CSVs'  # path to the dir containing Google Chrome Extension generated CSV files of attendance

teacher = 'Ravi Jain'  # (since teacher will always be present in the meeting)

min_percent = 70  # minimum percentage criteria

# Check the pattern and find some string which uniquely identify those particular lines in Attendance CSVs:

date_line_sub_str = 'Created on'

last_line_sub_str = 'Full Name'  # don't include double quotes ('"Whatever Sub-Text"' ❌ ; 'Whatever Sub-Text' ✔)


# Only if you're generating sample data:

path_to_sample_csv_dir = 'Sample CSV'  # path to the dir containing Sample CSV file (whose format will be followed to generate sample attendance data); must at least have one CSV in it


# Optional:

path_to_register = 'Attendance Register.xlsx'  # path to the Attendance Register

heading_row = 3  # row no. having headings like Roll Number, Name, Date

names_heading = "Full Name"  # heading for the names in the Attendance Register

start_column = None  # column no. from where the attendance insertion should start, None for default

include_year = False  # include year (-yyyy) in the dates in Attendance Register or not


########################################################################################################################


# Attributes:

studs = 'Students.data'
attrs = 'Attributes.py'
most_wanted = 'Most Wanted.txt'


# Functions:

def get_students():
    """Returns names from :var:`studs`."""
    # print(open(studs).read().split('\n'))  #debugging
    names = list(filter(lambda x: x, sorted(map(lambda x: x.strip().title(), open(studs).read().split('\n')))))  # doesn't matter what's outside, the data will be well formatted after coming inside the program! (title cased etc.)
    # print(names)  #debugging
    if len(names) != len(set(names)):  # duplicate names check
        from collections import Counter
        histogram = Counter(names)
        raise ValueError(f'"{studs}" contains duplicate names {list(filter(lambda x: histogram[x] > 1, histogram))}, please fix it and run the program again.')
    return names


def get_date(filename: str):
    """Returns :class:`datetime.date` (YYYY-MM-DD) from the CSV file name."""
    from datetime import date
    len_of_date = 10  # e.g. len('2021-11-22')
    for i in range(len(filename)-len_of_date+1):  # traversing through the filename (date can be anywhere in b/w)
        try:
            return date.fromisoformat(filename[i:i+len_of_date])
        except ValueError:
            continue
    else:
        raise ValueError(f'DATE NOT FOUND IN THE FILENAME "{filename}"')


# Debugging:

if __name__ == '__main__':

    for student in get_students():
        print(student)

    print()  # spacing

    from glob import iglob
    from os import chdir
    chdir(path_to_csv_dir)
    for csv in iglob('*.csv'):
        print(csv, '->', get_date(filename=csv))
