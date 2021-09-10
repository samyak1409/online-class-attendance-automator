# MADE REQUIRED CHANGES HERE:


# Must:

teacher = 'Ravi Jain'  # (since teacher will always be present in the meeting)

path_to_csv_dir = 'Attendance CSVs'  # path to the dir containing Google Chrome Extension generated CSV files of attendance

min_percent = 70  # minimum percentage criteria


# Only if you're generating sample data:

last_line_in_heading = 'Full Name'  # don't include double quotes ('"Full Name"' ❌ ; 'Full Name' ✔)

date_line_sub_str = 'Created on'  # some string which uniquely identifies date cell

path_to_sample_csv_dir = 'Sample CSV'  # path to the dir containing Sample CSV file (whose format will be followed to generate sample attendance data); must at least have one CSV in it


# Optional:

path_to_register = 'Attendance Register.xlsx'  # path to the Attendance Register

heading_row = 3  # row no. having headings like Roll Number, Name, Date

names_heading = last_line_in_heading  # heading for the names in the Attendance Register

start_column = None  # column no. from where the attendance insertion should start, None for default


########################################################################################################################


# Functions:

def get_students():
    """Yields a name from 'Students.data'."""
    for name in filter(lambda x: x, map(lambda x: x.strip(), sorted(open('Students.data').read().split('\n')))):
        yield name


def get_date(filename: str):
    """Returns :class:`datetime.date` (YYYY-MM-DD) from the CSV file name."""
    from datetime import date
    return date.fromisoformat(filename[:10])


# Debugging:

if __name__ == '__main__':

    for student in get_students():
        print(student)

    print()  # spacing

    from os import listdir
    csv = listdir('Sample Data')[0]
    print(csv, '->', get_date(filename=csv))
