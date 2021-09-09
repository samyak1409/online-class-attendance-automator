# Imports:
from csv import reader, writer
from datetime import timedelta
from random import sample, randint
from os import chdir
from glob import iglob
from Attributes import path_to_csv_dir, last_line_in_heading, date_line_sub_str, teacher, get_students, get_date


names = list(get_students())
strength = len(names)

chdir(path_to_csv_dir)

try:
    src_csv = next(iglob('*.csv'))  # CSV whose format will be followed
except StopIteration:
    exit(f'In order to generate some sample attendance data, \'{path_to_csv_dir}\' dir must have at least one CSV file to follow the format of. \nExiting.')
print('\nSource CSV:', src_csv)

# Reading Headings:
with open(src_csv, newline='', encoding='utf-8-sig') as file:
    # encoding='utf-8-sig' -> https://stackoverflow.com/questions/34399172/why-does-my-python-code-print-the-extra-characters-%c3%af-when-reading-from-a-tex

    data = reader(file)

    headings = []  # will be same for all CSVs
    print('\nHEADINGS ->')  # debugging

    while True:
        cell = next(data)[0]  # taking only one cell from every row
        print(cell)  # debugging

        headings.append([cell])  # add

        if cell == last_line_in_heading:
            break


# Main:

start_date = new_date = get_date(filename=src_csv)
remaining_str = src_csv[10:]

for i in range(1, int(input('\nHow many CSVs? > '))+1):

    new_date += timedelta(days=1)
    # print(new_date)  #debugging

    dst_csv = f'{new_date}{remaining_str}'

    with open(dst_csv, mode='w', newline='') as file:
        # w -> overwrite if exist else create and write

        writer_ = writer(file)

        # Copying headings:
        for line in headings:
            if date_line_sub_str in line[0]:  # update date
                line = [line[0].replace(str(start_date), str(new_date))]
            writer_.writerow(line)

        # Marking attendance randomly:
        for name in sorted(sample(population=names, k=randint(strength//2, strength))+[teacher]):
            writer_.writerow([name])

    print(f'{i})', dst_csv)

print('\nSUCCESS')
