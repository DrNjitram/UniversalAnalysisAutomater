import array
import binascii
import os
import re
import xlsxwriter
from collections import defaultdict


def own_round(value):
    if 'e' in value:
        return value[:7] + "E" + str(int(value[-3:]))
    elif -1 < float(value) < 1:
        if len(value) < 8:
            return value.ljust(8, "0")
        return value[:-len(value.lstrip("-0."))] + value.lstrip("-0.")[:7]
    elif len(value) > 8:
        return value[:8]
    else:
        return value.ljust(8, "0")


def get_column(cols, target):
    col_no = cols.index(target)
    return chr(ord('A') + col_no)


def avg(lst):
    return sum(lst)/len(lst)


def inc_column(col, inc):
    if col == "":
        return "A"
    if inc == 0:
        return col
    if col[-1] == "Z":
        return inc_column(inc_column(col[:-1], 1) + "A", inc - 1)
    elif inc > 1:
        return inc_column(col, inc - 1)
    else:
        return col[:-1] + chr(ord(col[-1]) + 1)


data_divisor = b'000c000c'
garbage_divisor = b'0000c8c20000a041'
delimiter = "\t"

path = os.path.abspath(os.getcwd())  # Get working directory

workbook = xlsxwriter.Workbook('auto generated DMA overview.xlsx')  # Create Excel file
overview = workbook.add_worksheet("Overview")  # Add overview sheet
averages = workbook.add_worksheet("Averages") # Add quick calculations sheet
end_points = workbook.add_worksheet("End Points") # Add quick end_piints sheet
chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})  # Chart for all data

all_series = {}  # Container for all data
all_end_points = {}  # Container for all end_points

for root, dirs, files in os.walk(path):  # For all files in folder and sub folders
    for filename in files:  # For all files in this folder
        if not filename.split('.')[-1].isnumeric() and filename.split('.')[-1].endswith(('txt', 'zip', 'exe', 'xlsx')):
            continue

        print(os.path.join(root, filename))  # Print current file

        # Clean up filename
        if filename.split('.')[-1].isnumeric():
            clean_filename = "".join(filename.split(".")[:-1])
        else:
            clean_filename = filename

        # Read file
        with open(os.path.join(root, filename), 'rb') as f:
            binary_content = f.read()

        # Read content into sub variables
        text_content_hex, hex_content_hex = binascii.hexlify(binary_content).split(data_divisor)

        # Discard garbage at end of file
        hex_content_hex, _ = hex_content_hex.split(garbage_divisor)

        # Decode and add delimiter into text
        text_content_text = binascii.unhexlify(text_content_hex)[:-1].decode('utf-16')
        text_content_text = re.sub(r'^([^ ]+) ', r'\g<1>' + delimiter, text_content_text, flags=re.M)

        # Create array of all floats
        floats_sequence = array.array('f', binascii.unhexlify(hex_content_hex)).tolist()

        # Get number of signals
        NSig = int(re.findall(r'Nsig' + delimiter + r'(\d+)', text_content_text)[0])

        # Create text from list of floats
        hex_content_text = "\n".join([delimiter.join([own_round(str(flt)) for flt in flt_list]) for flt_list in [floats_sequence[i:i + NSig] for i in range(0, len(floats_sequence), NSig)]])

        # Truncate filename because sheets cant be longer than 31 chars
        clean_filename = clean_filename[:31]

        # Add file as sheet to excel workbook
        try:
            curr_sheet = workbook.add_worksheet(clean_filename)
        except:
            clean_filename = filename[:31]
            curr_sheet = workbook.add_worksheet(clean_filename)

        with open(os.path.join(root, clean_filename + ".txt"), 'w') as f:
            f.write(text_content_text.replace("\r\n", "\n"))
            f.write("Orgfile" + delimiter + os.path.join(path, filename) + "\n")
            f.write("StartOfData\n")
            f.write(hex_content_text + "\n")

        columns = []
        index = 0

        # Write text to excel sheet
        for line in text_content_text.split("\n"):
            column = "A"
            index += 1

            if 'Sig' in line and 'NSig' not in line:
                columns.append(line.split(delimiter)[-1].strip())

            for number in line.split(delimiter):
                curr_sheet.write_string(column + str(index), number)
                column = chr(ord(column) + 1)

        curr_sheet.write("A" + str(index + 1), "Orgfile" + delimiter + os.path.join(path, filename))
        curr_sheet.write("A" + str(index + 2), "StartOfData")

        index += 2
        data_range = [index + 1, 0]
        lastline = ""

        # Split line into its parts
        hex_content_float = [[0 if number.strip() == "" else float(number) for number in line.split(delimiter)] for line in hex_content_text.split("\n")]

        # Write number data to excel sheet
        for line in hex_content_float:
            column = "A"
            index += 1

            #curr_sheet.write_row(index, ord(column), line)

            try:
                for number in line:
                    curr_sheet.write_number(column + str(index), number)
                    column = chr(ord(column) + 1)
            except:
                pass

        data_range[1] = index

        # Create series from this datafile
        current_series = {
            'name': clean_filename,
            'categories': '=\'' + clean_filename + '\'!$' + get_column(columns, 'Strain (%)') + '$' + str(data_range[0]) + ':$' + get_column(columns, 'Strain (%)') + '$' + str(data_range[1]),
            'values': '=\'' + clean_filename + '\'!$' + get_column(columns, 'Stress (MPa)') + '$' + str(data_range[0]) + ':$' + get_column(columns, 'Stress (MPa)') + '$' + str(data_range[1]),
            'data_labels': {'series_name': True, 'custom': [{'delete': True} for _ in range(data_range[0], data_range[1])]},
        }

        # Add series to chart
        chart.add_series(current_series)

        # Add series to data collection
        all_series[filename] = current_series
        try:
            all_end_points[filename] = (hex_content_float[-1][columns.index('Strain (%)')], hex_content_float[-1][columns.index('Stress (MPa)')], clean_filename)
        except:
            pass

# Set titles and ranges for axes
chart.set_x_axis({'name': 'Strain (%)', 'min': '0', 'max': '120'})
chart.set_y_axis({'name': 'Stress (MPa)', 'min': '0', 'max': '10'})

# Insert chart in sheet
overview.insert_chart('A1', chart)


if os.path.exists(os.path.join(path, 'grouping.txt')) and os.path.getsize(os.path.join(path, 'grouping.txt')) > 100:
    groupings = defaultdict(list)
    statistics = defaultdict(list)  # (Strain, Stress)
    found_filenames = []
    with open(os.path.join(path, 'grouping.txt'), 'r') as f:
        for line in f.readlines():
            try:

                if not line.strip().startswith("#") and line.split(":")[-1].strip() is not "" and line.strip() is not "":
                    part1, part2 = line.split(":")
                    found_filenames.append(part1.strip())
                    clean_part1 = part1
                    if part1.split('.')[-1].isnumeric():
                        clean_part1 = "".join(part1.split(".")[:-1])

                    for section in part2.split(";"):
                        group, name = section.split(",")

                        group = group.strip()[1:]
                        name = name.strip()[:-1]

                        if group in groupings:
                            new_group = groupings[group]
                            new_group.append((part1, clean_part1 + " " + name))
                            groupings[group] = new_group
                        else:
                            groupings[group] = [(part1, clean_part1 + " " + name)]

                        statistics_group = group + ";" + name
                        if statistics_group in statistics and part1 in all_end_points:
                            new_group = statistics[statistics_group]
                            new_group.append(all_end_points[part1])
                            statistics[statistics_group] = new_group
                        else:
                            statistics[statistics_group] = [all_end_points[part1]]

                elif line.split(":")[-1].strip() is "" and line.strip() is not "":
                    found_filenames.append(line.split(":")[0].strip())

            except:
                print("Error parsing line: " + line + "Remove any wrong , ; ( ) :\n")

    chart_index = 1
    for key, value in groupings.items():
        print("Creating chart group " + key)
        chart = workbook.add_chart({'type': 'scatter',
                                    'subtype': 'smooth'})
        chart.set_x_axis({'name': 'Strain (%)', 'min': '0'})
        chart.set_y_axis({'name': 'Stress (MPa)', 'min': '0'})
        chart.set_title({'name': key})

        for filename, title in value:
            current_series = all_series[filename]
            current_series['name'] = title
            chart.add_series(current_series)

        overview.insert_chart('A' + str(chart_index), chart)
        chart_index += 2

    index = 2
    averages.write_string("B1", "Average Strain")
    averages.write_string("C1", "Average Stress")
    for key, value in statistics.items():
        print("Creating statistics for group " + key)
        avg_strain = avg([val[0] for val in value])
        avg_stress = avg([val[1] for val in value])

        averages.write_string("A" + str(index), key)
        averages.write_number("B" + str(index), avg_strain)
        averages.write_number("C" + str(index), avg_stress)

        index += 1

    column = "A"
    for key, value in statistics.items():
        end_points.write_string(column + "1", key + " Sequence")
        end_points.write_string(inc_column(column, 1) + "1", key + " Strain")
        end_points.write_string(inc_column(inc_column(column, 1), 1) + "1", key + " Stress")
        index = 2

        for val in value:
            end_points.write_string(column + str(index), val[2])
            end_points.write_number(inc_column(column, 1) + str(index), val[0])
            end_points.write_number(inc_column(inc_column(column, 1), 1) + str(index), val[1])
            index += 1

        column = inc_column(inc_column(inc_column(column, 1), 1), 1)

    any_unfound = False
    for key in all_series.keys():
        if key not in found_filenames:
            if not any_unfound:
                print("Not found files:")
                any_unfound = True

            print(key)

    if not any_unfound:
        print("All filenames found properly")

else:
    print("Creating grouping file")
    with open(os.path.join(path, 'grouping.txt'), 'w') as f:
        f.write("#Format each line as: filename: (grouping, name in group); (other group, other name)\n")
        f.write("#Example: measurement1.001: (compound 1, amount); (compound 2, amount)")
        for filename in all_series.keys():
            f.write(filename + ":\n")

print("Finalising Excel Sheet")
try:
    workbook.close()
except:
    print("Close Excel sheet: auto generated DMA overview.xlsx and try again")

input("\nPress any key to close the program")