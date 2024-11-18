from os import listdir
from os.path import isfile, join
import xlsxwriter

# Change this to name the generated excel file
excel_workbook_name = "happy_scale.xlsx"
# Change this to the absolute or relative path of the folder that contains all your daily notes
DAILY_NOTES_DIR = ".\\"
# Change this list to include the names of files you DON'T want the script to parse
files_to_exclude = []

def is_valid_daily_note(daily_note):
	if len(daily_note) != 13:
		return False
	if daily_note[2] != '-' or daily_note[5] != '-' or daily_note[10] != '.':
		return False
	return True

if __name__ == "__main__":
	# Collect all daily notes
	daily_notes = listdir(DAILY_NOTES_DIR)
	daily_notes = [daily_note for daily_note in daily_notes if isfile(join(DAILY_NOTES_DIR, daily_note)) and is_valid_daily_note(daily_note) and daily_note not in files_to_exclude]

	# Organize the daily notes
	daily_notes_by_date = dict()

	for daily_note in daily_notes:
		daily_note_date = daily_note.split('.')[0].split('-')
		current_year = daily_note_date[2]
		current_month = daily_note_date[1]
		happy_score = -1
		# Parse the file to get the actual score /10
		try:
			daily_note_file = open(f"{DAILY_NOTES_DIR}\\{daily_note}", 'r')
			daily_note_file_data = daily_note_file.read().split('\n')
			happy_score = daily_note_file_data[daily_note_file_data.index("#Happy-scale")+1].split('/')[0]
			if not happy_score:
				happy_score = -1
		except Exception as E:
			print(f"An error has occurred while reading {daily_note}...")
			print(E)
		# Save the score with its corresponding file name in the dictionary
		try:
			daily_notes_by_date[current_year][current_month].append((daily_note, happy_score))
		except KeyError as E:
			if str(E)[1:-1] == current_year:
				daily_notes_by_date[current_year] = {}
			daily_notes_by_date[current_year][current_month] = [(daily_note, happy_score)]

	# Make sure the files are sorted by date
	for year in daily_notes_by_date:
		for month in daily_notes_by_date[year]:
			daily_notes_by_date[year][month].sort()

	# Write to the .xlsx file
	workbook = xlsxwriter.Workbook(excel_workbook_name)
	# Create a sheet for every year
	for year in daily_notes_by_date:
		current_year = workbook.add_worksheet(year)
		current_year.set_column("A:A", 20)
		current_year.set_column("B:B", 20)
		current_year.write("A1", "Date")
		current_year.write("B1", "Happy Score")
		row_number = 2
		for month in daily_notes_by_date[year]:
			for date_data in daily_notes_by_date[year][month]:
				current_year.write(f"A{row_number}", date_data[0][:-3])
				current_year.write_number(f"B{row_number}", float(date_data[1]))
				row_number += 1
		chart = workbook.add_chart({"type": "line"})
		chart.set_title({"name": f"Happy Scale Line Curve for {year}"})
		chart.set_x_axis({
			"name": "Date",
			"label_position": "low"
		})
		chart.set_y_axis({"name": "Happy Score"})
		chart.add_series({
			"name": "Happy Score",
			"categories": [str(current_year.name), 1, 0, len(daily_notes)+1, 0],
			"values": [str(current_year.name), 1, 1, len(daily_notes)+1, 1],
			"data_labels": {
				"value": True,
				"position": "above"
			},
			"line": {"color": "blue"},
			"marker": {"type": "automatic"}
		})
		current_year.insert_chart("C1", chart)

	workbook.close()
