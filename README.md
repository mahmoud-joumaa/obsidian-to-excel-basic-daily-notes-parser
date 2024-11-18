# Obsidian To Excel (Basic) Daily Notes Parser

A, rather basic, [Python](https://www.python.org/) program to graph (my cousin's) [Obsidian](https://obsidian.md/) daily notes into an excel workbook to track his "happy scores".

## The "Why"

My cousin tracks his happiness inside [Obsidian](https://obsidian.md/) and he thought it'd be cool if he could (automatically) visualize this inside an exported excel file.

## The "How"

1. The script parses the folder where your daily notes are stored (_the folder's location is hard-coded by the user_) and lists every file name in his directory (_excluding certain files that the user defines_).

2. The script then organizes the listed files by their names (_assumed to be in the format of 'dd-mm-YYYY.md'_).

	a. As the file names are being organized, the script parses each file and extracts the happy score from each file by searching for the line that comes directly after `#Happy-scale` so the file is very flexible in terms of format.

	b. Once all the files have been organized into their respective years and months, the lists are all sorted by name to ensure chronological order.

3. The chronologically-ordered list is then iterated through to write the data into an excel workbook with a user-defined name (_`happy_scale.xlsx` by default_).

	a. Each year has a separate spreadsheet in the excel workbook.

	b. Each spreadsheet has all the dates writter in the `A` column and all the scores written in the `B` column.

	c. If the script does not detect a happy score for a certain date, the score will default to `-1`.

4. Once the data is written into the corresponding columns, a line chart is drawn based on this data.

	a. The x-axis and the y-axis represent the dates and the scores, respectively.

	b. The line is of a blue color and the marker is set to automatic.

	c. The labels on the x-axis are set to be under the axis, and the data labels are set to be above each data point.

### Using the Script

To run the Python script, the following **must** be configured correctly:
- `DAILY_NOTES_DIR` at **line 8**: Set this variable to the folder which contains your daily notes (_e.g. `"C:\\Users\\<User-name>\\<Obsidian-vault>\\Daily Notes"`_)

To run the Python script, the following **may** be configured based on the user (_a.k.a. you_):
- `excel_workbook_name` at **line 6**: Set this variable to whatever the `.xlsx` file should be named (_e.g. `"happy.xlsx"`_).
> [!Note]
> The scirpt does some basic checking while parsing the file names to make sure of the format. However, this check is not fool-proof. It is up to the user to make sure the file names follow the format "dd-mm-YYYY" and exclude any files with names that do not follow this format.
- `files_to_exclude` at **line 10**: Set this variable to a list of file names to be ignored while parsing (_e.g. `["Daily Notes.md", "Daily Note Template.md"]`_)

Once the configured variables are set, the script can be started by running `python graph_happy_scale.py` in the terminal.

## Requirements

- [Obsidian](https://obsidian.md/)
- [Python](https://www.python.org/)
	- The [XlsWriter](https://pypi.org/project/XlsxWriter/) package stated in the `requirements.txt` file.

- To install Python, head over to [https://www.python.org/](https://www.python.org/) and follow the instructions on the website.
> [!Important]
> Ensure Python is added to the Path while downloading
- To install [XlsWriter](https://pypi.org/project/XlsxWriter/), run `python -m pip install xlswriter` in the terminal.
> [!Tip]
> It may be worth looking into Python virtual environments (venv) it you're starting to have a lot of dependencies downloaded globally.

## Known Issues

The script currently parses all files each time it's ran, so it may be a little resource-heavy.

Although an optimization probably exists, given that there are around 365 days a year on average, the daily notes would have to span multiple years for a considerable difference to be observed.

# Useful Resources

- [The official Python website](https://www.python.org/)
- [The full XlsWriter documentation](https://xlsxwriter.readthedocs.io/)

# Contact & FAQs

I'm not affiliated with any of the involved applications or technologies. At the time of working on this project, I'm simply a Computer Science graduate who enjoys working on custom solutions that help my friends have more fun when they use certain technologies. If you would like to reach out, feel free to send an email titled "Force Open Google Drive Files with Google Chrome Instead of the Default Browser" to mahmoud.j.eschool@gmail.com.

The source code for this project can be found at [https://github.com/mahmoud-joumaa/obsidian-to-excel-basic-daily-notes-parser](https://github.com/mahmoud-joumaa/obsidian-to-excel-basic-daily-notes-parser)
