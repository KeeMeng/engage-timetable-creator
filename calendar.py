import re
import datetime
from tkinter import messagebox
from os.path import dirname, join
from tkinter.filedialog import askopenfilename

messagebox.showinfo("Timetable Creator", "Select Week A first, and then Week B HTML File: ")

#GUI input Week A and Week B html files
files = [askopenfilename(), askopenfilename()]

path = dirname(files[0])

#default starting of ics file
lines = """
BEGIN:VCALENDAR
VERSION:2.0
CALSCALE:GREGORIAN
BEGIN:VTIMEZONE
TZID:Asia/Hong_Kong
TZURL:http://tzurl.org/zoneinfo-outlook/Asia/Hong_Kong
X-LIC-LOCATION:Asia/Hong_Kong
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
"""

#dictionary containing names of teachers and their zoom links
staff = {
	"name": "link"
}
#removed for privacy reasons

#loop for Week A and B
for f in files:
	if f != "":
		with open(f, "r", encoding='utf-8') as file:
		    html = file.read()

		today = re.search("Week Beginning (\d\d)\/(\d\d)\/(\d\d\d\d)<\/span>", html)
		(day, month, year) = today.group(1, 2, 3)

		# tutorial/assembly has a different format compared to other lessons
		assembly = re.search("(\d\d):(\d\d) - (\d\d):(\d\d).*Assembly / Tutorial", html)
		lines += f"""
BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T{assembly.group(1)}{assembly.group(2)}00
RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL=20210701T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T{assembly.group(3)}{assembly.group(4)}00
SUMMARY:Assembly / Tutorial
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:Assembly / Tutorial
TRIGGER:-PT5M
END:VALARM
END:VEVENT

BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T095500
RRULE:FREQ=WEEKLY;INTERVAL=2;BYDAY=MO,TU,WE,TH,FR;UNTIL=20210701T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T101000
SUMMARY:Break
END:VEVENT

BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T115500
RRULE:FREQ=WEEKLY;INTERVAL=2;BYDAY=MO,TU,WE,TH,FR;UNTIL=20210701T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T122000
SUMMARY:Break
END:VEVENT

BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T140500
RRULE:FREQ=WEEKLY;INTERVAL=2;BYDAY=MO,TU,WE,TH,FR;UNTIL=20210701T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T153000
SUMMARY:Break
END:VEVENT
"""

		#finding all other lessons
		regex = "<span class=[\'\"]ttLessonText[\'\"]>(<strong>Lesson \d<\/strong><br ?\/?>)?(.*?)<br ?\/?>(\d\d):(\d\d) - (\d\d):(\d\d)<br ?\/?>(.*?)<br ?\/?>(.*?)(<br ?\/?>)?<\/span>"
		for (_, lesson, start_hour, start_min, end_hour, end_min, location, teacher, _) in re.findall(regex, html):
			# print([lesson, start_hour, start_min, end_hour, end_min, location, teacher])
			#increase date
			if start_hour == "08":
				date = datetime.datetime(int(year),int(month),int(day)) 
				date += datetime.timedelta(days=1)
				day = f"{date.day:02d}"
				month = f"{date.month:02d}"
				year = f"{date.year:02d}"

# 				# Break 1
# 				lines += f"""
# BEGIN:VEVENT
# DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T095500
# RRULE:FREQ=WEEKLY;INTERVAL=2
# DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T101000
# SUMMARY:Break
# END:VEVENT
# """
# 				# Break 2
# 				lines += f"""
# BEGIN:VEVENT
# DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T115500
# RRULE:FREQ=WEEKLY;INTERVAL=2
# DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T122000
# SUMMARY:Break
# END:VEVENT
# """
	
# 				# Break 3
# 				lines += f"""
# BEGIN:VEVENT
# DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T140500
# RRULE:FREQ=WEEKLY;INTERVAL=2
# DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T153000
# SUMMARY:Break
# END:VEVENT
# """



			first_last_name = re.search("(Mr|Mrs|Ms) (.*?)$", teacher)
			try:
				zoom_link = staff[first_last_name.group(2)]
			except KeyError:
				teachers = re.findall("(Mr|Mrs|Ms) (.*?)((<br ?\/?>)| &amp;|$)", teacher)
				teacher = ""
				for (title, name, _, _) in teachers:
					teacher += f"{title} {name}, "
				teacher = teacher[:-2]
				zoom_link = ""
			except AttributeError:
				teacher = ""
				zoom_link = ""

			lines += f"""
BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T{start_hour}{start_min}00
RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL=20210701T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T{end_hour}{end_min}00
SUMMARY:{lesson}
URL:{zoom_link}
DESCRIPTION:{teacher}
LOCATION:{location}
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:{lesson}
TRIGGER:-PT5M
END:VALARM
END:VEVENT
"""


#creates ics file
with open(join(path, "timetable.ics"), "w") as ics_file:
	lines += "\nEND:VCALENDAR"
	ics_file.writelines(lines)

messagebox.showinfo("Timetable Creator", f"Saved to: {join(path, 'timetable.ics')}")
