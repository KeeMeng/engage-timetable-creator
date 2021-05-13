import re
import datetime
from tkinter import messagebox
from os.path import dirname, join
from tkinter.filedialog import askopenfilename

add_initials = True

#ends calendar on this day below
#last day     YYYY/MM/DD
school_end = "2021/07/01".replace("/","")

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
RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL={school_end}T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T{assembly.group(3)}{assembly.group(4)}00
SUMMARY:Assembly / Tutorial
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:Assembly / Tutorial
TRIGGER:-PT5M
END:VALARM
END:VEVENT
"""

		#finding all other lessons
		data = []
		regex = "<span class=[\'\"]ttLessonText[\'\"]>(<strong>Lesson \d<\/strong><br ?\/?>)?(.*?)<br ?\/?>(\d\d):(\d\d) - (\d\d):(\d\d)<br ?\/?>(.*?)<br ?\/?>(.*?)(<br ?\/?>)?<\/span>"
		for (_, lesson, start_hour, start_min, end_hour, end_min, location, teacher, _) in re.findall(regex, html):
			data.append((lesson, start_hour, start_min, end_hour, end_min, location, teacher))

		data.append((None, None, None, None, None, None, None))

		count = 0
		skip = False
		while count < len(data)-1:
			(lesson, start_hour, start_min, end_hour, end_min, location, teacher) = data[count]
			(lesson2, start_hour2, start_min2, end_hour2, end_min2, _, teacher2) = data[count+1]

			if lesson2 != None:
				time_diff = (int(start_hour2) * 60 + int(start_min2)) - (int(end_hour) * 60 + int(end_min)) 
				if time_diff > 5 and time_diff < 180:
					# print(time_diff)
					break_start_hour = int(end_hour)
					break_start_min = int(end_min) + 5
					if break_start_min > 60:
						break_start_min -= 60
						break_start_hour += 1

					break_end_hour = int(start_hour2)
					break_end_min = int(start_min2) - 5
					if break_end_min > 60:
						break_end_min -= 60
						break_end_hour += 1
					break_start_hour = f"{break_start_hour:02d}"
					break_start_min = f"{break_start_min:02d}"
					break_end_hour = f"{break_end_hour:02d}"
					break_end_min= f"{break_end_min:02d}"

					# print([end_hour, end_min, start_hour2, start_min2])
					# print([break_start_hour, break_start_min, break_end_hour, break_end_min])

					lines += f"""
BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T{break_start_hour}{break_start_min}00
RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL={school_end}T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T{break_end_hour}{break_end_min}00
SUMMARY:Break
END:VEVENT
"""
			
			if skip:
				skip = False
				count += 1
				continue

			if (teacher == teacher2 or (lesson == "Study Period" and lesson2 == "Study Period")) and lesson == lesson2:
				end_hour = end_hour2
				end_min = end_min2
				skip = True

			# print([lesson, start_hour, start_min, end_hour, end_min, location, teacher])
			#increase date
			if start_hour == "08":
				date = datetime.datetime(int(year),int(month),int(day)) 
				date += datetime.timedelta(days=1)
				day = f"{date.day:02d}"
				month = f"{date.month:02d}"
				year = f"{date.year:02d}"

			first_last_name = re.search("(Mr|Mrs|Ms) (.*?)$", teacher)
			try:
				zoom_link = staff[first_last_name.group(2)]
				if add_initials:
					initials = f" ({''.join([word[0] for word in first_last_name.group(2).split(' ')]).upper()})"
				else:
					initials = ""
			except KeyError:
				teachers = re.findall("(Mr|Mrs|Ms) (.*?)((<br ?\/?>)| &amp;|$)", teacher)
				teacher = ""
				for (title, name, _, _) in teachers:
					teacher += f"{title} {name}, "
				teacher = teacher[:-2]
				zoom_link = ""
				initials = ""
			except AttributeError:
				teacher = ""
				zoom_link = ""
				initials = ""
			lines += f"""
BEGIN:VEVENT
DTSTART;TZID=Asia/Hong_Kong:{year}{month}{day}T{start_hour}{start_min}00
RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL={school_end}T000000Z
DTEND;TZID=Asia/Hong_Kong:{year}{month}{day}T{end_hour}{end_min}00
SUMMARY:{lesson}{initials}
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
			count += 1


#creates ics file
with open(join(path, "timetable.ics"), "w") as ics_file:
	lines += "\nEND:VCALENDAR"
	ics_file.writelines(lines)

messagebox.showinfo("Timetable Creator", f"Saved to: {join(path, 'timetable.ics')}")
