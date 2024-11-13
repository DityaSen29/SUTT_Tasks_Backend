import pandas as pd
import json

def get_instructor_list(df, index):
    instructors = []
    current_idx = index

    first_instructor = df.iloc[current_idx]['INSTRUCTOR-IN-CHARGE / Instructor']
    if pd.notna(first_instructor):
        instructors.append(first_instructor.strip())

    while current_idx + 1 < len(df):
        current_idx += 1
        row = df.iloc[current_idx]
        if pd.notna(row['SEC']):
            current_idx -= 1
            break
        if pd.notna(row['INSTRUCTOR-IN-CHARGE / Instructor']):
            instructors.append(row['INSTRUCTOR-IN-CHARGE / Instructor'].strip())
            
        #implemented such that we can see till where a particular section lasts till
        #for multiple instuctors for same course

    return instructors, current_idx

def get_section_type(section_code):
    if not section_code:
        return "Unknown"
        
    prefix = section_code[0].upper()
    if prefix == 'L':
        return "Lecture"
    elif prefix == 'P':
        return "Practical"
    elif prefix == 'T':
        return "Tutorial"
    else:
        return "Unknown"

slot_to_time = {
    1: "8:00 - 9:00", 
    2: "9:00 - 10:00", 
    3: "10:00 - 11:00", 
    4: "11:00 - 12:00",
    5: "12:00 - 1:00", 
    6: "1:00 - 2:00", 
    7: "2:00 - 3:00", 
    8: "3:00 - 4:00", 
    9: "4:00 - 5:00"
}

def parse_days_and_hours(days_and_hours):
    if not days_and_hours:
        return []

    days_hours = days_and_hours.split()
    timings = []
    i = 0

    while i < len(days_hours):
        days = []
        
        while i < len(days_hours) and days_hours[i].isalpha():
            days.append(days_hours[i])
            i += 1
        
        if days_hours[i].isdigit():
            slot = int(days_hours[i])
            time_range = slot_to_time.get(slot, "")
            for day in days:
                timings.append({"day": day, "slots": [time_range]})
            i += 1
        else:
            i += 1

    return timings

def parse_excel_as_dict(file_path):
    excel_data = pd.ExcelFile(file_path)
    timetable_data = {}

    for sheet in excel_data.sheet_names:
        print(f"Parsing {sheet}")
        df = pd.read_excel(file_path, sheet_name=sheet, skiprows=1)
        df.columns = ['COM COD', 'COURSE NO.', 'COURSE TITLE', 'CREDIT_L', 'CREDIT_P', 'CREDIT_U', 
                      'SEC', 'INSTRUCTOR-IN-CHARGE / Instructor', 'ROOM', 'DAYS & HOURS', 
                      'MIDSEM DATE & SESSION', 'COMPRE DATE & SESSION']

        course_data = {
            "course_code": None,
            "course_title": None,
            "credits": {
                "lecture": "-",
                "practical": "-",
                "units": "-"
            },
            "sections": []
        }

        idx = 0

        while idx < len(df):
            row = df.iloc[idx]

            if pd.isna(row['SEC']):
                idx += 1
                continue

            if pd.notna(row['COM COD']):
                course_data["course_code"] = row['COURSE NO.']
                course_data["course_title"] = row['COURSE TITLE']
                course_data["credits"]["lecture"] = str(row['CREDIT_L']).strip() if pd.notna(row['CREDIT_L']) else "-"
                course_data["credits"]["practical"] = str(row['CREDIT_P']).strip() if pd.notna(row['CREDIT_P']) else "-"
                course_data["credits"]["units"] = str(row['CREDIT_U']).strip() if pd.notna(row['CREDIT_U']) else "-"

            if pd.notna(row['SEC']):
                instructors, last_idx = get_instructor_list(df, idx)
                
                section_data = {
                    "section_type": get_section_type(row['SEC']),
                    "section_number": row['SEC'],
                    "instructors": instructors,
                    "room": row['ROOM'] if pd.notna(row['ROOM']) else "",
                    "timing": parse_days_and_hours(row['DAYS & HOURS']) if pd.notna(row['DAYS & HOURS']) else [],
                    "midsem_date": row['MIDSEM DATE & SESSION'] if pd.notna(row['MIDSEM DATE & SESSION']) else "N/A",
                    "compre_date": row['COMPRE DATE & SESSION'] if pd.notna(row['COMPRE DATE & SESSION']) else "N/A"
                }

                course_data["sections"].append(section_data)
                idx = last_idx

            idx += 1

        timetable_data[sheet] = course_data

    return timetable_data


def save_json(data, output_file="timetable_data.json"):
    with open(output_file, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, indent=2, ensure_ascii=False)
    print("JSON file has been created as timetable_data.json")

if __name__ == "__main__":
    try:
        file_path = r"C:\Users\Ditya\Downloads\TimetableWorkbook_SUTT_Task1.xlsx"
        timetable_data = parse_excel_as_dict(file_path)
        save_json(timetable_data)
    except Exception as e:
        print(f"An error occurred: {e}")
