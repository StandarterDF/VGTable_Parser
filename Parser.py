import pandas, re, os, json
import os.path as path

VERBOSE = False

"""
# VGTU Table Lesson per day Structure
-   0) Week day
-   1) Time (From \ To)
-   2) None
-   3-[n-1]) ["Name", "Class Number"]
-   [n]) "Class Number"

# VGTU Structure
#
# Номер занятия, Название занятия, Начало занятия, Конец занятия, Преподаватель, Аудитория, Подгруппа
#
{
    "Numerator": {
        "Monday": [],
        "Tuesday": [],
        "Wednesday": [],
        "Thursday": [],
        "Friday": [], 
        "Saturday": []
    },
    "Denominator": {
        "Monday": [],
        "Tuesday": [],
        "Wednesday": [],
        "Thursday": [],
        "Friday": [], 
        "Saturday": []
    }
}
"""
DaysRus = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"]
DaysEng = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
SepInt = [1, 0]
SepStr = ["Numerator", "Denominator"]

CurDir = path.dirname(path.abspath(__file__))
Excel = CurDir +  "/Excel"

def ListToList(Item, List1: list, List2: list):
    return List2[List1.index(Item)]

def LessonToList(DayList: dict, DayLesson: list, Separator: int):
    # Номер занятия (0), Название занятия (1), Начало занятия (2), Конец занятия (3), Преподаватель (4), Аудитория (5), Подгруппа (6)
    Lesson = DayLesson[Separator]
    Result = [0, 0, 0, 0, 0, 0, 0]
    Result[0] = DayList[Lesson[0]].index(DayLesson) + 1
    LessonNumber = Result[0]
    if Lesson[-1] != "None":
        RegexTime = re.findall("[0-9]{1,}:[0-9]{1,}:[0-9]{1,}", Lesson[1])
        RegexName = re.findall("\(.*\)", Lesson[3])[0] if len(re.findall("\(.*\)", Lesson[3])) > 0 else "Неизвестно"
        Result[1] = Lesson[3].replace(RegexName, "")
        Result[2] = RegexTime[0]
        Result[3] = RegexTime[1]
        Result[4] = re.sub("([\( ] | [ \)] | [()])", "", RegexName)
        Result[5] = Lesson[-1]
        Result[6] = 0
        if VERBOSE: print(Result)
        return [0, Result]
    else:
        Result = []
        for Data in range(3, len(Lesson) - 1, 2):
            if Lesson[Data] != "None":
                RegexTime = re.findall("[0-9]{1,}:[0-9]{1,}:[0-9]{1,}", Lesson[1])
                RegexName = re.findall("\(.*\)", Lesson[Data])[0] if len(re.findall("\(.*\)", Lesson[Data])) > 0 else "Неизвестно"
                TempResult = [0, 0, 0, 0, 0, 0, 0]
                TempResult[0] = LessonNumber
                TempResult[1] = Lesson[Data].replace(RegexName, "")
                TempResult[2] = RegexTime[0]
                TempResult[3] = RegexTime[1]
                TempResult[4] = re.sub("([\( ] | [ \)] | [()])", "", RegexName)
                TempResult[5] = Lesson[Data + 1] if Lesson[Data + 1] != "None" else "0000" 
                TempResult[6] = (((Data - 2) + 1)//2)
                Result.append(TempResult)
                #print(TempResult)
        #print(Result)
        return [1, Result]
def ExcelToDayList(Filename: str) -> dict:
    DB = pandas.read_excel(Filename, skiprows=3).fillna("None").values

    DictByDay = {
        "Понедельник": [],
        "Вторник": [],
        "Среда": [],
        "Четверг": [],
        "Пятница": [],
        "Суббота": []
    }
    
    for i in range(2, len(DB)):
        if DB[i][0] == "None": DB[i][0] = DB[i - 1][0]
        if DB[i][1] == "None": DB[i][1] = DB[i - 1][1]
    # for line in DB:
    #     print(list(line), f"Lenght: {len(list(line))}")
    for i in range(2, len(DB), 2):
        ListDay = []
        ListDay.append(list(DB[i]))
        ListDay.append(list(DB[i - 1]))
        DictByDay[DB[i][0]].append(ListDay)
    return DictByDay

def DayListToJSON(DayList: dict, GroupName):
    ResultJSON = {
        "Numerator": {
            "Monday": [],
            "Tuesday": [],
            "Wednesday": [],
            "Thursday": [],
            "Friday": [], 
            "Saturday": []
        },
        "Denominator": {
            "Monday": [],
            "Tuesday": [],
            "Wednesday": [],
            "Thursday": [],
            "Friday": [], 
            "Saturday": []
        }
    }
    for Day in DayList.keys():
        for Lesson in DayList[Day]:
            #print(Lesson)
            for DateSeparator in range(2):
                ResStyle, TLesson = LessonToList(DayList, Lesson, DateSeparator)
                if ResStyle == 0:
                    ResultJSON[ListToList(DateSeparator, SepInt, SepStr)][ListToList(Day, DaysRus, DaysEng)].append(TLesson)
                else:
                    for TTLesson in TLesson:
                        ResultJSON[ListToList(DateSeparator, SepInt, SepStr)][ListToList(Day, DaysRus, DaysEng)].append(TTLesson)
                #print(TLesson)
    return {GroupName: ResultJSON}


FinalResult = {
    "groups": {
        
    }
}

for TableIterator in os.listdir(Excel):
    try:
        Group = ExcelToDayList(Excel + "/" + TableIterator)
        GroupTable = DayListToJSON(Group, TableIterator.split(".")[0])
        FinalResult["groups"].update(GroupTable)
        print(f"#=> Parse file {TableIterator} (SUCCESS)")
    except Exception as E:
        print(f"#=> Can't parse file {TableIterator} ({str(E).upper()})")
        raise E

with open(CurDir + "/db.js", "w", encoding="utf-8") as FWriter:
    FWriter.write("TableTest = " + str(FinalResult).replace('"', '').replace("'", '"').replace("\n", ""))
