def get_lesson_from_data(data):
    lesson_name = ""
    if data is None or data == "" or data[-18:] == "(Kursa Yerleşmedi)":
        lesson_name = ""
        data = ""
    else:
        if data[-2] == " ":
            if data[-4].isdigit():
                lesson_name = data[-4] + data[-3] + data[-1]
            else:
                lesson_name = data[-3] + data[-1]
        elif data[-3] == " ":
            lesson_name = data[-5] + data[-4] + data[-2] + data[-1]
    lesson_teacher = data.split("\n")[0][8:]

    return Lesson(lesson_name, lesson_teacher)


class Student:
    def __init__(self, data: dict, data_dict: dict):
        self.satir = data["satir"]
        self.ad = data["ad"]
        self.numara = data["numara"]

        self.lesson_classes = list()
        self.lesson_names = list()

        for i in data_dict.keys():
            setattr(self, i, get_lesson_from_data(data[i]))
            self.lesson_classes.append(i + ".name")
            self.lesson_names.append(i)

    def __iter__(self):
        lessons = dict()
        for i in self.lesson_names:
            lessons[i] = getattr(self, i).name

        return iter(lessons.items())

    def lesson_from_name(self, name_string):
        if name_string == "cografya":
            return self.cografya
        if name_string == "biyoloji":
            return self.biyoloji
        if name_string == "fizik":
            return self.fizik
        if name_string == "kimya":
            return self.kimya
        if name_string == "matematik":
            return self.matematik
        if name_string == "tarih":
            return self.tarih
        if name_string == "edebiyat":
            return self.edebiyat


class Lesson:
    def __init__(self, name, teacher):
        self.name = name
        self.teacher = teacher
        pass
