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
    def __init__(self, data):
        self.satir = data["satir"]
        self.numara = data["numara"]
        self.ad = data["ad"]

        self.biyoloji = get_lesson_from_data(data["biyoloji"])
        self.cografya = get_lesson_from_data(data["cografya"])
        self.fizik = get_lesson_from_data(data["fizik"])
        self.kimya = get_lesson_from_data(data["kimya"])
        self.matematik = get_lesson_from_data(data["matematik"])
        self.tarih = get_lesson_from_data(data["tarih"])
        self.edebiyat = get_lesson_from_data(data["edebiyat"])

    def __iter__(self):
        lesson_classes = [self.biyoloji.name, self.cografya.name, self.fizik.name, self.kimya.name,
                          self.matematik.name, self.tarih.name, self.edebiyat.name]
        lesson_names = ["biyoloji", "cografya", "fizik", "kimya", "matematik", "tarih", "edebiyat"]
        lessons = dict()

        for i, k in zip(lesson_names, lesson_classes):
            lessons[i] = k

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
