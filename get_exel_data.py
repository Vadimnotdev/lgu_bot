import openpyxl
import json
import re
import matplotlib.pyplot as plt

# Маппинг дней недели и ячеек для времени и пар
def load_mapping(file_name="mapping.json"):
    with open(file_name, "r", encoding="utf-8") as f:
        return json.load(f)

mapping = load_mapping()
DAYS_MAPPING = mapping["DAYS_MAPPING"]
VO_DAYS_MAPPING = mapping["VO_DAYS_MAPPING"]
GROUP_ROOM_MAPPING = mapping["GROUP_ROOM_MAPPING"]
def get_group_and_room_cells(sheet):
    """
    Возвращает список ячеек с группами (если заполнены) и колонку аудиторий.
    """
    group_cells = []
    
    # Проверяем, есть ли данные в C7
    if sheet["C7"].value:
        group_cells.append("C7")

    # Проверяем, есть ли данные в E7
    if sheet["E7"].value:
        group_cells.append("E7")

    # Если нет данных в обеих ячейках, используем только C7
    if not group_cells:
        group_cells = ["C7"]

    return {"group_cells": group_cells, "room_column": "D"}



def get_group_data(sheet, group_name):
    """
    Найти группу на листе и вернуть столбец для расписания.
    """
    group_cell_mapping = {"C7": "C", "E7": "E"}  # Группы на листе
    for cell, column in group_cell_mapping.items():
        cell_value = sheet[cell].value
        if cell_value and group_name in cell_value:
            return column
    return None

def determine_vo_sheet_mapping(sheet_name):
    normalized_sheet_name = sheet_name.strip().lower()

    if "лист1" in normalized_sheet_name:
        return VO_DAYS_MAPPING["лист1"]
    elif "лист2" in normalized_sheet_name:
        return VO_DAYS_MAPPING["лист2"]
    elif "лист3" in normalized_sheet_name:
        return VO_DAYS_MAPPING["лист3"]
    elif "лист4" in normalized_sheet_name:
        return VO_DAYS_MAPPING["лист4"]

    # Обработка листов с цифровыми названиями
    if re.match(r"^\d+-\d+", normalized_sheet_name):
        return VO_DAYS_MAPPING.get("лист1")
    return None


def normalize_teacher_name(teacher_name):
    if teacher_name:
        teacher_name = re.sub(r"(преп\.|ст\.преп\.|куратор|к\.э\.н\.|к\.п\.н\.|доцент)\s*", "", teacher_name)
        teacher_name = re.sub(r"\s+", " ", teacher_name).strip().lower()
        return re.sub(r"\.\s+", ".", teacher_name)
    return None

def clean_text(value):
    if value:
        return " ".join(str(value).split())
    return None

def display_schedule(schedule, entity):
    if not schedule:
        return f"Расписание для {entity} отсутствует."

    result = []
    for entry in schedule:
        pair_info = (
            f"📅 День недели: {entry.get('day', 'Не указано')}\n"
            f"📅 Дата: {entry['date']}\n"
            f"№ Пары: {entry['pair']}\n"
            f"🕒 Время: {entry['time'][0]} - {entry['time'][1]}\n"
            f"📖 Предмет: {entry['subject']}\n"
            f"🏫 Аудитория: {entry['room']}"
        )
        if 'teacher' in entry:
            pair_info += f"\n👩‍🏫 Преподаватель: {entry['teacher']}"
        if 'group' in entry:
            pair_info += f"\n👨‍🎓 Группа: {entry['group']}"

        # Добавляем пару в результат
        result.append(pair_info)

    return "\n-----------------------------------\n".join(result)



def search_teacher(sheet, sheet_name, teacher_name, target_day=None, education_type="SPO"):
    """
    Ищет расписание преподавателя по всем дням недели.
    Если target_day передан, то ищет только для этого дня.
    """
    teacher_name = normalize_teacher_name(teacher_name)
    result = []

    # Получение маппинга для дней недели
    if education_type == "SPO":
        days_mapping = DAYS_MAPPING
    else:
        vo_mapping = determine_vo_sheet_mapping(sheet_name)
        if vo_mapping is None:
            return []
        days_mapping = vo_mapping

    # Получаем все возможные столбцы групп (например, ["C7", "E7"])
    group_cells = get_group_and_room_cells(sheet)["group_cells"]
    columns = [cell[0] for cell in group_cells]  # Преобразуем ["C7", "E7"] в ["C", "E"]

    for day in days_mapping.keys():
        if target_day and target_day.lower() != day.lower():
            continue

        day_info = days_mapping[day]
        pairs_cells = day_info["pairs_cells"]
        time_cells = day_info["time_cells"]
        date_cell = day_info["date_cell"]

        # Извлекаем дату
        date = sheet[date_cell].value

        for column in columns:  # Перебираем "C" и "E"
            room_column = "D" if column == "C" else "F"

            for pair_cell, (time_start, time_end) in zip(pairs_cells, time_cells):
                pair_number = clean_text(sheet[f"A{pair_cell[1:]}"].value)
                subject = clean_text(sheet[f"{column}{pair_cell[1:]}"].value)
                teacher = clean_text(sheet[f"{column}{int(pair_cell[1:]) + 1}"].value)
                room = clean_text(sheet[f"{room_column}{pair_cell[1:]}"].value)

                # Проверяем совпадение имени преподавателя
                if teacher and teacher_name in normalize_teacher_name(teacher):
                    group_cell = f"{column}7"  # Определяем ячейку с группой (C7 или E7)
                    group_name = clean_text(sheet[group_cell].value)  # Получаем название группы

                    result.append({
                        "day": day.capitalize(),
                        "date": date or "Не указана",
                        "time": (sheet[time_start].value, sheet[time_end].value),
                        "subject": subject or "Не указано",
                        "room": room or "Не указана",
                        "group": group_name or "Не указана", 
                        "pair": pair_number or "Не указано"
                    })

    return result

def extract_schedule(sheet, column, day, education_type, sheet_name):
    schedule = []

    if education_type == "SPO":
        day_info = DAYS_MAPPING[day.lower()]
    else:
        vo_mapping = determine_vo_sheet_mapping(sheet_name)
        if not vo_mapping:
            return []
        day_info = vo_mapping[day.lower()]

    pairs_cells = day_info["pairs_cells"]
    time_cells = day_info["time_cells"]
    room_column = "D" if column == "C" else "F"
    date_cell = day_info["date_cell"]
    date = sheet[date_cell].value

    for pair_cell, (time_start, time_end) in zip(pairs_cells, time_cells):
        pair_number = clean_text(sheet[f"A{pair_cell[1:]}"].value)
        subject = clean_text(sheet[f"{column}{pair_cell[1:]}"].value)
        teacher = clean_text(sheet[f"{column}{int(pair_cell[1:]) + 1}"].value)
        room = clean_text(sheet[f"{room_column}{pair_cell[1:]}"].value)

        if subject or teacher or room:
            schedule.append({
                "day": day.capitalize(),
                "pair": pair_number or "Не указано",
                "time": (sheet[time_start].value, sheet[time_end].value),
                "subject": subject or "Не указано",
                "teacher": teacher or "Не указан",
                "room": room or "Не указана",
                "date": date or "Не указана"
            })
    return schedule

# Порядок дней недели для сортировки
DAYS_ORDER = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота"]

def generate_week_schedule_image(result, teacher_name):
    """
    Генерирует PNG-изображение расписания на всю неделю для преподавателя с разделением по дням и указанием даты.
    """
    # Сортируем расписание по дням недели и номерам пар
    result.sort(key=lambda entry: (
        DAYS_ORDER.index(entry.get('day', '').lower()) if entry.get('day', '').lower() in DAYS_ORDER else len(DAYS_ORDER),
        int(entry.get('pair', '0')) if entry.get('pair', '0').isdigit() else 0
    ))

    # Разбиваем расписание по дням недели
    day_groups = {day: [] for day in DAYS_ORDER}
    for entry in result:
        day = entry.get('day', 'Не указано').lower()
        if day in day_groups:
            day_groups[day].append(entry)

    # Создаём изображение
    fig, ax = plt.subplots(figsize=(8, 12))
    ax.set_title(f"\nРасписание для преподавателя {teacher_name}", fontsize=18, fontweight='bold')
    ax.axis("off")

    # Генерация текста расписания
    schedule_text = ""
    for day in DAYS_ORDER:
        if day_groups[day]:
            # Получаем дату для текущего дня
            date = day_groups[day][0].get('date', 'Не указана')
            schedule_text += f"\n-- {day.capitalize()} ({date}) --\n\n"  # Добавляем дату рядом с днём недели
            for entry in day_groups[day]:
                schedule_text += (
                    f"  Пара № {entry['pair']} | {entry['time'][0]} - {entry['time'][1]} | {entry['subject']}\n"
                    f"  Аудитория: {entry['room']}\n"
                    f"  Группа: {entry['group']}\n\n"
                )

    # Добавляем текст на изображение
    ax.text(0.1, 0.9, schedule_text, fontsize=10, va="top", ha="left", wrap=True)

    # Сохраняем изображение
    file_name = f"schedule_{teacher_name}.png"
    plt.savefig(file_name, bbox_inches="tight")
    plt.close(fig)

    return file_name


