import pandas as pd
import re
import csv
import os
import hashlib
from datetime import datetime


def extract_city(employer_name, address, default_city="Ярославль"):
    """
    Извлекает конкретный город из названия организации или адреса.
    """
    cities = [
        "Переславль-Залесский", "Переславль", "Ростов Великий", "Ростов",
        "Рыбинск", "Углич", "Ярославль", "Семибратово", "Данилов",
        "Тутаев", "Гаврилов-Ям", "Любим", "Мышкин", "Пошехонье",
        "Пречистое", "Некоуз", "Большое Село"
    ]

    # Проверяем адрес
    if address and pd.notna(address) and str(address).strip():
        address_str = str(address).lower()
        for city in cities:
            if city.lower() in address_str:
                return city

    # Проверяем название организации
    if employer_name and pd.notna(employer_name) and str(employer_name).strip():
        employer_str = str(employer_name).lower()
        for city in cities:
            if city.lower() in employer_str:
                return city

    return default_city


def get_vacancy_hash(row):
    """
    Создает уникальный хеш вакансии на основе ключевых полей.
    """
    key = f"{row.get('vacancy_name_raw', '')}{row.get('employer_name', '')}{row.get('salary_min', '')}{row.get('date_created', '')}"
    return hashlib.md5(key.encode('utf-8')).hexdigest()


def process_vacancies_csv(input_file, output_file='vacancies_for_teens_14plus.xlsx'):
    # === Чтение и очистка файла ===
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    cleaned_lines = []
    for line in lines:
        if line.strip().startswith('"rvr ",'):
            line = line[5:]
        cleaned_lines.append(line)

    temp_file = 'cleaned_temp.csv'
    with open(temp_file, 'w', encoding='utf-8', newline='') as f:
        f.writelines(cleaned_lines)

    # === Структура колонок ===
    expected_columns = [
        "source", "vacancy_id", "url", "date_created", "date_updated", "is_hidden",
        "vacancy_name_raw", "vacancy_name_lower", "duties", "salary_min", "salary_max", "salary_avg",
        "currency", "is_gross", "experience_months", "employer_id", "employer_name", "region_code",
        "industry", "okved_sections", "okved_groups", "okved_codes", "okved_names", "languages",
        "education_level", "accept_kids", "driver_license", "schedule_label", "employment_label",
        "region", "federal_district", "city", "address", "role_id", "role_name",
        "skills", "hard_skills", "soft_skills", "salary_period", "role_id_dashboard", "role_name_dashboard"
    ]

    df = pd.read_csv(
        temp_file,
        header=None,
        names=expected_columns[:50],
        on_bad_lines='skip',
        quoting=csv.QUOTE_ALL,
        dtype=str
    )

    if os.path.exists(temp_file):
        os.remove(temp_file)

    # === Извлечение года из даты ===
    def extract_year(date_str):
        if pd.isna(date_str) or date_str == 'nan' or not date_str:
            return ''
        try:
            date_obj = datetime.strptime(date_str.strip(), '%m/%d/%Y %I:%M:%S %p')
            return str(date_obj.year)
        except:
            try:
                date_obj = datetime.strptime(date_str.strip(), '%m/%d/%Y')
                return str(date_obj.year)
            except:
                return ''

    df['year'] = df['date_created'].apply(extract_year)

    # === Фильтрация подростковых вакансий ===
    teen_keywords = [
        'несовершеннолетн', 'квота.*14', 'квота.*16', 'квота.*18', 'от 14 лет', 'от 16 лет',
        'подросток', 'трудоустройство несовершеннолетних', 'до 18 лет', 'возраст.*14',
        'возраст.*16', 'возраст.*18', '14+', '16+', 'для молодежи', 'начало карьеры',
        'без опыта', 'ученик', 'помощник', 'разнорабочий', 'подсобный рабочий',
        'возраст от 15', 'возраст от 16', 'возраст от 17', 'возраст от 18'
    ]
    pattern = '|'.join(teen_keywords)
    mask = (
            df['duties'].astype(str).str.contains(pattern, case=False, na=False) |
            df['vacancy_name_raw'].astype(str).str.contains(pattern, case=False, na=False)
    )
    if 'requirements' in df.columns:
        mask |= df['requirements'].astype(str).str.contains(pattern, case=False, na=False)

    df_filtered = df[mask].copy() if mask.any() else df.copy()

    # === Удаление дубликатов ===
    seen_hashes = set()
    unique_indices = []

    for idx, row in df_filtered.iterrows():
        v_hash = get_vacancy_hash(row)
        if v_hash not in seen_hashes:
            seen_hashes.add(v_hash)
            unique_indices.append(idx)

    df_filtered = df_filtered.loc[unique_indices].copy()

    # === Выбор нужных колонок ===
    selected_cols = [
        'year', 'vacancy_name_raw', 'employer_name', 'salary_min', 'salary_max', 'currency',
        'schedule_label', 'employment_label', 'region', 'city', 'address', 'duties',
        'requirements', 'url'
    ]
    available_cols = [col for col in selected_cols if col in df_filtered.columns]
    df_final = df_filtered[available_cols].copy()

    # === Исправление города ===
    def fix_city(row):
        current_city = row.get('city', 'Ярославль')
        return extract_city(row.get('employer_name', ''), row.get('address', ''), current_city)

    if 'city' in df_final.columns:
        df_final['city'] = df_final.apply(fix_city, axis=1)

    # === Переименование колонок ===
    df_final.rename(columns={
        'year': 'Год',
        'vacancy_name_raw': 'Вакансия',
        'employer_name': 'Работодатель',
        'salary_min': 'Зарплата от',
        'salary_max': 'Зарплата до',
        'currency': 'Валюта',
        'schedule_label': 'График работы',
        'employment_label': 'Тип занятости',
        'region': 'Регион',
        'city': 'Город',
        'address': 'Адрес',
        'duties': 'Обязанности',
        'requirements': 'Требования',
        'url': 'Ссылка'
    }, inplace=True)

    # === Очистка текста ===
    def clean_text(x):
        if pd.isna(x) or x == 'nan' or x == '':
            return ''
        if isinstance(x, str):
            x = re.sub(r'<[^>]+>', '', x)
            x = x.replace('&nbsp;', ' ')
            x = re.sub(r'\s+', ' ', x)
            return x.strip()
        return str(x)

    for col in df_final.columns:
        df_final[col] = df_final[col].apply(clean_text)

    # === Сохранение ===
    df_final.to_excel(output_file, index=False)
    df_final.to_csv(output_file.replace('.xlsx', '.csv'), index=False, encoding='utf-8-sig', sep=';')


if __name__ == "__main__":
    process_vacancies_csv("Ярославль_вакансия_от_14лет.csv")