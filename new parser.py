import pandas as pd
import re
import csv
import os
from datetime import datetime


def process_vacancies_csv(input_file, output_file='vacancies_for_teens_14plus.xlsx'):
    # === Шаг 1: Чтение и очистка файла ===
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    cleaned_lines = []
    for line in lines:
        if line.strip().startswith('"rvr",'):
            line = line[5:]  # убираем "rvr",
        cleaned_lines.append(line)

    temp_file = 'cleaned_temp.csv'
    with open(temp_file, 'w', encoding='utf-8', newline='') as f:
        f.writelines(cleaned_lines)

    # === Шаг 2: Задаём структуру колонок ===
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

    # Удаляем временный файл
    if os.path.exists(temp_file):
        os.remove(temp_file)

    # === Шаг 3: Извлечение года из даты создания ===
    def extract_year(date_str):
        if pd.isna(date_str) or date_str == 'nan' or not date_str:
            return ''
        try:
            # Формат: 3/1/2024 08:22:13 PM или 6/3/2025 01:29:01 AM
            date_obj = datetime.strptime(date_str.strip(), '%m/%d/%Y %I:%M:%S %p')
            return str(date_obj.year)
        except:
            try:
                # Альтернативный формат без времени
                date_obj = datetime.strptime(date_str.strip(), '%m/%d/%Y')
                return str(date_obj.year)
            except:
                return ''

    df['year'] = df['date_created'].apply(extract_year)

    # === Шаг 4: Фильтрация подростковых вакансий ===
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
    if not mask.any():
        print("⚠️ Не найдено вакансий по ключевым словам. Показаны все.")

    # === Шаг 5: Выбор нужных колонок ===
    selected_cols = [
        'year', 'vacancy_name_raw', 'employer_name', 'salary_min', 'salary_max', 'currency',
        'schedule_label', 'employment_label', 'region', 'city', 'address', 'duties',
        'requirements', 'url'
    ]
    available_cols = [col for col in selected_cols if col in df_filtered.columns]
    df_final = df_filtered[available_cols].copy()

    # Переименование
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

    # === Шаг 6: Очистка текста с обработкой NaN ===
    def clean_text(x):
        if pd.isna(x) or x == 'nan' or x == '':
            return ''
        if isinstance(x, str):
            x = re.sub(r'<[^>]+>', '', x)  # удаляем HTML-теги
            x = x.replace('&nbsp;', ' ')
            x = re.sub(r'\s+', ' ', x)  # сырая строка для \s+
            return x.strip()
        return str(x)

    for col in df_final.columns:
        df_final[col] = df_final[col].apply(clean_text)

    # === Шаг 7: Сохранение ===
    df_final.to_excel(output_file, index=False)
    df_final.to_csv(output_file.replace('.xlsx', '.csv'), index=False, encoding='utf-8-sig', sep=';')

    print(f"\n✅ Обработано {len(df_final)} вакансий.")
    print(f"Файл сохранён: {output_file}")

    # Предпросмотр
    preview_cols = ['Год', 'Вакансия', 'Работодатель', 'Город', 'Зарплата от', 'График работы']
    preview = df_final.head(10)[preview_cols]
    print("\n🔍 Предпросмотр (первые 10 вакансий):")
    print(preview.to_string(index=False))

    # Статистика по годам
    if 'Год' in df_final.columns and not df_final['Год'].empty:
        year_stats = df_final['Год'].value_counts().sort_index()
        print("\n📊 Распределение вакансий по годам:")
        for year, count in year_stats.items():
            if year:
                print(f"   {year}: {count} вакансий")


if __name__ == "__main__":
    process_vacancies_csv("Ярославль_вакансия_от_14лет.csv")