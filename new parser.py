import pandas as pd
import re
import csv
import os
from datetime import datetime
import hashlib

# === Словари для определения реального города ===
# === Словари для определения реального города ===
CITY_KEYWORDS = {
    'Рыбинск': ['рыбинск', 'рыбинский', 'рыбинская', 'рыбинское'],
    'Ростов': ['ростов', 'ростовский', 'ростовская', 'ростовское', 'ростов великий'],
    'Переславль-Залесский': ['переславль', 'переславский', 'переславская', 'переславское'],
    'Углич': ['углич', 'угличский', 'угличская', 'угличское'],
    'Тутаев': ['тутаев', 'тутаевский', 'тутаевская', 'тутаевское'],
    'Гаврилов-Ям': ['гаврилов-ям', 'гаврилов-ямский', 'гаврилов-ямская'],
    'Данилов': ['данилов', 'даниловский', 'даниловская', 'даниловское'],
    'Семибратово': ['семибратово', 'семибратовский', 'семибратовская', 'семибратовское'],
    'Некоуз': ['некоуз', 'некоузский', 'некоузская', 'некоузское'],
    'Ярославль': ['ярославль', 'ярославский', 'ярославская', 'ярославское'],
    'Брейтово': ['брейтово', 'брейтовский', 'брейтовская'],
    'Мышкин': ['мышкин', 'мышкинский', 'мышкинская'],
    'Пошехонье': ['пошехонье', 'пошехонский', 'пошехонская'],
    'Любим': ['любим', 'любимский', 'любимская'],
    'Пречистое': ['пречистое', 'пречистенский', 'пречистенская'],
    'Борисоглебский': ['борисоглебский', 'борисоглебская'],
    'Угодичи': ['угодичи', 'угодичская'],
    'Петровское': ['петровское', 'петровский', 'петровская'],
    'Марково': ['марково', 'марковский', 'марковская'],
    'Шурма': ['шурма', 'шуркольская', 'шурскольская'],
    'Скнятино': ['скнятино', 'скнятинская'],
    'Вахрушево': ['вахрушево', 'вахрушевская'],
    'Покров-Ситское': ['покрово-ситская', 'покров-ситская'],
    'Дубки': ['дубки', 'дубковская'],
    'Красный Профинтерн': ['красный профинтерн'],
}


def process_vacancies_csv(input_file, output_file='vacancies_for_teens_ver2.xlsx'):
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

    # Читаем CSV с более гибкими настройками
    df = pd.read_csv(
        temp_file,
        header=None,
        names=expected_columns,
        on_bad_lines='warn',  # Вместо 'skip' — предупреждать о проблемах
        quoting=csv.QUOTE_MINIMAL,  # Изменено с QUOTE_ALL
        dtype=str,
        skipinitialspace=True,
        engine='python'  # Более гибкий парсер
    )

    # Проверяем количество колонок
    print(f"📊 Ожидаемо колонок: {len(expected_columns)}")
    print(f"📊 Фактически колонок: {len(df.columns)}")

    # Если колонок больше — обрезаем
    if len(df.columns) > len(expected_columns):
        df = df.iloc[:, :len(expected_columns)]
    # Если колонок меньше — добавляем пустые
    elif len(df.columns) < len(expected_columns):
        for i in range(len(df.columns), len(expected_columns)):
            df[i] = ''
        df.columns = expected_columns

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

    # === Шаг 3.5: Функция определения реального города ===
    def extract_actual_city(employer_name, city_field):
        """Извлекает реальный город из названия работодателя"""
        employer_lower = str(employer_name).lower() if employer_name else ''

        # Проверяем название работодателя на наличие топонимов
        for city, keywords in CITY_KEYWORDS.items():
            for keyword in keywords:
                if keyword in employer_lower:
                    return city

        # Если в поле city_field указан конкретный город (не Ярославль)
        if city_field and str(city_field).strip() and str(city_field).strip() != 'Ярославль':
            return str(city_field).strip()

        # По умолчанию оставляем как есть
        return city_field if city_field else 'Ярославль'

    # === Шаг 3.6: Функция для обнаружения дубликатов ===
    def generate_vacancy_hash(employer, position, salary, description):
        """Генерирует хеш для идентификации уникальной вакансии"""
        # Берём первые 100 символов описания для сравнения
        desc_fragment = str(description)[:100] if description else ''
        key = f"{str(employer).lower()}|{str(position).lower()}|{str(salary)}|{desc_fragment}"
        return hashlib.md5(key.encode('utf-8')).hexdigest()

    # === Шаг 3.7: Функция извлечения адреса из текста ===
    def extract_address_from_text(duties_text, city_field):
        """Извлекает адрес из текста обязанностей"""
        if pd.isna(duties_text) or not duties_text:
            return city_field if city_field else ''

        text = str(duties_text)

        # Паттерны для поиска адреса
        patterns = [
            r'г\.\s*([А-Яа-яЁё\-]+),\s*([^.;]+?)(?:\.|$)',  # г. Ярославль, улица...
            r'адрес:\s*([^.;]+?)(?:\.|$)',  # адрес: ...
            r'по адресу:\s*([^.;]+?)(?:\.|$)',  # по адресу: ...
            r'место работы:\s*([^.;]+?)(?:\.|$)',  # место работы: ...
            r'ул\.\s*([А-Яа-яЁё\-]+\s*\d+[А-Яа-я]?)',  # ул. ... д.15
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                address = match.group(0).strip()
                if len(address) > 5 and len(address) < 100:  # Фильтр по длине
                    return address

        return city_field if city_field else ''

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

    # === ПЕРЕИМЕНОВАНИЕ (сначала переименуем!) ===
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

    # === Исправление городов (теперь русские имена работают) ===
    print("🔍 Исправляем названия городов...")
    df_final['Город'] = df_final.apply(
        lambda row: extract_actual_city(row.get('Работодатель', ''), row.get('Город', '')),
        axis=1
    )

    # === Удаление дубликатов (теперь русские имена работают) ===
    print("🔄 Удаляем дубликаты вакансий...")
    df_final['vacancy_hash'] = df_final.apply(
        lambda row: generate_vacancy_hash(
            row.get('Работодатель', ''),
            row.get('Вакансия', ''),
            row.get('Зарплата от', ''),
            row.get('Обязанности', '')
        ),
        axis=1
    )

    total_before = len(df_final)
    df_final = df_final.drop_duplicates(subset=['vacancy_hash'], keep='first')
    duplicates_removed = total_before - len(df_final)
    df_final = df_final.drop(columns=['vacancy_hash'])
    print(f"   Удалено дубликатов: {duplicates_removed}")

    # === Извлечение адресов из текста ===
    print("🏠 Извлекаем адреса из текста...")
    df_final['Адрес'] = df_final.apply(
        lambda row: extract_address_from_text(row.get('Обязанности', ''), row.get('Город', '')),
        axis=1
    )

    # === Шаг 6: Очистка текста с обработкой NaN ===
    def clean_text(x):
        # Безопасная проверка на пустое значение
        try:
            if x is None:
                return ''
            if isinstance(x, float) and pd.isna(x):
                return ''
            x_str = str(x).strip()
            if x_str == '' or x_str.lower() == 'nan':
                return ''
        except:
            return ''

        if isinstance(x, str):
            x = re.sub(r'<[^>]+>', '', x)  # удаляем HTML-теги
            x = x.replace('&nbsp;', ' ')
            x = re.sub(r'\s+', ' ', x)
            return x.strip()
        return str(x).strip()

    for col in df_final.columns:
        df_final[col] = df_final[col].apply(clean_text)

    # === Шаг 7: Сохранение ===
    df_final.to_excel(output_file, index=False)
    df_final.to_csv(output_file.replace('.xlsx', '.csv'), index=False, encoding='utf-8-sig', sep=';')

    print(f"\n✅ Обработано {len(df_final)} вакансий.")
    # Статистика по городам
    if 'Город' in df_final.columns:
        city_stats = df_final['Город'].value_counts()
        print("\n📊 Распределение вакансий по городам:")
        for city, count in city_stats.items():
            print(f"   {city}: {count} вакансий")
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