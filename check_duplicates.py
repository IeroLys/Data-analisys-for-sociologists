import pandas as pd
import hashlib


def generate_vacancy_hash(employer, position, salary, description):
    desc_fragment = str(description)[:100] if description else ''
    key = f"{str(employer).lower()}|{str(position).lower()}|{str(salary)}|{desc_fragment}"
    return hashlib.md5(key.encode('utf-8')).hexdigest()


# Читаем исходный файл
df = pd.read_csv('vacancies_for_teens_14plus.csv', sep=';', encoding='utf-8-sig')

# Создаём хеш
df['vacancy_hash'] = df.apply(
    lambda row: generate_vacancy_hash(
        row.get('Работодатель', ''),
        row.get('Вакансия', ''),
        row.get('Зарплата от', ''),
        row.get('Обязанности', '')
    ),
    axis=1
)

# Находим дубликаты
duplicates = df[df.duplicated(subset=['vacancy_hash'], keep=False)]

print(f"📊 Всего дубликатов: {len(duplicates)}")
print(f"📊 Уникальных вакансий: {len(df) - len(duplicates) + duplicates['vacancy_hash'].nunique()}")

# Показываем 5 групп дубликатов
hash_groups = duplicates['vacancy_hash'].unique()[:5]

for i, h in enumerate(hash_groups, 1):
    group = duplicates[duplicates['vacancy_hash'] == h]
    print(f"\n{'=' * 100}")
    print(f"🔴 Группа дубликатов #{i} ({len(group)} записей):")
    print(f"{'=' * 100}")

    for idx, row in group.iterrows():
        print(f"\n📌 Запись #{idx}:")
        print(f"   Вакансия: {str(row.get('Вакансия', ''))[:80]}")
        print(f"   Работодатель: {str(row.get('Работодатель', ''))[:60]}")
        print(f"   Город: {row.get('Город', '')}")
        print(f"   Зарплата: {row.get('Зарплата от', '')} - {row.get('Зарплата до', '')}")
        print(f"   Источник: {str(row.get('Ссылка', ''))[:80]}")

# Сохраняем все дубликаты
duplicates.to_csv('all_duplicates.csv', index=False, encoding='utf-8-sig', sep=';')
print(f"\n💾 Все дубликаты сохранены в: all_duplicates.csv")