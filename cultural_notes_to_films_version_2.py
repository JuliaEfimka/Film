import os
import zipfile
import pandas as pd
from tqdm import tqdm
from pysrt import SubRipFile
import tiktoken
import matplotlib.pyplot as plt
import openai

# Установите ваш API-ключ OpenAI
openai.api_key = 'API-KEY'

# Категории для анализа
categories = ["set expression", "pun", "joke", "location", "venue", "event", "drug", "famous team", "food", "drink", "corporate humor", "science", "movie", "song", "book"]

# 1. Функция для распаковки zip-архивов
def extract_zip_files(zip_folder, output_folder):
    """
    Распаковывает только ZIP-файлы, относящиеся к Emily in Paris.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(zip_folder):
        if filename.endswith('.zip') and "Silicon_Valley" in filename:
            zip_path = os.path.join(zip_folder, filename)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                extract_to = os.path.join(output_folder, os.path.splitext(filename)[0])
                zip_ref.extractall(extract_to)
                print(f"Распакован: {filename} -> {extract_to}")

# 2. Функция для подсчета токенов
def count_tokens(text):
    """
    Подсчет количества токенов в тексте.
    """
    encoding = tiktoken.encoding_for_model("gpt-4")
    return len(encoding.encode(text))

# 3. Функция для генерации культурных заметок
def generate_cultural_note_and_category(subtitle_line):
    """
    Генерирует культурную заметку и её категорию для строки субтитров.
    """
    if not subtitle_line.strip():
        return "", ""

    prompt_note = (
        f"Extract cultural references (set expression, pun, joke, location, venue, event, drug, famous team, food, drink, corporate humor, science, movie, song, book) from the given text. "
        f"Focus only on less obvious or culturally specific elements that require detailed local knowledge. Avoid trivial explanations or generalizations.\n\n"
        f"For each, provide a concise explanation in Russian, starting directly with the explanation without repeating the source text. "
        f"Avoid verbosity and do not explain common or widely understood expressions. If no explanation is needed, return '-'.\n\n"
        f"Text: '{subtitle_line}'"
    )
    try:
        response_note = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt_note}],
            max_tokens=120,
            temperature=0
        )
        cultural_note = response_note.choices[0].message.content.strip()
        cultural_note = cultural_note.lstrip('-– ').strip()  # Убираем дефисы и пробелы в начале
        if not cultural_note or cultural_note == "-":
            return "", ""

        prompt_category = (
            f"Return only one word representing the category from these options: {', '.join(categories)}. "
            f"No explanations, only the category name.\n\n"
            f"Cultural Note: '{cultural_note}'"
        )
        response_category = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt_category}],
            max_tokens=20,
            temperature=0
        )
        category = response_category.choices[0].message.content.strip()
        return cultural_note, category
    except Exception as e:
        return f"Error: {e}", "Error"

# 4. Функция для обработки SRT-файлов
def process_srt(input_srt_path, percentage=100):
    """
    Обрабатывает SRT файл, извлекая культурные заметки и их категории.
    """
    try:
        subs = SubRipFile.open(input_srt_path, encoding='utf-8')
    except Exception as e:
        print(f"Error reading SRT file: {e}")
        return [], 0

    subtitle_lines = [sub.text.strip() for sub in subs if sub.text.strip()]
    total_items = len(subtitle_lines)
    limit = max(1, int(total_items * percentage / 100))
    subtitle_lines = subtitle_lines[:limit]

    # Подсчет токенов для оценки стоимости
    total_tokens = sum(count_tokens(line) for line in subtitle_lines)
    print(f"Подсчитано {total_tokens} токенов для файла: {input_srt_path}")

    cultural_notes = []
    with tqdm(total=len(subtitle_lines), desc="Processing subtitles") as pbar:
        for line in subtitle_lines:
            try:
                cultural_note, category = generate_cultural_note_and_category(line)
            except Exception as e:
                cultural_note, category = f"Error: {e}", "Error"
            cultural_notes.append({"Subtitle": line, "Cultural Note": cultural_note, "Category": category})
            pbar.update(1)

    return cultural_notes, total_tokens

# 5. Обработка сезона
def process_season(season_folder, output_excel_path, percentage=1):
    """
    Обрабатывает все серии в папке сезона и объединяет данные в один Excel.
    Также считает общую стоимость обработки сезона.
    """
    all_cultural_notes = []
    total_tokens_season = 0

    for filename in os.listdir(season_folder):
        if filename.endswith('.srt'):
            input_srt_path = os.path.join(season_folder, filename)
            notes, total_tokens = process_srt(input_srt_path, percentage)
            if notes:
                all_cultural_notes.extend(notes)
                total_tokens_season += total_tokens

    # Подсчет стоимости сезона
    cost = (total_tokens_season * 2) / 1000000 * 1.25
    print(f"Общее количество токенов для сезона {os.path.basename(season_folder)}: {total_tokens_season}")
    print(f"Итоговая стоимость обработки сезона (точно): ${cost:.6f}")

    # Создание общего Excel
    df = pd.DataFrame(all_cultural_notes)
    df.to_excel(output_excel_path, index=False, engine='openpyxl')
    print(f"Season file saved to {output_excel_path}")

    # Построение графика
    plot_cultural_categories(df, output_excel_path)

# 6. Функция для построения графика распределения категорий
def plot_cultural_categories(df, output_excel_path):
    """
    Создаёт график распределения категорий культурных заметок.
    """
    category_counts = df[df['Category'] != ""]['Category'].value_counts()
    plt.figure(figsize=(10, 6))
    bars = plt.bar(category_counts.index, category_counts.values, width=0.5)

    # Добавляем числа на столбцы
    for bar in bars:
        plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), int(bar.get_height()),
                 ha='center', va='bottom', fontsize=10)

    plt.title('Distribution of Cultural Note Categories')
    plt.xlabel('Category')
    plt.ylabel('Count')
    plt.xticks(rotation=45, ha='center', fontsize=10)
    plt.tight_layout()
    plt.savefig(output_excel_path.replace('.xlsx', '.png'))
    plt.close()

# Основной скрипт
if __name__ == "__main__":
    zip_folder = r'C:\Users\JuliaEfimova\Desktop\PET\Субтитры к фильму'
    output_folder = r'C:\Users\JuliaEfimova\Desktop\PET\Субтитры к фильму\extracted_srt_3'

    # Распаковка архивов
    extract_zip_files(zip_folder, output_folder)

    # Обработка сезонов
    season_folders = [
        os.path.join(output_folder, "Silicon_Valley - season 1.en")
    ]

    for season_folder in season_folders:
        output_excel_path = os.path.join(season_folder, f"{os.path.basename(season_folder)}_notes.xlsx")
        process_season(season_folder, output_excel_path, percentage=100)
