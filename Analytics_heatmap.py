import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Папки с сезонами
season_folders = [
    r'C:\Users\JuliaEfimova\Desktop\PET\Субтитры к фильму\extracted_srt_3\Silicon_Valley - season 1.en'
]

# Названия сезонов
season_names = ["Season 1"]

def count_categories_per_season(season_folders):
    """
    Подсчитывает категории для каждого сезона.
    """
    all_seasons_data = {}

    for season_folder, season_name in zip(season_folders, season_names):
        excel_path = os.path.join(season_folder, f"{os.path.basename(season_folder)}_notes.xlsx")
        try:
            df = pd.read_excel(excel_path, sheet_name="Sheet2")
            category_counts = df['Category'].value_counts()
            all_seasons_data[season_name] = category_counts
        except Exception as e:
            print(f"Ошибка чтения файла {excel_path}: {e}")

    return pd.DataFrame(all_seasons_data).fillna(0).astype(int)

def save_category_data_to_excel(stats_df, output_path):
    """
    Сохраняет полную таблицу категорий в Excel.
    """
    stats_df.to_excel(output_path, engine='openpyxl')
    print(f"Все категории сохранены в Excel: {output_path}")

def plot_heatmap(stats_df, output_path):
    """
    Строит тепловую карту категорий по сезонам.
    """
    plt.figure(figsize=(12, 8))
    sns.heatmap(stats_df, annot=True, fmt="d", cmap="YlGnBu", cbar=True)
    plt.title("Category Distribution Heatmap", fontsize=16)
    plt.xlabel("Season", fontsize=12)
    plt.ylabel("Category", fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"Тепловая карта сохранена: {output_path}")

if __name__ == "__main__":
    # Подсчет категорий
    stats_df = count_categories_per_season(season_folders)

    # Сохранение всех категорий в Excel
    full_excel_path = r'C:\Users\JuliaEfimova\Desktop\PET\Субтитры к фильму\all_categories-2.xlsx'
    save_category_data_to_excel(stats_df, full_excel_path)

    # Построение тепловой карты
    heatmap_path = r'C:\Users\JuliaEfimova\Desktop\PET\Субтитры к фильму\category_heatmap-2.png'
    plot_heatmap(stats_df, heatmap_path)
