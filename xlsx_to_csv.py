import os
import pandas as pd

excel_path = 'Test_Task_Analyst_Warehouse_Logistics.xlsx'
os.makedirs("csv_pages", exist_ok=True)


# 1. Конвертація Excel -> CSV
xls = pd.ExcelFile(excel_path)

for sheet_name in xls.sheet_names:
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    for col in df.select_dtypes(['object']).columns:
        df[col] = df[col].astype(str).str.strip()

    csv_filename = f"csv_pages/{sheet_name}.csv"
    df.to_csv(csv_filename, index=True, encoding='utf-8-sig')


# 2. Завантаження даних
goods_df = pd.read_csv(
    "csv_pages/Довідник товарів.csv",
    index_col=0,
    encoding='utf-8-sig'
)

shipments_df = pd.read_csv(
    "csv_pages/Реєстр відвантажень.csv",
    index_col=0,
    encoding='utf-8-sig'
)

# 3. Первинна перевірка
anomalies = []

# Перевірка пропусків
for name, data in [
    ("товари", goods_df),
    ("відвантаження", shipments_df)
]:
    for col in data.columns:
        missing = data[col].isna().sum()

        if missing > 0:
            anomalies.append({
                "джерело": name,
                "тип": "пропущені_значення",
                "стовпець": col,
                "кількість": missing,
                "опис": f"Виявлено {missing} пропущених значень"
            })


# Перевірка дублікатів
for name, data in [
    ("товари", goods_df),
    ("відвантаження", shipments_df)
]:
    duplicates = data.duplicated().sum()

    if duplicates > 0:
        anomalies.append({
            "джерело": name,
            "тип": "дублікати",
            "стовпець": "",
            "кількість": duplicates,
            "опис": f"Виявлено {duplicates} повних дублікатів"
        })


# 4. Перевірка довідника товарів
numeric_goods_cols = [
    'Норма упаковки (шт)',
    'Довжина (мм)',
    'Ширина (мм)',
    'Висота (мм)',
    'Вага за одиницю (кг)'
]

for col in numeric_goods_cols:
    if col in goods_df.columns:
        invalid = goods_df[col] <= 0

        if invalid.any():
            anomalies.append({
                "джерело": "товари",
                "тип": "некоректне_значення",
                "стовпець": col,
                "кількість": invalid.sum(),
                "опис": "Виявлено значення <= 0"
            })


# 5. Перевірка реєстру відвантажень
if 'Кількість (шт)' in shipments_df.columns:
    invalid_qty = shipments_df['Кількість (шт)'] <= 0

    if invalid_qty.any():
        anomalies.append({
            "джерело": "відвантаження",
            "тип": "некоректне_значення",
            "стовпець": "Кількість (шт)",
            "кількість": invalid_qty.sum(),
            "опис": "Кількість товару повинна бути > 0"
        })


# 6. Перевірка відповідності товарів
missing_products = shipments_df[
    ~shipments_df["ID Товару"].isin(goods_df["ID Товару"])
]

if len(missing_products) > 0:
    anomalies.append({
        "джерело": "відвантаження",
        "тип": "відсутній_товар",
        "стовпець": "ID Товару",
        "кількість": len(missing_products),
        "опис": "Товар відсутній у довіднику товарів"
    })


# 7. Об'єднання даних
df = pd.merge(
    shipments_df,
    goods_df,
    on="ID Товару",
    how="inner"
)


# 8. Розрахунок об'єму та ваги
df["Об'єм (м3)"] = (
    df['Довжина (мм)']
    * df['Ширина (мм)']
    * df['Висота (мм)']
    / 1e9
) * df['Кількість (шт)']

df['Вага (кг)'] = (
    df['Вага за одиницю (кг)']
    * df['Кількість (шт)']
)


# 9. Логічна перевірка об'єднаних даних
# Перевірка: кількість товару > 0,
# але фізичні характеристики товару некоректні
zero_characteristics = (
    (df['Кількість (шт)'] > 0)
    &
    (
        (df['Довжина (мм)'] <= 0)
        | (df['Ширина (мм)'] <= 0)
        | (df['Висота (мм)'] <= 0)
        | (df['Вага за одиницю (кг)'] <= 0)
    )
)

if zero_characteristics.any():
    anomalies.append({
        "джерело": "об'єднані_дані",
        "тип": "логічна_аномалія",
        "стовпець": "характеристики_товару",
        "кількість": zero_characteristics.sum(),
        "опис": (
            "Кількість товару > 0, але одна або декілька "
            "фізичних характеристик товару <= 0"
        )
    })


# Перевірка: кількість товару > 0,
# але розрахований об'єм або вага дорівнює 0
zero_volume_weight = (
    (df['Кількість (шт)'] > 0)
    &
    (
        (df["Об'єм (м3)"] <= 0)
        | (df["Вага (кг)"] <= 0)
    )
)

if zero_volume_weight.any():
    anomalies.append({
        "джерело": "об'єднані_дані",
        "тип": "логічна_аномалія",
        "стовпець": "об'єм_вага",
        "кількість": zero_volume_weight.sum(),
        "опис": (
            "За ненульової кількості товару розрахований "
            "об'єм або вага дорівнює 0"
        )
    })


# 10. Перевірка статистичних викидів
def check_iqr(data, column):
    q1 = data[column].quantile(0.25)
    q3 = data[column].quantile(0.75)

    iqr = q3 - q1

    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr

    outliers = (
        (data[column] < lower)
        | (data[column] > upper)
    )

    return outliers, lower, upper


for col in [
    'Кількість (шт)',
    "Об'єм (м3)",
    'Вага (кг)'
]:
    if col in df.columns:
        outliers, lower, upper = check_iqr(df, col)

        if outliers.any():
            anomalies.append({
                "джерело": "об'єднані_дані",
                "тип": "статистичний_викид",
                "стовпець": col,
                "кількість": outliers.sum(),
                "опис": (
                    f"IQR-викиди: допустимий діапазон "
                    f"від {lower:.3f} до {upper:.3f}"
                )
            })



# 11. Збереження звіту про аномалії
anomalies_df = pd.DataFrame(anomalies)

anomalies_df.to_excel(
    "csv_pages/anomalies_report.xlsx",
    index=False
)

print("\n=== ЗВІТ ПРО АНОМАЛІЇ ===")

if anomalies_df.empty:
    print("Аномальних даних не виявлено.")
else:
    print(anomalies_df.to_string(index=False))


# 12. Збереження підсумкового датасету
df = df[(df != 0).all(axis=1)]
df.to_excel("csv_pages/Warehouse_Logistics.xlsx", index=False)
