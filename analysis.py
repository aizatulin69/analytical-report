import pandas as pd

goods_df = pd.read_csv("csv_pages/Довідник товарів.csv", index_col=0, encoding='utf-8-sig')
shipments_df = pd.read_csv("csv_pages/Реєстр відвантажень.csv", index_col=0, encoding='utf-8-sig')

# З'єднуємо таблиці
df = pd.read_csv("csv_pages/Warehouse_Logistics.xlsx", index_col=0, encoding='utf-8-sig')

# Розраховуємо об'єм (м3) та вагу (кг) для кожного рядка
# Габарити в довіднику вказані в мм, тому ділимо на 1,000,000,000 для переведення в м³
df['Об\'єм (м3)'] = (df['Довжина (мм)'] * df['Ширина (мм)'] * df['Висота (мм)'] / 1e9) * df['Кількість (шт)']
df['Вага (кг)'] = df['Вага за одиницю (кг)'] * df['Кількість (шт)']


# БЛОК 1: Щоденні показники та виявлення пікових днів
daily = df.groupby('Дата відвантаження').agg(
    Рядків=('Номер замовлення', 'count'),
    Документів=('Номер замовлення', 'nunique'),
    Обєм_м3=('Об\'єм (м3)', 'sum'),
    Вага_кг=('Вага (кг)', 'sum'),
    Кількість_шт=('Кількість (шт)', 'sum')
).reset_index()

# БЛОК 2: Структура замовлень («Типове відвантаження»)
# Згортаємо дані до рівня окремого замовлення (документа)
orders = df.groupby(['Дата відвантаження', 'Номер замовлення', 'Клієнт']).agg(
    Рядків=('ID Товару', 'count'),
    Загалом_шт=('Кількість (шт)', 'sum'),
    Обєм_м3=('Об\'єм (м3)', 'sum'),
    Вага_кг=('Вага (кг)', 'sum')
).reset_index()
summary_stats = pd.DataFrame({
    'Медіана (Типове)': orders[['Рядків', 'Загалом_шт', 'Обєм_м3', 'Вага_кг']].median(),
    'Середнє': orders[['Рядків', 'Загалом_шт', 'Обєм_м3', 'Вага_кг']].mean(),
    'Максимум': orders[['Рядків', 'Загалом_шт', 'Обєм_м3', 'Вага_кг']].max()
})


# БЛОК 3: ТОП 10 товарів та ТОП 5 клієнтів
top_products = df.groupby(['ID Товару', 'Товар', 'Категорія']).agg(
    Кількість_звернень=('Номер замовлення', 'count'),
    Загальна_кількість_шт=('Кількість (шт)', 'sum')
).reset_index().sort_values(by='Кількість_звернень', ascending=False).head(10)

top_clients = df.groupby('Клієнт').agg(
    Замовлень=('Номер замовлення', 'nunique'),
    Загальний_обєм_м3=('Об\'єм (м3)', 'sum'),
    Загальна_вага_кг=('Вага (кг)', 'sum')
).reset_index().sort_values(by='Загальний_обєм_м3', ascending=False).head(5)

with pd.ExcelWriter("Звіт.xlsx", engine='openpyxl') as writer:
    summary_stats.to_excel(writer, sheet_name="Типове відвантаження", index=False)
    daily.to_excel(writer, sheet_name='Щоденна звітність', index=False)
    top_products.to_excel(writer, sheet_name='ТОП 10 Товарів', index=False)
    top_clients.to_excel(writer, sheet_name='ТОП 5 Клієнтів', index=False)