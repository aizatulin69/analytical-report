import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

df = pd.read_csv("csv_pages/Warehouse_Logistics.csv")

# Згортаємо дані до рівня окремого замовлення
orders_agg = df.groupby(['Номер замовлення', 'Клієнт', 'Дата відвантаження']).agg(
    Кількість_рядків=('ID Товару', 'count'),
    Загальна_кількість_шт=('Кількість (шт)', 'sum'),
    Загальний_обєм_м3=("Об'єм (м3)", 'sum'),
    Загальна_вага_кг=('Вага (кг)', 'sum')
).reset_index()

# Кластеризація замовлень за допомогою KMeans з scikit-learn
features = ['Кількість_рядків', 'Загальна_кількість_шт', 'Загальний_обєм_м3', 'Загальна_вага_кг']
X = orders_agg[features].values

# Стандартизація даних
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# Навчання моделі (k = 3 кластери)
kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
orders_agg['Код_кластера'] = kmeans.fit_predict(X_scaled)

# Маппінг зрозумілих назв кластерів
cluster_names = {0: 'K1', 1: 'K2', 2: 'K3'}
orders_agg['Тип замовлення (Кластер)'] = orders_agg['Код_кластера'].map(cluster_names)

# Аналіз закономірностей та збереження
correlation_matrix = orders_agg[features].corr()

cluster_summary = orders_agg.groupby('Тип замовлення (Кластер)').agg(
    Кількість_замовлень=('Номер замовлення', 'count'),
    Сер_рядків=('Кількість_рядків', 'mean'),
    Сер_шт=('Загальна_кількість_шт', 'mean'),
    Сер_обєм_м3=('Загальний_обєм_м3', 'mean'),
    Сер_вага_кг=('Загальна_вага_кг', 'mean')
).reset_index()

# Збереження у новий Excel-файл
output_filename = 'Аналіз на кластери.xlsx'
with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    orders_agg.to_excel(writer, sheet_name='Кластеризація замовлень', index=False)
    cluster_summary.to_excel(writer, sheet_name='Профілі кластерів', index=False)
    correlation_matrix.to_excel(writer, sheet_name='Матриця кореляцій')

print(f"Аналіз успішно завершено! Результати збережено у файл '{output_filename}'.")
print("\n--- Профілі виділених кластерів ---")
print(orders_agg.groupby('Тип замовлення (Кластер)')[features].mean())