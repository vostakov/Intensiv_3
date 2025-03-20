import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from catboost import CatBoostRegressor
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

train_data = pd.DataFrame()
future_df = pd.DataFrame()

train_data = pd.read_excel('C:/Users/User/Documents/GitHub/Intensiv_3/Dataset/data/train.xlsx')
train_data['dt'] = pd.to_datetime(train_data['dt'])
train_data['День'] = train_data['dt'].dt.day
train_data['Месяц'] = train_data['dt'].dt.month
train_data['Год'] = train_data['dt'].dt.year
train_data['ДеньНедели'] = train_data['dt'].dt.dayofweek
train_data['ЦенаL1'] = train_data['Цена на арматуру'].shift(1)
train_data['ЦенаL2'] = train_data['Цена на арматуру'].shift(2)
train_data.dropna(inplace=True)

def train_and_predict():
    global future_df
    try:
        weeks = int(week_choice.get())

        if weeks < 1 or weeks > 6:
            messagebox.showwarning("Warning", "Количество недель должно быть от 1 до 6!")
            return

        X_train = train_data[['День', 'Месяц', 'Год', 'ДеньНедели', 'ЦенаL1', 'ЦенаL2']]
        y_train = train_data['Цена на арматуру']
        model = CatBoostRegressor(iterations=500, learning_rate=0.1, depth=6, silent=True)
        model.fit(X_train, y_train)

        future_dates = pd.date_range(start=train_data['dt'].max() + pd.Timedelta(days=1), periods=weeks, freq='W-MON')
        
        future_data = pd.DataFrame({
            'День': future_dates.day,
            'Месяц': future_dates.month,
            'Год': future_dates.year,
            'ДеньНедели': future_dates.dayofweek,
        })
        
        last_price = train_data['Цена на арматуру'].iloc[-1]
        future_data['ЦенаL1'] = last_price
        future_data['ЦенаL2'] = last_price
        
        future_predictions = model.predict(future_data)
        future_df = pd.DataFrame({'dt': future_dates, 'Цена на арматуру': future_predictions})

        plot_graph()
        display_predictions()

    except Exception as e:
        messagebox.showerror("Error", f"Ошибка при прогнозировании: {e}")

def plot_graph():
    global future_df
    fig, ax = plt.subplots(figsize=(10, 5))
    
    ax.plot(train_data['dt'], train_data['Цена на арматуру'], label='Исторические данные', color='blue')
    ax.plot(future_df['dt'], future_df['Цена на арматуру'], label='Предсказанные цены', color='red', linestyle='-', linewidth=2)
    ax.set_title('Прогноз цен на арматуру')
    ax.set_xlabel('Дата')
    ax.set_ylabel('Цена на арматуру')
    ax.legend()
    ax.grid()

    for widget in graph_frame.winfo_children():
        widget.destroy()

    canvas = FigureCanvasTkAgg(fig, graph_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

def display_predictions():
    global future_df
    for widget in table_frame.winfo_children():
        widget.destroy()

    table = ttk.Treeview(table_frame, columns=("Дата", "Цена на арматуру"), show="headings", height=10)
    table.heading("Дата", text="Дата")
    table.heading("Цена на арматуру", text="Цена на арматуру")
    table.grid(row=0, column=0, padx=10, pady=10)

    for _, row in future_df.iterrows():
        table.insert("", "end", values=(row["dt"].strftime('%Y-%m-%d'), row["Цена на арматуру"]))

    export_button = tk.Button(table_frame, text="Экспортировать в Excel", command=export_to_excel)
    export_button.grid(row=1, column=0, padx=10, pady=10)

    back_button = tk.Button(table_frame, text="Вернуться", command=reset_app)
    back_button.grid(row=2, column=0, padx=10, pady=10)

def export_to_excel():
    try:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            future_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Данные успешно экспортированы в {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ошибка при экспорте данных: {e}")

def reset_app():
    global future_df
    future_df = pd.DataFrame()
    for widget in graph_frame.winfo_children():
        widget.destroy()
    for widget in table_frame.winfo_children():
        widget.destroy()
    start_screen()

def start_screen():
    for widget in root.winfo_children():
        widget.grid_forget()

    label = tk.Label(root, text="Выберите количество недель для прогноза:", font=("Arial", 14))
    label.grid(row=0, column=0, padx=10, pady=10)

    global week_choice
    week_choice = ttk.Combobox(root, values=[1, 2, 3, 4, 5, 6], state="readonly", font=("Arial", 12))
    week_choice.grid(row=1, column=0, padx=10, pady=10)
    week_choice.set(1)

    predict_button = tk.Button(root, text="Прогнозировать", font=("Arial", 12), command=train_and_predict)
    predict_button.grid(row=2, column=0, padx=10, pady=10)

root = tk.Tk()
root.title("Прогноз цен на арматуру")
root.geometry("800x600")

start_screen()

graph_frame = tk.Frame(root)
graph_frame.grid(row=0, column=1, rowspan=4, padx=10, pady=10)

table_frame = tk.Frame(root)
table_frame.grid(row=0, column=0, padx=10, pady=10)

root.mainloop()