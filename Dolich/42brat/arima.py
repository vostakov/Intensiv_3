#это вторая версия приложения на ARIMA

import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pickle

train_data = pd.read_excel('C:/Users/User/Documents/GitHub/Intensiv_3/Dataset/data/train.xlsx')
train_data['dt'] = pd.to_datetime(train_data['dt'])
train_data.set_index('dt', inplace=True)

future_df = pd.DataFrame()

formulas_cache_file = "formulas_cache.pkl"

def load_saved_formulas():
    try:
        with open(formulas_cache_file, "rb") as f:
            return pickle.load(f)
    except FileNotFoundError:
        return {}

def save_formulas():
    with open(formulas_cache_file, "wb") as f:
        pickle.dump(formulas_cache, f)

formulas_cache = load_saved_formulas()

def train_and_predict():
    global future_df
    try:
        weeks = int(week_choice.get())
        if weeks < 1 or weeks > 6:
            messagebox.showwarning("Warning", "Количество недель должно быть от 1 до 6!")
            return

        model = ARIMA(train_data['Цена на арматуру'], order=(5, 1, 0))
        model_fit = model.fit()
        forecast = model_fit.forecast(steps=10)
        future_dates = pd.date_range(start=train_data.index.max() + pd.Timedelta(days=7), periods=10, freq='W-MON')
        future_df = pd.DataFrame({'dt': future_dates, 'Цена на арматуру': forecast})

        plot_graph(weeks)
        display_predictions(weeks)
        analyze_future_trend(weeks)
    except Exception as e:
        messagebox.showerror("Error", f"Ошибка при прогнозировании: {e}")

def plot_graph(weeks):
    global future_df
    for widget in graph_frame.winfo_children():
        widget.destroy()

    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(train_data.index, train_data['Цена на арматуру'], label='Исторические данные', color='blue', linewidth=2)
    ax.plot(future_df['dt'][:weeks], future_df['Цена на арматуру'][:weeks], label='Предсказанные цены', color='red', linestyle='-', linewidth=2)
    ax.plot([train_data.index[-1], future_df['dt'].iloc[0]], [train_data['Цена на арматуру'].iloc[-1], future_df['Цена на арматуру'].iloc[0]], color='red', linestyle='--', linewidth=2)
    ax.set_title('Прогноз цен на арматуру')
    ax.set_xlabel('Дата')
    ax.set_ylabel('Цена на арматуру')
    ax.legend()
    ax.grid()

    canvas = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

def display_predictions(weeks):
    global future_df
    for widget in table_frame.winfo_children():
        widget.destroy()

    table = ttk.Treeview(table_frame, columns=("Дата", "Цена на арматуру"), show="headings", height=10)
    table.heading("Дата", text="Дата")
    table.heading("Цена на арматуру", text="Цена на арматуру")
    table.pack(fill=tk.BOTH, expand=True)

    for _, row in future_df.iloc[:weeks].iterrows():
        table.insert("", "end", values=(row["dt"].strftime('%Y-%m-%d'), row["Цена на арматуру"]))

    export_button = tk.Button(table_frame, text="Экспортировать в Excel", command=export_to_excel)
    export_button.pack(pady=5)

def analyze_future_trend(weeks):
    for widget in trend_frame.winfo_children():
        widget.destroy()

    if weeks < 6:
        week_6_price = future_df.iloc[weeks]["Цена на арматуру"]
        week_5_price = future_df.iloc[weeks - 1]["Цена на арматуру"]
    else:
        week_6_price = future_df.iloc[6]["Цена на арматуру"]
        week_5_price = future_df.iloc[5]["Цена на арматуру"]
    
    trend_message = ""
    color = "green" if week_6_price > week_5_price else "red"
    
    if week_6_price > week_5_price:
        trend_message = "цена будет расти"
    else:
        trend_message = "цена будет падать"
    
    trend_label = tk.Label(trend_frame, text=trend_message, font=("Arial", 12), fg=color)
    trend_label.pack()

    try:
        budget = float(budget_entry.get())
        quantity = float(quantity_entry.get())  
        current_price = train_data['Цена на арматуру'].iloc[-1]
        profit_or_loss = calculate_profit_or_loss(budget, current_price, week_6_price, quantity)
        profit_loss_message = f"Прогнозируемая прибыль/убыток: {profit_or_loss:.2f} ₽"
        profit_loss_color = "green" if profit_or_loss > 0 else "red"
        
        profit_loss_label = tk.Label(trend_frame, text=profit_loss_message, font=("Arial", 12), fg=profit_loss_color)
        profit_loss_label.pack()
    
    except ValueError:
        messagebox.showerror("Error", "Пожалуйста, введите корректный бюджет для закупки и количество тонн.")

def calculate_profit_or_loss(budget, current_price, week_6_price, quantity):
    if selected_formula.get() == "Классическая формула":
        return (week_6_price - current_price) * quantity
    elif selected_formula.get() == "Своя формула":
        try:
            custom_formula = custom_formula_text.get("1.0", tk.END).strip()
            result = eval(custom_formula)
            return result
        except Exception as e:
            messagebox.showerror("Error", f"Ошибка в формуле: {e}")
            return 0
    return 0

def save_custom_formula():
    formula_name = formula_name_entry.get().strip()
    custom_formula = custom_formula_text.get("1.0", tk.END).strip()
    if formula_name and custom_formula:
        formulas_cache[formula_name] = custom_formula
        save_formulas()
        update_formulas_list()

def delete_custom_formula():
    selected_formula_name = formulas_listbox.get(tk.ACTIVE)
    if selected_formula_name:
        del formulas_cache[selected_formula_name]
        save_formulas()
        update_formulas_list()

def update_formulas_list():
    formulas_listbox.delete(0, tk.END)
    for formula_name in formulas_cache.keys():
        formulas_listbox.insert(tk.END, formula_name)

def export_to_excel():
    try:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            future_df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Данные успешно экспортированы в {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ошибка при экспорте данных: {e}")

def show_classic_formula_description():
    description = ("Классическая формула прибыли/убытка:\n"
                   "Прибыль/убыток рассчитывается как разница между прогнозируемой ценой (на 6-й неделе) и текущей ценой "
                   "арматуры, умноженная на количество, которое можно закупить на заданный бюджет.\n\n"
                   "Формула:\n"
                   "Прибыль/убыток = (Прогнозируемая цена через 6 недель - Текущая цена) * Количество закупленной арматуры.")
    messagebox.showinfo("Классическая формула", description)

root = tk.Tk()
root.title("Прогноз цен на арматуру")
root.geometry("900x800")  

main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

control_frame = tk.Frame(main_frame)
control_frame.pack(fill=tk.X, padx=10, pady=10)

tk.Label(control_frame, text="Выберите недели:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
week_choice = ttk.Combobox(control_frame, values=[1, 2, 3, 4, 5, 6], state="readonly", font=("Arial", 10), width=5)
week_choice.pack(side=tk.LEFT, padx=5)
week_choice.set(1)

tk.Label(control_frame, text="Бюджет закупки:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
budget_entry = tk.Entry(control_frame, font=("Arial", 10), width=10)
budget_entry.pack(side=tk.LEFT, padx=5)

tk.Label(control_frame, text="Кол-во тонн:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
quantity_entry = tk.Entry(control_frame, font=("Arial", 10), width=10)
quantity_entry.pack(side=tk.LEFT, padx=5)

tk.Label(control_frame, text="Формула прибыли:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
selected_formula = ttk.Combobox(control_frame, values=["Классическая формула", "Своя формула"], state="readonly", font=("Arial", 10), width=15)
selected_formula.pack(side=tk.LEFT, padx=5)
selected_formula.set("Классическая формула")

question_button = ttk.Button(control_frame, text="?", command=show_classic_formula_description)
question_button.pack(side=tk.LEFT, padx=5)

tk.Button(control_frame, text="Прогнозировать", font=("Arial", 10), command=train_and_predict).pack(side=tk.LEFT, padx=5)

custom_formula_frame = tk.Frame(main_frame)
custom_formula_frame.pack(fill=tk.X, padx=10, pady=10)

tk.Label(custom_formula_frame, text="Введите свою формулу:", font=("Arial", 10)).pack()

custom_formula_text = tk.Text(custom_formula_frame, height=3, font=("Arial", 10), width=40)
custom_formula_text.pack(fill=tk.X, pady=5)

tk.Label(custom_formula_frame, text="Название формулы:", font=("Arial", 10)).pack()

formula_name_entry = tk.Entry(custom_formula_frame, font=("Arial", 10), width=40)
formula_name_entry.pack(fill=tk.X, pady=5)

tk.Button(custom_formula_frame, text="Сохранить формулу", font=("Arial", 10), command=save_custom_formula).pack(side=tk.LEFT, padx=5)

tk.Button(custom_formula_frame, text="Удалить формулу", font=("Arial", 10), command=delete_custom_formula).pack(side=tk.LEFT, padx=5)

formulas_frame = tk.Frame(main_frame)
formulas_frame.pack(fill=tk.X, padx=10, pady=10)

tk.Label(formulas_frame, text="Сохранённые формулы:", font=("Arial", 10)).pack()

formulas_listbox = tk.Listbox(formulas_frame, height=6, font=("Arial", 10))
formulas_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

update_formulas_list()

graph_frame = tk.Frame(main_frame)
graph_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

trend_frame = tk.Frame(main_frame)
trend_frame.pack(fill=tk.X, padx=10, pady=10)

table_frame = tk.Frame(main_frame)
table_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

root.mainloop()
