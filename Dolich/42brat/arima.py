#это вторая версия приложения на ARIMA

import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pickle
from PIL import Image, ImageTk

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

selected_formula_choice = "Классическая формула"

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

    last_6_weeks = train_data.tail(6)
    ax.plot(last_6_weeks.index, last_6_weeks['Цена на арматуру'], label='Последние 6 недель', color='blue', linewidth=2)

    ax.plot(future_df['dt'][:weeks], future_df['Цена на арматуру'][:weeks], label='Предсказанные цены', color='red', linestyle='-', linewidth=2)

    ax.scatter(future_df['dt'][:weeks], future_df['Цена на арматуру'][:weeks], color='red', zorder=5)

    for i in range(weeks):
        ax.text(future_df['dt'][i], future_df['Цена на арматуру'][i], f"Неделя {i+1}", color='black', fontsize=9, ha='center')

    ax.set_title('Прогноз цен на арматуру')
    ax.set_xlabel('Дата')
    ax.set_ylabel('Цена на арматуру')
    ax.legend()
    ax.grid()

    canvas = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

def display_predictions(weeks):
    for widget in table_frame.winfo_children():
        widget.destroy()

    predictions_table = tk.Frame(table_frame)
    predictions_table.pack(fill=tk.BOTH, expand=True)

    tk.Label(predictions_table, text="Дата", font=("Arial", 10)).grid(row=0, column=0, padx=5, pady=5)
    tk.Label(predictions_table, text="Прогноз цены", font=("Arial", 10)).grid(row=0, column=1, padx=5, pady=5)

    for i in range(weeks):
        tk.Label(predictions_table, text=f"{future_df['dt'][i].strftime('%d-%m-%Y')}", font=("Arial", 10)).grid(row=i+1, column=0, padx=5, pady=5)
        tk.Label(predictions_table, text=f"{future_df['Цена на арматуру'][i]:.2f} ₽", font=("Arial", 10)).grid(row=i+1, column=1, padx=5, pady=5)

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
    if selected_formula_choice == "Классическая формула":
        return (week_6_price - current_price) * quantity
    elif selected_formula_choice == "Своя формула":
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

def open_formulas_menu():
    global formulas_menu
    formulas_menu = tk.Toplevel(root)
    formulas_menu.title("Меню формул")
    formulas_menu.geometry("500x500")

    formulas_frame = tk.Frame(formulas_menu)
    formulas_frame.pack(pady=20)

    tk.Label(formulas_frame, text="Сохранённые формулы:", font=("Arial", 12)).pack()

    global formulas_listbox
    formulas_listbox = tk.Listbox(formulas_frame, height=6, font=("Arial", 10))
    formulas_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

    update_formulas_list()

    global formula_name_entry
    formula_name_entry = tk.Entry(formulas_frame, font=("Arial", 10), width=20)
    formula_name_entry.pack(pady=5)

    global custom_formula_text
    custom_formula_text = tk.Text(formulas_frame, font=("Arial", 10), width=40, height=6)
    custom_formula_text.pack(pady=5)

    tk.Button(formulas_menu, text="Сохранить формулу", font=("Arial", 10), command=save_custom_formula).pack(pady=5)
    tk.Button(formulas_menu, text="Удалить формулу", font=("Arial", 10), command=delete_custom_formula).pack(pady=5)
    tk.Button(formulas_menu, text="Описание классической формулы", font=("Arial", 10), command=show_classic_formula_description).pack(pady=5)

    tk.Label(formulas_menu, text="Выберите формулу для прогноза:", font=("Arial", 12)).pack(pady=5)

    global selected_formula
    selected_formula = ttk.Combobox(formulas_menu, values=["Классическая формула"] + list(formulas_cache.keys()), state="readonly", font=("Arial", 10), width=20)
    selected_formula.set("Классическая формула")
    selected_formula.pack(pady=5)

    selected_formula.bind("<<ComboboxSelected>>", update_selected_formula)

def update_selected_formula(event=None):
    global selected_formula_choice
    selected_formula_choice = selected_formula.get()

root = tk.Tk()
root.title("Программа прогнозирования")
root.geometry("900x800")

main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

control_frame = tk.Frame(main_frame)
control_frame.pack(fill=tk.X, padx=10, pady=10)

calc_icon = Image.open("C:/Users/User/Documents/GitHub/Intensiv_3/Dolich/42brat/calculator_icon.png")
calc_icon = calc_icon.resize((30, 30))
calc_icon = ImageTk.PhotoImage(calc_icon)

calc_button = tk.Button(control_frame, image=calc_icon, command=open_formulas_menu, width=30, height=30)
calc_button.pack(side=tk.LEFT, padx=5)

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

predict_button = tk.Button(control_frame, text="Прогнозировать", font=("Arial", 10), command=train_and_predict)
predict_button.pack(side=tk.LEFT, padx=5)

export_button = tk.Button(control_frame, text="Экспортировать в Excel", font=("Arial", 10), command=export_to_excel)
export_button.pack(side=tk.LEFT, padx=5)

graph_frame = tk.Frame(main_frame)
graph_frame.pack(fill=tk.BOTH, expand=True)

table_frame = tk.Frame(main_frame)
table_frame.pack(fill=tk.BOTH, expand=True)

trend_frame = tk.Frame(main_frame)
trend_frame.pack(fill=tk.BOTH, expand=True)

root.mainloop()
