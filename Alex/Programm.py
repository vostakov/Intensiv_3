import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import joblib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class App:
    def __init__(self, master):
        self.master = master
        master.title("Прогнозирование цены на арматуру")
        self.master.geometry("800x600")  # Установка размера окна

        # Переменные для хранения данных
        self.data = None
        self.latest_date = None
        self.model = self.load_model()
        self.load_data()  # Загрузим данные сразу при старте

        # Переменная для хранения результатов
        self.saved_results = None

        # Создание интерфейса
        self.create_interface()

    def load_model(self):
        """Метод для загрузки модели."""
        try:
            model = joblib.load('catboost_model.pkl')
            return model
        except FileNotFoundError:
            messagebox.showerror("Ошибка", "Не удалось загрузить модель. Убедитесь, что модель сохранена.")
            return None

    def load_data(self):
        """Метод для загрузки данных из файла train.xlsx."""
        try:
            self.data = pd.read_excel('train.xlsx')
            self.latest_date = self.data['dt'].max()  # Находим последнюю дату в загруженных данных
        except FileNotFoundError:
            messagebox.showerror("Ошибка", "Файл train.xlsx не найден.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при загрузке данных: {e}")

    def create_interface(self):
        """Создает интерфейс программы."""
        # Фрейм для вывода советов
        advice_frame = tk.Frame(self.master)
        advice_frame.pack(padx=10, pady=(55, 10), side=tk.RIGHT, anchor='n')  # Правый верхний угол

        # Создаем текстовое поле для совета по покупке
        self.purchase_advice_text = tk.Text(advice_frame, width=40, height=5, state='disabled')  # Увеличили высоту
        self.purchase_advice_text.pack(pady=(0, 5))

        # Новый текстовый блок для советов по количеству арматуры
        self.quantity_advice_text = tk.Text(advice_frame, width=40, height=5, state='disabled')  # Тот же размер
        self.quantity_advice_text.pack(pady=(0, 5))

        # Новый текстовый блок для советов по бюджету
        self.budget_advice_text = tk.Text(advice_frame, width=40, height=5, state='disabled')  # Тот же размер
        self.budget_advice_text.pack(pady=(0, 5))

        # Блок для отображения прибыли/убытка
        self.profit_loss_text = tk.Text(advice_frame, width=40, height=2, state='disabled', wrap='word')
        self.profit_loss_text.pack(pady=(10, 5))

        # Фрейм для ввода параметров
        input_frame = tk.Frame(self.master)
        input_frame.pack(side=tk.TOP, padx=10, pady=10, fill=tk.X)

        # Список выбора для количества недель
        self.weeks_label = tk.Label(input_frame, text="Недель:")
        self.weeks_label.pack(side=tk.LEFT)
        self.weeks_var = tk.StringVar(value='1')
        self.weeks_dropdown = ttk.Combobox(input_frame, textvariable=self.weeks_var, values=[str(i) for i in range(1, 11)], state="readonly")
        self.weeks_dropdown.pack(side=tk.LEFT, padx=(0, 10))

        # Поле для ввода бюджета
        self.budget_label = tk.Label(input_frame, text="Бюджет:")
        self.budget_label.pack(side=tk.LEFT)
        self.budget_entry = tk.Entry(input_frame)
        self.budget_entry.pack(side=tk.LEFT, padx=(0, 10))

        # Поле для ввода количества арматуры
        self.quantity_label = tk.Label(input_frame, text="Арматура (тонны):")
        self.quantity_label.pack(side=tk.LEFT)
        self.quantity_entry = tk.Entry(input_frame)
        self.quantity_entry.pack(side=tk.LEFT)

        # Кнопка для прогноза
        self.predict_button = tk.Button(input_frame, text="Сделать прогноз", command=self.make_prediction)
        self.predict_button.pack(side=tk.LEFT, padx=(10, 0))

        # Фрейм для вывода результатов
        result_frame = tk.Frame(self.master)
        result_frame.pack(padx=10, pady=10, side=tk.LEFT, fill=tk.BOTH)

        # Поле для вывода прогноза
        self.result_text_forecast = tk.Text(result_frame, width=50, height=12)
        self.result_text_forecast.pack(side=tk.TOP, fill=tk.BOTH)

        # Поле для вывода информации
        self.info_tab_text = tk.Text(result_frame, width=50, height=12, state='disabled')
        self.info_tab_text.pack(side=tk.TOP, pady=(10, 0), fill=tk.BOTH)

        # Поле для ввода названия файла
        self.filename_label = tk.Label(result_frame, text="Название файла:")
        self.filename_label.pack(side=tk.TOP, pady=(10, 0))
        self.filename_entry = tk.Entry(result_frame, width=40)
        self.filename_entry.pack(side=tk.TOP, pady=(0, 10))

        # Кнопка для скачивания результатов
        self.download_button = tk.Button(result_frame, text="Скачать результат", command=self.download_results)
        self.download_button.pack(side=tk.TOP, padx=(10, 0), pady=(5, 0))

        # График
        self.figure_info, self.ax_info = plt.subplots(figsize=(12, 10))
        self.canvas_info = FigureCanvasTkAgg(self.figure_info, master=self.master)
        self.canvas_info.get_tk_widget().pack(pady=10)

    def make_prediction(self):
        """Метод для прогноза цен на арматуру на заданное количество недель."""
        if self.model is None:
            messagebox.showerror("Ошибка", "Сначала загрузите модель!")
            return
        
        if self.data is None:
            messagebox.showerror("Ошибка", "Сначала загрузите данные!")
            return

        try:
            future_weeks = int(self.weeks_var.get())
            purchase_quantity = int(self.quantity_entry.get())
            budget = float(self.budget_entry.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные данные для недель, количества арматуры и бюджета.")
            return

        if self.latest_date is not None:
            start_date = pd.to_datetime(self.latest_date)
        else:
            messagebox.showerror("Ошибка", "Загрузите данные для получения стартовой даты.")
            return

        initial_price = self.data.loc[self.data['dt'] == start_date, 'Цена на арматуру'].values[0]

        future_dates = pd.date_range(start=start_date, periods=future_weeks, freq='W-MON')

        future_data = pd.DataFrame({
            'День': future_dates.day,
            'Месяц': future_dates.month,
            'Год': future_dates.year,
            'ДеньНедели': future_dates.dayofweek,
            'ЦенаL1': initial_price,
            'ЦенаL2': initial_price,
        })

        future_predictions = self.model.predict(future_data)

        result_df = pd.DataFrame({'Дата': future_dates, 'Прогноз цены': future_predictions})
        self.saved_results = result_df  # Сохраняем результаты для дальнейшего использования
        self.result_text_forecast.delete(1.0, tk.END)
        self.result_text_forecast.insert(tk.END, result_df.to_string(index=False))

        # Обновление информации
        max_price = future_predictions.max()
        self.update_info(start_date, future_weeks, purchase_quantity, budget, future_predictions)

        # Отображение графика
        self.plot_forecast(future_weeks, future_predictions)

        # Анализ прогноза и выдача советов
        budget_needed, profit_or_loss = self.analyze_predictions(future_predictions, initial_price, purchase_quantity)

        # Обновление прогноза прибыли или убытка
        self.update_profit_loss(budget, budget_needed)

    def update_info(self, start_date=None, weeks=None, quantity=None, budget=None, predictions=None):
        """Обновляет текстовое поле информацией о прогнозе."""
        self.info_tab_text.config(state='normal')
        self.info_tab_text.delete(1.0, tk.END)
        
        if predictions is not None:
            if predictions[-1] > predictions[0]:
                status = "Рост"
            else:
                status = "Падение"
            
            self.info_tab_text.insert(tk.END, f"Статус цены: {status}\n")
            self.info_tab_text.insert(tk.END, f"Максимальная стоимость: {predictions.max():.2f}\n")

        if start_date:
            self.info_tab_text.insert(tk.END, f"Стартовая дата: {start_date.strftime('%d-%m-%Y')}\n")
        if weeks:
            self.info_tab_text.insert(tk.END, f"Количество недель: {weeks}\n")
        if quantity:
            self.info_tab_text.insert(tk.END, f"Количество арматуры для покупки: {quantity} тонн\n")
        if budget is not None:
            self.info_tab_text.insert(tk.END, f"Ваш бюджет: {budget:.2f} руб.\n")
        
        self.info_tab_text.config(state='disabled')

    def plot_forecast(self, future_weeks, future_predictions):
        """Метод для отображения графика прогноза цен."""
        self.ax_info.clear()
        
        self.ax_info.plot(range(1, future_weeks + 1), future_predictions, marker='o', linestyle='-', color='blue')
        self.ax_info.set_title('Прогноз цен на арматуру')
        self.ax_info.set_xlabel('Неделя')
        self.ax_info.set_ylabel('Цена')
        self.ax_info.set_xticks(range(1, future_weeks + 1))
        self.ax_info.set_xticklabels([f'Неделя {i}' for i in range(1, future_weeks + 1)])
        self.ax_info.grid()
        
        self.canvas_info.draw()

    def analyze_predictions(self, predictions, initial_price, purchase_quantity):
        """Анализирует прогнозируемые значения и выдает советы о покупке."""
        advice_weeks = 0
        
        # Определяем минимальную стоимость прогнозируемых недель
        min_price_forecast = predictions.min()
        first_week_price = predictions[0]

        if predictions[0] > initial_price:
            advice_weeks += 1  

        for price in predictions[1:]:
            if price > initial_price:
                advice_weeks += 1
            else:
                break  

        # Определяем советы
        if advice_weeks > 0:
            advice = f"Рекомендуем покупать арматуру на {advice_weeks + 1} недел(и).\n"
            price_to_use = min_price_forecast
        else:
            advice = "Цены падают. Рекомендуем подождать с покупкой.\n"
            price_to_use = first_week_price

        self.purchase_advice_text.config(state='normal')
        self.purchase_advice_text.delete(1.0, tk.END)
        self.purchase_advice_text.insert(tk.END, advice)
        self.purchase_advice_text.config(state='disabled')

        # Рассчитаем рекомендованное количество арматуры
        try:
            total_quantity_advice = (advice_weeks + 1) * purchase_quantity
            quantity_advice = f"Рекомендуемое количество арматуры: {total_quantity_advice} тонн.\n"

            # Рассчитаем рекомендованный бюджет
            budget_needed = total_quantity_advice * price_to_use
            budget_advice = f"Рекомендуемый бюджет: {budget_needed:.2f} руб.\n"
        except ValueError:
            quantity_advice = "Ошибка в вводе количества арматуры.\n"
            budget_needed = 0  # Если что-то пошло не так с расчетами, устанавливаем 0

        self.quantity_advice_text.config(state='normal')
        self.quantity_advice_text.delete(1.0, tk.END)
        self.quantity_advice_text.insert(tk.END, quantity_advice)
        self.quantity_advice_text.config(state='disabled')

        # Обновим бюджет совет в соответствующем текстовом поле
        self.budget_advice_text.config(state='normal')
        self.budget_advice_text.delete(1.0, tk.END)
        self.budget_advice_text.insert(tk.END, budget_advice)
        self.budget_advice_text.config(state='disabled')

        return budget_needed, total_quantity_advice

    def update_profit_loss(self, budget, budget_needed):
        """Обновляет текстовое поле с прогнозом прибыли или убытка."""
        self.profit_loss_text.config(state='normal')
        self.profit_loss_text.delete(1.0, tk.END)

        profit_or_loss = budget - budget_needed
        if profit_or_loss >= 0:
            self.profit_loss_text.insert(tk.END, f"Остаток: {profit_or_loss:.2f} руб. (Зеленый)")
            self.profit_loss_text.tag_configure('green', foreground='green')
            self.profit_loss_text.tag_add('green', '1.0', '1.end')
        else:
            self.profit_loss_text.insert(tk.END, f"Убыток: {-profit_or_loss:.2f} руб. (Красный)")
            self.profit_loss_text.tag_configure('red', foreground='red')
            self.profit_loss_text.tag_add('red', '1.0', '1.end')

        self.profit_loss_text.config(state='disabled')

    def download_results(self):
        """Метод для скачивания прогноза в формате Excel."""
        if self.saved_results is None:
            messagebox.showerror("Ошибка", "Нет доступных результатов для скачивания.")
            return
        
        filename = self.filename_entry.get().strip()  # Получаем имя файла из поля ввода
        if not filename:
            messagebox.showerror("Ошибка", "Введите название файла для сохранения.")
            return
        
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'  # Добавляем расширение, если оно отсутствует
        
        try:
            self.saved_results.to_excel(filename, index=False)
            messagebox.showinfo("Информация", f"Результаты успешно сохранены в {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении файла: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()