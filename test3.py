import tkinter as tk
from tkinter import messagebox, simpledialog
from tkinter import ttk
import pandas as pd

# Глобальная переменная для хранения данных сотрудников
employees = []

# Функция для добавления нового сотрудника
def add_employee():
    name = simpledialog.askstring("Имя", "Введите имя сотрудника:")
    position = simpledialog.askstring("Должность", "Введите должность сотрудника:")
    schedule = simpledialog.askstring("График", "Введите график работы:")

    if name and position and schedule:
        employees.append({'Имя': name, 'Должность': position, 'График': schedule})
        update_employee_list()
    else:
        messagebox.showwarning("Вход", "Пожалуйста, заполните все поля.")

# Функция для редактирования выбранного сотрудника
def edit_employee():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Редактирование", "Пожалуйста, выберите сотрудника для редактирования.")
        return

    item_id = selected_item[0]
    employee = employees[int(item_id)]

    name = simpledialog.askstring("Имя", "Введите имя сотрудника:", initialvalue=employee['Имя'])
    position = simpledialog.askstring("Должность", "Введите должность сотрудника:", initialvalue=employee['Должность'])
    schedule = simpledialog.askstring("График", "Введите график работы:", initialvalue=employee['График'])

    if name and position and schedule:
        employees[int(item_id)] = {'Имя': name, 'Должность': position, 'График': schedule}
        update_employee_list()
    else:
        messagebox.showwarning("Вход", "Пожалуйста, заполните все поля.")

# Функция для удаления выбранного сотрудника
def delete_employee():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Удаление", "Пожалуйста, выберите сотрудника для удаления.")
        return

    item_id = selected_item[0]
    employees.pop(int(item_id))
    update_employee_list()

# Функция для обновления списка сотрудников в интерфейсе
def update_employee_list():
    # Очищаем текущие данные в Treeview
    for item in tree.get_children():
        tree.delete(item)

    # Добавляем сотрудников в Treeview
    for index, employee in enumerate(employees):
        tree.insert('', 'end', iid=index, values=(employee['Имя'], employee['Должность'], employee['График']))

# Функция для экспорта данных в Excel
def export_to_excel():
    df = pd.DataFrame(employees)
    df.to_excel("employees_schedule.xlsx", index=False)
    messagebox.showinfo("Экспорт", "Данные экспортированы в employees_schedule.xlsx")

# Создание основного окна интерфейса
root = tk.Tk()
root.title("Управление графиками сотрудников")
root.geometry("600x400")  # Задаем размер окна
root.configure(bg="#f0f0f0")  # Устанавливаем фон

# Кнопки управления
button_frame = ttk.Frame(root)
button_frame.pack(pady=10)

add_button = ttk.Button(button_frame, text="Добавить сотрудника", command=add_employee)
add_button.grid(row=0, column=0, padx=5)

edit_button = ttk.Button(button_frame, text="Редактировать сотрудника", command=edit_employee)
edit_button.grid(row=0, column=1, padx=5)

delete_button = ttk.Button(button_frame, text="Удалить сотрудника", command=delete_employee)
delete_button.grid(row=0, column=2, padx=5)

export_button = ttk.Button(button_frame, text="Экспортировать в Excel", command=export_to_excel)
export_button.grid(row=0, column=3, padx=5)

# Настройка Treeview для отображения сотрудников
columns = ('Имя', 'Должность', 'График')
tree = ttk.Treeview(root, columns=columns, show='headings')
tree.heading('Имя', text='Имя')
tree.heading('Должность', text='Должность')
tree.heading('График', text='График')

# Установка ширины колонок и цвета заголовков
tree.column('Имя', width=200)
tree.column('Должность', width=150)
tree.column('График', width=150)

# Настройка стиля Treeview
style = ttk.Style()
style.configure("Treeview",
                background="#ffffff",
                foreground="#000000",
                rowheight=25,
                fieldbackground="#ffffff")
style.configure("Treeview.Heading",
                background="#007acc",
                foreground="Grey",
                font=('Arial', 10, 'bold'))
style.map('Treeview',
          background=[('selected', '#007acc')],
          foreground=[('selected', 'white')])

tree.pack(pady=20)

# Запуск главного цикла приложения
root.mainloop()

