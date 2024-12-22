import tkinter as tk
from tkinter import messagebox, simpledialog
from tkinter import ttk
from employee_manager import EmployeeManager

#СПЕЦИЯ
employee_manager = EmployeeManager()

#ДОБАВЛЕНИЕ
def add_employee():
    name = simpledialog.askstring("Имя", "Введите имя сотрудника:")
    position = simpledialog.askstring("Должность", "Введите должность сотрудника:")
    schedule = simpledialog.askstring("График", "Введите график работы:")

    if employee_manager.add_employee(name, position, schedule):
        update_employee_list()
    else:
        messagebox.showwarning("Вход", "Пожалуйста, заполните все поля.")

#РЕДАКТИРОВАНИЕ
def edit_employee():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Редактирование", "Пожалуйста, выберите сотрудника для редактирования.")
        return

    item_id = selected_item[0]
    employee = employee_manager.get_employees()[int(item_id)]

    name = simpledialog.askstring("Имя", "Введите имя сотрудника:", initialvalue=employee['Имя'])
    position = simpledialog.askstring("Должность", "Введите должность сотрудника:", initialvalue=employee['Должность'])
    schedule = simpledialog.askstring("График", "Введите график работы:", initialvalue=employee['График'])

    if employee_manager.edit_employee(int(item_id), name, position, schedule):
        update_employee_list()
    else:
        messagebox.showwarning("Вход", "Пожалуйста, заполните все поля.")

#УДАЛЕНИЕ
def delete_employee():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Удаление", "Пожалуйста, выберите сотрудника для удаления.")
        return

    item_id = selected_item[0]
    employee_manager.delete_employee(int(item_id))
    update_employee_list()

#ОБНОВЛЕНИЕ
def update_employee_list():
    for item in tree.get_children():
        tree.delete(item)

    for index, employee in enumerate(employee_manager.get_employees()):
        tree.insert('', 'end', iid=index, values=(employee['Имя'], employee['Должность'], employee['График']))

#ШедевроЭксельчик
def export_to_excel():
    filename = employee_manager.export_to_excel()
    messagebox.showinfo("Экспорт", f"Данные экспортированы в {filename}")

# Создание основного окна интерфейса
root = tk.Tk()
root.title("Управление графиками сотрудников")
root.geometry("700x500")  #ОКНО
root.configure(bg="#E0F7FA")  #ФОН

#УПРАВЛЕНИЕ
button_frame = ttk.Frame(root)
button_frame.pack(pady=10)

#СТИЛЬКНОПОК
style = ttk.Style()
style.configure("TButton",
                padding=10,
                relief="flat",
                background="#B2EBF2",
                foreground="#004D40",
                font=('Arial', 10, 'bold'))
style.map("TButton",
          background=[('active', '#80DEEA')])  #НАВЕДЕНИЕ

# Кнопки
add_button = ttk.Button(button_frame, text="Добавить сотрудника", command=add_employee)
add_button.grid(row=0, column=0, padx=5)

edit_button = ttk.Button(button_frame, text="Редактировать сотрудника", command=edit_employee)
edit_button.grid(row=0, column=1, padx=5)

delete_button = ttk.Button(button_frame, text="Удалить сотрудника", command=delete_employee)
delete_button.grid(row=0, column=2, padx=5)

export_button = ttk.Button(button_frame, text="Экспортировать в Excel", command=export_to_excel)
export_button.grid(row=0, column=3, padx=5)

# ДЕРЕВО
columns = ('Имя', 'Должность', 'График')
tree = ttk.Treeview(root, columns=columns, show='headings')
tree.heading('Имя', text='Имя')
tree.heading('Должность', text='Должность')
tree.heading('График', text='График')

# КОЛОНКИ И ЦВЕТА ЗАГОЛОВКОВ
tree.column('Имя', width=200)
tree.column('Должность', width=150)
tree.column('График', width=150)

# Настройка стиля Treeview
style.configure("Treeview",
                background="#ffffff",
                foreground="#000000",
                rowheight=25,
                fieldbackground="#ffffff")
style.configure("Treeview.Heading",
                background="#4DD0E1",
                foreground="#008080",
                font=('Arial', 10, 'bold'))
style.map('Treeview',
          background=[('selected', '#B2EBF2')],
          foreground=[('selected', '#008080')])

tree.pack(pady=20, fill=tk.BOTH, expand=True)  #ЗАКРЫТЬ ПРОБЕЛ

update_employee_list()  #ОБНОВА
root.mainloop()