import pandas as pd
import json
import os

#добавление, редактирование, удаление и экспорт данных о сотрудниках

class EmployeeManager:
    def __init__(self, filename='employees.json'):
        self.filename = filename
        self.employees = []
        self.load_employees()

    def add_employee(self, name, position, schedule):
        if name and position and schedule:
            self.employees.append({'Имя': name, 'Должность': position, 'График': schedule})
            self.save_employees()
            return True
        return False

    def edit_employee(self, index, name, position, schedule):
        if 0 <= index < len(self.employees) and name and position and schedule:
            self.employees[index] = {'Имя': name, 'Должность': position, 'График': schedule}
            self.save_employees()
            return True
        return False

    def delete_employee(self, index):
        if 0 <= index < len(self.employees):
            self.employees.pop(index)
            self.save_employees()
            return True
        return False

    def export_to_excel(self, filename="employees_schedule.xlsx"):
        df = pd.DataFrame(self.employees)
        df.to_excel(filename, index=False)
        return filename

    def get_employees(self):
        return self.employees

    def save_employees(self):
        with open(self.filename, 'w', encoding='utf-8') as f:
            json.dump(self.employees, f, ensure_ascii=False, indent=4)

    def load_employees(self):
        if os.path.exists(self.filename):
            with open(self.filename, 'r', encoding='utf-8') as f:
                self.employees = json.load(f)
