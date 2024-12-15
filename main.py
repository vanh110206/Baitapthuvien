import tkinter as tk
from tkinter import messagebox
import pandas as pd
import datetime
import csv
import os

root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("500x700")  

data = {
    "Tên nhân viên": [],
    "Ngày sinh": [],
    "CMND": [],
    "Giới tính": [],
    "Ngày cấp": [],
    "Nơi cấp": [],
    "Chức danh": [],
    "Mã": [],
    "Đơn vị": []
}


tk.Label(root, text="Tên nhân viên:").grid(row=0, column=0, padx=10, pady=10)
tk.Label(root, text="Ngày sinh (DD/MM/YYYY):").grid(row=1, column=0, padx=10, pady=10)
tk.Label(root, text="CMND:").grid(row=2, column=0, padx=10, pady=10)
tk.Label(root, text="Giới tính:").grid(row=3, column=0, padx=10, pady=10)
tk.Label(root, text="Ngày cấp (DD/MM/YYYY):").grid(row=4, column=0, padx=10, pady=10)
tk.Label(root, text="Nơi cấp:").grid(row=5, column=0, padx=10, pady=10)
tk.Label(root, text="Chức danh:").grid(row=6, column=0, padx=10, pady=10)
tk.Label(root, text="Mã:").grid(row=7, column=0, padx=10, pady=10)
tk.Label(root, text="Đơn vị:").grid(row=8, column=0, padx=10, pady=10)


name_entry = tk.Entry(root)
dob_entry = tk.Entry(root)
id_entry = tk.Entry(root)
gender_entry = tk.Entry(root)
issue_date_entry = tk.Entry(root)
place_entry = tk.Entry(root)
position_entry = tk.Entry(root)
employee_code_entry = tk.Entry(root)
unit_entry = tk.Entry(root)

name_entry.grid(row=0, column=1, padx=10, pady=10)
dob_entry.grid(row=1, column=1, padx=10, pady=10)
id_entry.grid(row=2, column=1, padx=10, pady=10)
gender_entry.grid(row=3, column=1, padx=10, pady=10)
issue_date_entry.grid(row=4, column=1, padx=10, pady=10)
place_entry.grid(row=5, column=1, padx=10, pady=10)
position_entry.grid(row=6, column=1, padx=10, pady=10)
employee_code_entry.grid(row=7, column=1, padx=10, pady=10)
unit_entry.grid(row=8, column=1, padx=10, pady=10)

# Hàm lưu thông tin nhân viên vào CSV
def save_employee():
    name = name_entry.get()
    dob = dob_entry.get()
    id_card = id_entry.get()
    gender = gender_entry.get()
    issue_date = issue_date_entry.get()
    place = place_entry.get()
    position = position_entry.get()
    employee_code = employee_code_entry.get()
    unit = unit_entry.get()

    # Kiểm tra định dạng ngày sinh và ngày cấp
    try:
        dob = datetime.datetime.strptime(dob, '%d/%m/%Y').strftime('%d%m%Y')
        issue_date = datetime.datetime.strptime(issue_date, '%d/%m/%Y').strftime('%d%m%Y')
    except ValueError:
        messagebox.showerror("Lỗi", "Ngày sinh hoặc ngày cấp không đúng định dạng. Vui lòng nhập theo định dạng DD/MM/YYYY.")
        return
    
   

    # Lưu vào file CSV với mã hóa utf-8
    with open('employees.csv', mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([name, dob, id_card, gender, issue_date, place, position, employee_code, unit])  # Thêm mã nhân viên và đơn vị vào file

    messagebox.showinfo("Thông báo", "Thông tin nhân viên đã được lưu.")
    clear_fields()

# Hàm xóa các trường nhập liệu
def clear_fields():
    name_entry.delete(0, tk.END)
    dob_entry.delete(0, tk.END)
    id_entry.delete(0, tk.END)
    gender_entry.delete(0, tk.END)
    issue_date_entry.delete(0, tk.END)
    place_entry.delete(0, tk.END)
    position_entry.delete(0, tk.END)
    employee_code_entry.delete(0, tk.END)
    unit_entry.delete(0, tk.END)

# Hàm hiển thị các nhân viên có sinh nhật hôm nay
def show_birthday_today():
    today = datetime.datetime.now().strftime('%d%m')  # Lấy ngày hiện tại dưới định dạng DDMM
    employees = []
    
    # Đọc dữ liệu từ file CSV
    with open('employees.csv', mode='r', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            if row:  # Kiểm tra dòng không rỗng
                birth_date = row[1]  # Ngày sinh (theo định dạng ddmmyyyy)
                if birth_date[:4] == today:  # So sánh ngày sinh (chỉ ngày và tháng)
                    employees.append(row)
    
    # Kiểm tra và hiển thị kết quả
    if employees:
        result = "\n".join([f"{emp[0]} - Ngày sinh: {emp[1][:2]}/{emp[1][2:4]}/{emp[1][4:]} " for emp in employees])
        messagebox.showinfo("Sinh nhật hôm nay", result)
    else:
        messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào có sinh nhật hôm nay.")

# Hàm xuất danh sách nhân viên ra file Excel
def export_to_excel():
    employees = []
    with open('employees.csv', mode='r', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            if row:  # Kiểm tra dòng không rỗng
                employees.append(row)
    
    # Chuyển dữ liệu thành DataFrame và sắp xếp theo tuổi giảm dần
    df = pd.DataFrame(employees, columns=["Tên nhân viên", "Ngày sinh", "CMND", "Giới tính", "Ngày cấp", "Nơi cấp", "Chức danh", "Mã", "Đơn vị"])
    
    # Kiểm tra và chuyển đổi ngày sinh
    try:
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d%m%Y', errors='coerce')
        df['Tuổi'] = df['Ngày sinh'].apply(lambda x: datetime.datetime.now().year - x.year if pd.notnull(x) else None)
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi chuyển đổi ngày sinh: {e}")
        return
    
    # Sắp xếp theo tuổi giảm dần
    df_sorted = df.sort_values(by='Tuổi', ascending=False)
    
    # Xuất ra file Excel
    try:
        df_sorted.to_excel('employee_list.xlsx', index=False)
        messagebox.showinfo("Thông báo", "Danh sách nhân viên đã được xuất ra file Excel.")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi khi xuất file Excel: {e}")

tk.Button(root, text="Lưu thông tin", command=save_employee).grid(row=9, column=0, columnspan=2, pady=10)
tk.Button(root, text="Sinh nhật hôm nay", command=show_birthday_today).grid(row=10, column=0, columnspan=2, pady=10)
tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel).grid(row=11, column=0, columnspan=2, pady=10)
root.mainloop()