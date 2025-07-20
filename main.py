from itertools import count
import pandas as pd
from tabulate import tabulate
import jdatetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Treeview


def jalali_to_gregorian(jdate_str):
    y, m, d = map(int,jdate_str.split('-'))
    jdate= jdatetime.date(y, m , d)
    return jdate.togregorian()

def load_file():
    file_path= filedialog.askopenfilename(filetypes=[("Excel file", "*.xlsx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)


def run_lottery():
    try:
        file_path = file_entry.get()
        df = pd.read_excel(file_path)

        group = group_entry.get().strip()
        start_input = start_entry.get().strip()
        end_input = end_entry.get().strip()
        how_many = int(count_entry.get())

        filtered = df

        if start_input or end_input:
            if 'time' not in df.columns:
                raise Exception("ستون 'time' در فایل پیدا نشد.")

            df['time'] = df['time'].apply(jalali_to_gregorian)
            df['time'] = pd.to_datetime(df['time'])

            if start_input:
                start = pd.to_datetime(jalali_to_gregorian(start_input))
                filtered = filtered[filtered['time'] >= start]

            if end_input:
                end = pd.to_datetime(jalali_to_gregorian(end_input))
                filtered = filtered[filtered['time'] <= end]

        if group:
            filtered = filtered[filtered['group'] == group]

        print("تعداد افراد بعد از فیلتر:", len(filtered))

        if len(filtered) < how_many:
            messagebox.showerror("خطا", f'فقط {len(filtered)} نفر یافت شد. لطفاً تعداد کمتری انتخاب کنید.')
        else:
            selected = filtered.sample(n=how_many)

            for row in result_table.get_children():
                result_table.delete(row)

            for _, row in selected.iterrows():
                result_table.insert("", tk.END, values=list(row))

            selected.to_excel("final.xlsx", index=False)
            messagebox.showinfo("موفق", "نتیجه در فایل final.xlsx ذخیره شد.")

    except Exception as e:
        messagebox.showerror("خطا", str(e))

#---رابط گرافیکی ---
root = tk.Tk()
root.title("قرعه‌کشی باشگاه مشتریان سفیر")
root.geometry("750x500")

tk.Label(root, text="فایل اکسل:").pack()
file_entry = tk.Entry(root, width=50)
file_entry.pack()
tk.Button(root, text="انتخاب فایل", command=load_file).pack(pady=5)

tk.Label(root, text="گروه (تهران یا شهرستان):").pack()
group_entry = tk.Entry(root)
group_entry.pack()

tk.Label(root, text="از تاریخ (مثلاً 1403-01-01):").pack()
start_entry = tk.Entry(root)
start_entry.pack()

tk.Label(root, text="تا تاریخ (مثلاً 1404-01-01):").pack()
end_entry = tk.Entry(root)
end_entry.pack()

tk.Label(root, text="چند نفر انتخاب شوند؟").pack()
count_entry = tk.Entry(root)
count_entry.pack()

tk.Button(root, text=" انجام قرعه‌کشی", command=run_lottery, bg="green", fg="white").pack(pady=11)

# جدول نمایش نتیجه
columns = ["name", "phone" ,"national" ,"group", "time"]

result_table = Treeview(root, columns=columns, show='headings')
for col in columns:
    result_table.heading(col, text=col)
    result_table.column(col, width=100)
result_table.pack(fill=tk.BOTH, expand=True)

root.mainloop()
