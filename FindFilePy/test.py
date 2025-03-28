import tkinter as tk
from tkinter import messagebox

# 正确密码
CORRECT_PASSWORD = "123456"  # 可修改

def check_password():
    entered = password_entry.get()
    if entered == CORRECT_PASSWORD:
        messagebox.showinfo("成功", "Hello, world!")
        root.destroy()
    else:
        messagebox.showerror("错误", "密码错误！")

# 创建窗口
root = tk.Tk()
root.title("密码验证")
root.geometry("300x150")

# 标签
label = tk.Label(root, text="请输入密码：")
label.pack(pady=10)

# 密码输入框
password_entry = tk.Entry(root, show="*")
password_entry.pack()

# 确认按钮
submit_button = tk.Button(root, text="确认", command=check_password)
submit_button.pack(pady=10)

# 运行主循环
root.mainloop()