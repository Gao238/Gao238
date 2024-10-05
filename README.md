import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document

def load_words_from_docx(file_path):
    doc = Document(file_path)
    words = set()
    for para in doc.paragraphs:
        words.update(para.text.split())
    return words

def count_different_words():
    file1_path = filedialog.askopenfilename(title="Chọn file Word thứ nhất", filetypes=[("Word files", "*.docx")])
    file2_path = filedialog.askopenfilename(title="Chọn file Word thứ hai", filetypes=[("Word files", "*.docx")])

    if not file1_path or not file2_path:
        messagebox.showerror("Lỗi", "Bạn phải chọn cả hai file.")
        return

    words1 = load_words_from_docx(file1_path)
    words2 = load_words_from_docx(file2_path)

    different_words = words1.symmetric_difference(words2)

    result_text.set(f"Số từ khác nhau: {len(different_words)}\nCác từ khác nhau: {', '.join(different_words)}")

# Tạo giao diện
root = tk.Tk()
root.title("Đếm số từ khác biệt giữa 2 file Word")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

btn_count = tk.Button(frame, text="Đếm từ khác nhau", command=count_different_words)
btn_count.pack(pady=5)

result_text = tk.StringVar()
result_label = tk.Label(frame, textvariable=result_text, wraplength=400)
result_label.pack(pady=5)

root.mainloop()
