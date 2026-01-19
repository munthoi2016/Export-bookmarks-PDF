import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader
from openpyxl import Workbook
import os

def extract_bookmarks(outlines, level=0, parent=None, bookmarks_list=None):
    """
    Đệ quy duyệt toàn bộ bookmark, lấy tên, số trang, cấp độ và bookmark cha
    """
    if bookmarks_list is None:
        bookmarks_list = []

    for item in outlines:
        # Nếu là danh sách con
        if isinstance(item, list):
            extract_bookmarks(item, level + 1, parent=last_title, bookmarks_list=bookmarks_list)
        else:
            title = item.title
            page_number = reader.get_destination_page_number(item) + 1
            last_title = title
            bookmarks_list.append({
                "title": title,
                "page": page_number,
                "level": level,
                "parent": parent
            })

    return bookmarks_list


def export_bookmarks_to_excel(pdf_path, excel_path):
    global reader
    reader = PdfReader(pdf_path)

    # ✅ Đảm bảo tương thích với mọi bản PyPDF2
    outlines = None
    if hasattr(reader, "outline") and not callable(reader.outline):
        outlines = reader.outline
    elif hasattr(reader, "get_outlines"):
        outlines = reader.get_outlines()  # bản mới (>=3.0)
    elif hasattr(reader, "getOutlines"):
        outlines = reader.getOutlines()   # bản cũ (<3.0)
    else:
        raise Exception("Phiên bản PyPDF2 không hỗ trợ đọc bookmark")

    if not outlines:
        raise Exception("File PDF không có bookmark hoặc không đọc được")

    bookmarks = extract_bookmarks(outlines)

    wb = Workbook()
    ws = wb.active
    ws.title = "Bookmarks"

    ws.append(["Title", "Page", "Level", "Parent"])

    for bm in bookmarks:
        indent_title = " " * (bm["level"] * 4) + bm["title"]
        ws.append([indent_title, bm["page"], bm["level"], bm["parent"] or ""])

    wb.save(excel_path)
    return len(bookmarks)


# ---------------- GUI ----------------
def select_pdf():
    pdf_path = filedialog.askopenfilename(
        title="Chọn file PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    if pdf_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, pdf_path)


def select_excel():
    excel_path = filedialog.asksaveasfilename(
        title="Lưu file Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if excel_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, excel_path)


def run_export():
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()

    if not pdf_path or not os.path.exists(pdf_path):
        messagebox.showerror("Lỗi", "Vui lòng chọn file PDF hợp lệ.")
        return
    if not excel_path:
        messagebox.showerror("Lỗi", "Vui lòng chọn nơi lưu file Excel.")
        return

    try:
        count = export_bookmarks_to_excel(pdf_path, excel_path)
        messagebox.showinfo("Hoàn tất", f"✅ Xuất thành công {count} bookmark vào:\n{excel_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất bookmark:\n{e}")


# ---------------- TẠO GIAO DIỆN ----------------
root = tk.Tk()
root.title("Xuất Bookmark PDF ra Excel")
root.geometry("500x200")
root.resizable(False, False)

tk.Label(root, text="Chọn file PDF:").pack(anchor="w", padx=10, pady=(10, 0))
frame_pdf = tk.Frame(root)
frame_pdf.pack(fill="x", padx=10)
pdf_entry = tk.Entry(frame_pdf, width=50)
pdf_entry.pack(side="left", fill="x", expand=True)
tk.Button(frame_pdf, text="Chọn...", command=select_pdf).pack(side="right")

tk.Label(root, text="Lưu thành file Excel:").pack(anchor="w", padx=10, pady=(10, 0))
frame_excel = tk.Frame(root)
frame_excel.pack(fill="x", padx=10)
excel_entry = tk.Entry(frame_excel, width=50)
excel_entry.pack(side="left", fill="x", expand=True)
tk.Button(frame_excel, text="Chọn...", command=select_excel).pack(side="right")

tk.Button(root, text="📘 Xuất Bookmark", bg="#4CAF50", fg="white",
          font=("Arial", 11, "bold"), command=run_export).pack(pady=20)

root.mainloop()
