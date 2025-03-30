
import os
import msoffcrypto
from io import BytesIO
from tkinter import Tk, Label, Button, Text, filedialog, messagebox, Toplevel, Entry
from docx import Document  # 使用 python-docx 解析 .docx 文件

class WordDecryptorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Word解密工具")
        self.root.geometry("600x400")
        
        # 界面元素
        self.label = Label(root, text="選擇加密的Word文件", font=("Arial", 12))
        self.label.pack(pady=10)
        
        self.upload_button = Button(root, text="上傳文件", command=self.upload_file)
        self.upload_button.pack(pady=5)
        
        self.text_area = Text(root, wrap="word", font=("Arial", 10))
        self.text_area.pack(fill="both", expand=True, padx=10, pady=10)

    def upload_file(self):
        # 開啟文件選擇對話框
        file_path = filedialog.askopenfilename(
            title="選擇Word文件",
            filetypes=[("Word文件", "*.docx *.doc")]
        )
        
        if file_path:
            try:
                # 檢查文件是否存在
                if not os.path.exists(file_path):
                    messagebox.showerror("錯誤", "文件不存在！")
                    return
                
                # 彈出密碼輸入視窗
                password = self.ask_password()
                if password:
                    # 嘗試解密文件
                    content = self.decrypt_word_file(file_path, password)
                    if content:
                        self.text_area.delete(1.0, "end")  # 清除舊內容
                        self.text_area.insert("end", content)  # 顯示新內容
            except Exception as e:
                messagebox.showerror("錯誤", f"處理文件時發生錯誤: {str(e)}")

    def ask_password(self):
        # 密碼輸入對話框
        password_window = Toplevel(self.root)
        password_window.title("輸入密碼")
        password_window.geometry("300x150")
        
        Label(password_window, text="請輸入密碼:").pack(pady=10)
        
        password_entry = Entry(password_window, show="")  # 不隱藏密碼
        password_entry.pack(pady=5)
        
        password = None
        
        def on_ok():
            nonlocal password
            password = password_entry.get()
            password_window.destroy()
        
        Button(password_window, text="確定", command=on_ok).pack(pady=5)
        
        password_window.wait_window()
        return password

    def decrypt_word_file(self, file_path, password):
        try:
            with open(file_path, 'rb') as encrypted_file:
                # 使用 msoffcrypto 解密
                office_file = msoffcrypto.OfficeFile(encrypted_file)
                if not office_file.is_encrypted():
                    messagebox.showinfo("提示", "文件未加密，無需解密。")
                    return None
                
                # 解密文件
                decrypted_content = BytesIO()
                office_file.load_key(password=password)  # 提供密碼
                office_file.decrypt(decrypted_content)
                
                # 使用 python-docx 解析解密後的內容
                doc = Document(decrypted_content)
                content = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
                return content
        except Exception as e:
            messagebox.showerror("錯誤", f"解密失敗，密碼錯誤或文件無法處理: {str(e)}")
            return None

if __name__ == "__main__":
    root = Tk()
    app = WordDecryptorUI(root)
    root.mainloop()
