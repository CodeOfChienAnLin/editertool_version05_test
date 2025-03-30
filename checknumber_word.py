
import os
import msoffcrypto
from io import BytesIO
from tkinter import Tk, Label, Button, Text, filedialog, messagebox, Toplevel, Entry
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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
                    # 呼叫解密函數
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
        """
        解密加密的 Word 文件並返回內容（包含編號）
        :param file_path: Word 文件路徑
        :param password: 解密密碼
        :return: 解密後的文字內容（若成功），否則返回 None
        """
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
                
                # 解析文件並保留編號
                return self.parse_word_file(decrypted_content)
        except Exception as e:
            messagebox.showerror("錯誤", f"解密失敗，密碼錯誤或文件無法處理: {str(e)}")
            return None

    def parse_word_file(self, file_stream):
        """
        解析 Word 文件，保留編號信息
        :param file_stream: 解密後的文件流
        :return: 包含編號的完整內容
        """
        doc = Document(file_stream)
        content = []
        
        for para in doc.paragraphs:
            # 提取段落中的編號信息
            numbering = self.extract_numbering(para)
            if numbering:
                # 將編號與段落內容結合
                content.append(f"{numbering} {para.text}")
            else:
                content.append(para.text)
        
        return "\n".join(content)

    def extract_numbering(self, paragraph):
        """
        從段落中提取編號信息
        :param paragraph: 段落對象
        :return: 編號字串（若存在），否則返回 None
        """
        p = paragraph._element
        num_pr = p.find(qn('w:numPr'))
        if num_pr is not None:
            # 提取編號ID和層級
            num_id = num_pr.find(qn('w:numId')).attrib.get(qn('w:val'))
            level = num_pr.find(qn('w:ilvl')).attrib.get(qn('w:val'))
            
            # 查找編號定義
            num_def = self.find_numbering_definition(paragraph.part, num_id)
            if num_def:
                # 提取編號格式
                num_text = self.extract_number_text(num_def, level)
                return num_text
        return None

    def find_numbering_definition(self, part, num_id):
        """
        查找編號定義
        :param part: 文件部分
        :param num_id: 編號ID
        :return: 編號定義的XML元素
        """
        numbering_part = part.numbering_part
        if numbering_part:
            num_def = numbering_part.element.find(qn(f'w:num[@w:numId="{num_id}"]'))
            return num_def
        return None

    def extract_number_text(self, num_def, level):
        """
        從編號定義中提取編號文本
        :param num_def: 編號定義的XML元素
        :param level: 編號層級
        :return: 編號字串
        """
        lvl = num_def.find(qn(f'w:lvl[@w:ilvl="{level}"]'))
        if lvl is not None:
            num_fmt = lvl.find(qn('w:numFmt'))
            if num_fmt is not None:
                num_fmt_val = num_fmt.attrib.get(qn('w:val'))
                if num_fmt_val == 'decimal':
                    return "1."  # 示例：返回數字編號
                elif num_fmt_val == 'upperLetter':
                    return "A."  # 示例：返回大寫字母編號
                elif num_fmt_val == 'lowerLetter':
                    return "a."  # 示例：返回小寫字母編號
                elif num_fmt_val == 'upperRoman':
                    return "I."  # 示例：返回大寫羅馬數字編號
                elif num_fmt_val == 'lowerRoman':
                    return "i."  # 示例：返回小寫羅馬數字編號
                # 其他格式可根據需求擴展
        return None

if __name__ == "__main__":
    root = Tk()
    app = WordDecryptorUI(root)
    root.mainloop()
