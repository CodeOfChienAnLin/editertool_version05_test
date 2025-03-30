"""
Module for handling typo correction using the OpenCC library with protected words.
"""
import json
import os
import opencc

class TypoCorrector:
    """
    Class to handle typo correction using OpenCC while respecting a protected words list.
    """
    
    def __init__(self, protected_words_file=None):
        """
        Initialize the typo corrector with a protected words list
        
        Args:
            protected_words_file (str, optional): Path to the JSON file containing protected words
        """
        # 初始化OpenCC轉換器（使用正確的配置路徑）
        try:
            # 使用不帶.json後綴的配置名稱
            self.converter_t2s = opencc.OpenCC('t2s')  # 繁體到簡體
            self.converter_s2t = opencc.OpenCC('s2t')  # 簡體到繁體
        except Exception as e:
            print(f"OpenCC初始化錯誤: {e}")
            # 如果初始化失敗，使用空函數作為替代
            self.converter_t2s = self.converter_s2t = lambda x: x
        
        # 載入受保護詞彙列表
        self.protected_words = []
        if protected_words_file and os.path.exists(protected_words_file):
            try:
                with open(protected_words_file, 'r', encoding='utf-8') as f:
                    self.protected_words = json.load(f)
                print(f"已載入 {len(self.protected_words)} 個受保護詞彙")
            except Exception as e:
                print(f"載入受保護詞彙時發生錯誤: {e}")
                # 如果文件不存在或格式錯誤，創建一個空的JSON文件
                self.save_protected_words(protected_words_file)
    
    def add_protected_word(self, word):
        """
        Add a word to the protected words list
        
        Args:
            word (str): Word to add to the protected list
        """
        if word and word not in self.protected_words:
            self.protected_words.append(word)
    
    def remove_protected_word(self, word):
        """
        Remove a word from the protected words list
        
        Args:
            word (str): Word to remove from the protected list
        """
        if word in self.protected_words:
            self.protected_words.remove(word)
    
    def save_protected_words(self, file_path):
        """
        Save the protected words list to a JSON file
        
        Args:
            file_path (str): Path to save the JSON file
        """
        try:
            # 確保目錄存在
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.protected_words, f, ensure_ascii=False, indent=4)
            print(f"受保護詞彙已保存至 {file_path}")
        except Exception as e:
            print(f"保存受保護詞彙時發生錯誤: {e}")
    
    def correct_text(self, text):
        """
        Correct typos in text while respecting protected words
        
        Args:
            text (str): Text to correct
        
        Returns:
            str: Corrected text
        """
        if not text:
            return text
        
        # 將受保護詞彙替換為佔位符
        placeholders = {}
        for i, word in enumerate(self.protected_words):
            if word in text:
                placeholder = f"__PROTECTED_WORD_{i}__"
                text = text.replace(word, placeholder)
                placeholders[placeholder] = word
        
        try:
            # 繁體到簡體再到繁體的轉換（用於糾正錯別字）
            corrected_text = text
            if hasattr(self.converter_t2s, 'convert') and hasattr(self.converter_s2t, 'convert'):
                simplified = self.converter_t2s.convert(text)
                corrected_text = self.converter_s2t.convert(simplified)
        except Exception as e:
            print(f"轉換過程中發生錯誤: {e}")
            corrected_text = text
        
        # 恢復受保護詞彙
        for placeholder, word in placeholders.items():
            corrected_text = corrected_text.replace(placeholder, word)
        
        return corrected_text
