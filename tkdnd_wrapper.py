class TkDND:
    """TkDND 包裝類，用於在 tkinter 中啟用拖放功能"""
    
    def __init__(self, root):
        """初始化 TkDND 包裝器
        
        參數:
            root: tkinter 或 ttk 的根視窗
        """
        self.root = root
        self.has_tkdnd = False
        
        # 初始化 TkDND 功能
        try:
            # 嘗試載入 tkdnd 套件
            self.root.tk.call('package', 'require', 'tkdnd')
            self.has_tkdnd = True
        except Exception:
            # 如果 tkdnd 作為外部包不可用，則使用內建的
            # Windows 拖放支援（功能有限但適用於基本文件拖放）
            print("TkDND 套件不可用，使用替代方案...")
            
            # 在 Windows 上使用 OLE 拖放
            if hasattr(self.root, 'tk') and self.root.tk.call('tk', 'windowingsystem') == 'win32':
                try:
                    # 啟用 OLE 拖放
                    self.root.tk.call('wm', 'attributes', self.root._w, '-dropoverride', 'true')
                    print("已啟用 Windows OLE 拖放")
                except Exception as e:
                    print(f"無法啟用 Windows OLE 拖放: {str(e)}")

    def bindtarget(self, widget, callback, dndtype):
        """將控件綁定為拖放目標
        
        參數:
            widget: 要綁定的控件
            callback: 拖放事件的回調函數
            dndtype: 拖放類型
        """
        if self.has_tkdnd:
            try:
                # 使用 TkDND 的方法
                widget.drop_target_register(dndtype)
                widget.dnd_bind('<<Drop>>', callback)
                print("已使用 TkDND 註冊拖放目標")
                return True
            except Exception as e:
                print(f"TkDND 註冊失敗: {str(e)}")
        
        # 如果 TkDND 不可用或註冊失敗，使用替代方案
        try:
            # 在 Windows 上使用通用拖放事件
            self.root.bind('<Drop>', callback)
            
            # 如果是 Windows，嘗試啟用 OLE 拖放
            if hasattr(widget, 'tk') and widget.tk.call('tk', 'windowingsystem') == 'win32':
                try:
                    widget.tk.call('wm', 'attributes', widget._w, '-dropoverride', 'true')
                    print("已使用 Windows OLE 拖放註冊目標")
                    return True
                except Exception as e:
                    print(f"Windows OLE 拖放註冊失敗: {str(e)}")
            
            # 嘗試使用通用檔案拖放事件
            widget.bind('<Drop>', callback)
            print("已使用通用拖放事件註冊目標")
            return True
        except Exception as e:
            print(f"所有拖放註冊方法都失敗: {str(e)}")
            return False
