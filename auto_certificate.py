import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import datetime, os, json, time
import win32com.client

class CertificateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("全方位獎狀生成器 - 偵錯選取版")
        self.root.geometry("1100x950")
        
        self.templates = []
        self.current_index = -1
        
        self.setup_ui()
        self.add_new_template_item()

    def setup_ui(self):
        self.paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4)
        self.paned.pack(fill="both", expand=True, padx=10, pady=10)
        
        # --- 左側：帶有勾選功能的清單 ---
        left_frame = tk.LabelFrame(self.paned, text="1. 選擇執行模板 (勾選欲執行的項目)", font=("Microsoft JhengHei", 10, "bold"))
        self.paned.add(left_frame, width=320)
        
        # 使用 Treeview 來模擬 Checkbox 清單
        columns = ("run", "name")
        self.tree_list = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="browse")
        self.tree_list.heading("run", text="執行", anchor="center")
        self.tree_list.heading("name", text="模板名稱", anchor="w")
        self.tree_list.column("run", width=40, anchor="center")
        self.tree_list.column("name", width=200)
        self.tree_list.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 綁定點擊事件（模擬 Checkbox）
        self.tree_list.bind('<ButtonRelease-1>', self.on_tree_click)
        
        btn_f = tk.Frame(left_frame)
        btn_f.pack(fill="x", padx=5)
        tk.Button(btn_f, text="+ 新增", command=self.add_new_template_item).pack(side="left", fill="x", expand=True)
        tk.Button(btn_f, text="- 刪除", command=self.delete_template_item).pack(side="left", fill="x", expand=True)
        tk.Button(btn_f, text="全選/全不選", command=self.toggle_all_selection).pack(side="left", fill="x", expand=True)
        
        cfg_f = tk.Frame(left_frame)
        cfg_f.pack(fill="x", padx=5, pady=5)
        tk.Button(cfg_f, text="📂 匯入 Config", command=self.load_config).pack(side="left", fill="x", expand=True)
        tk.Button(cfg_f, text="💾 儲存 Config", command=self.save_config, bg="#fff9c4").pack(side="left", fill="x", expand=True)

        # --- 右側：編輯設定 ---
        self.right_frame = tk.LabelFrame(self.paned, text="2. 編輯選定模板細節", padx=15, pady=15)
        self.paned.add(self.right_frame)
        
        self.temp_name_var = tk.StringVar()
        self.temp_name_var.trace_add("write", self.sync_name_to_tree)
        self.ppt_path, self.excel_path = tk.StringVar(), tk.StringVar()

        tk.Label(self.right_frame, text="名稱:").grid(row=0, column=0, sticky="e")
        tk.Entry(self.right_frame, textvariable=self.temp_name_var, width=50).grid(row=0, column=1, columnspan=2, sticky="w")
        tk.Label(self.right_frame, text="PPT:").grid(row=1, column=0, sticky="e")
        tk.Entry(self.right_frame, textvariable=self.ppt_path, width=50).grid(row=1, column=1, padx=5)
        tk.Button(self.right_frame, text="瀏覽", command=lambda: self.ppt_path.set(filedialog.askopenfilename())).grid(row=1, column=2)
        tk.Label(self.right_frame, text="Excel:").grid(row=2, column=0, sticky="e")
        tk.Entry(self.right_frame, textvariable=self.excel_path, width=50).grid(row=2, column=1, padx=5)
        tk.Button(self.right_frame, text="瀏覽", command=lambda: self.excel_path.set(filedialog.askopenfilename())).grid(row=2, column=2)

        self.tree_rules = ttk.Treeview(self.right_frame, columns=("c", "t"), show="headings", height=10)
        self.tree_rules.heading("c", text="Excel 欄位"); self.tree_rules.heading("t", text="PPT 標籤")
        self.tree_rules.grid(row=4, column=0, columnspan=3, sticky="nsew")

        edit_f = tk.Frame(self.right_frame)
        edit_f.grid(row=5, column=0, columnspan=3, pady=10)
        self.ent_col = tk.Entry(edit_f, width=15); self.ent_col.pack(side="left", padx=2)
        self.ent_tag = tk.Entry(edit_f, width=15); self.ent_tag.pack(side="left", padx=2)
        tk.Button(edit_f, text="加入規則", command=self.add_rule, bg="#007bff", fg="white").pack(side="left", padx=5)
        tk.Button(edit_f, text="刪除選取", command=self.delete_rule).pack(side="left")

        self.log_area = scrolledtext.ScrolledText(self.root, height=12, bg="#1e1e1e", fg="#ffffff", font=("Consolas", 10))
        self.log_area.pack(fill="both", padx=20, pady=5)
        tk.Button(self.root, text="🚀 啟動勾選項目之批次任務", command=self.run_batch_process, bg="#28a745", fg="white", font=("", 12, "bold"), height=2).pack(fill="x", padx=100, pady=10)

    # --- UI 控制邏輯 ---
    def on_tree_click(self, event):
        item_id = self.tree_list.identify_row(event.y)
        column = self.tree_list.identify_column(event.x)
        
        if not item_id: return
        
        idx = self.tree_list.index(item_id)
        
        # 如果點擊的是 "執行" 這一欄 (第#1欄)，切換打勾狀態
        if column == "#1":
            current_status = self.templates[idx].get("is_active", True)
            self.templates[idx]["is_active"] = not current_status
            self.update_tree_list()
        
        # 切換編輯對象
        if self.current_index != -1: self.save_to_mem()
        self.current_index = idx
        self.load_to_ui(self.templates[idx])

    def update_tree_list(self):
        """重新整理左側清單顯示"""
        self.tree_list.delete(*self.tree_list.get_children())
        for t in self.templates:
            status = " [ v ] " if t.get("is_active", True) else " [   ] "
            self.tree_list.insert("", "end", values=(status, t["name"]))

    def toggle_all_selection(self):
        if not self.templates: return
        # 只要有一個沒勾，就全選；否則全取消
        any_unselected = any(not t.get("is_active", True) for t in self.templates)
        for t in self.templates:
            t["is_active"] = any_unselected
        self.update_tree_list()

    def sync_name_to_tree(self, *args):
        if self.current_index != -1:
            new_name = self.temp_name_var.get()
            self.templates[self.current_index]["name"] = new_name
            # 更新特定行
            item_id = self.tree_list.get_children()[self.current_index]
            status = " [ v ] " if self.templates[self.current_index].get("is_active", True) else " [   ] "
            self.tree_list.item(item_id, values=(status, new_name))

    def load_to_ui(self, d):
        self.temp_name_var.set(d["name"])
        self.ppt_path.set(d.get("ppt", ""))
        self.excel_path.set(d.get("excel", ""))
        self.tree_rules.delete(*self.tree_rules.get_children())
        for r in d.get("rules", []):
            self.tree_rules.insert("", "end", values=r)

    def save_to_mem(self):
        if self.current_index == -1: return
        self.templates[self.current_index].update({
            "ppt": self.ppt_path.get(),
            "excel": self.excel_path.get(),
            "rules": [self.tree_rules.item(i)["values"] for i in self.tree_rules.get_children()]
        })

    def add_new_template_item(self):
        if self.current_index != -1: self.save_to_mem()
        new_item = {"name": f"新模板 {len(self.templates)+1}", "ppt": "", "excel": "", "rules": [], "is_active": True}
        self.templates.append(new_item)
        self.update_tree_list()
        self.current_index = len(self.templates) - 1
        self.load_to_ui(new_item)

    def delete_template_item(self):
        if self.current_index == -1: return
        if messagebox.askyesno("刪除", "確定刪除選定模板？"):
            del self.templates[self.current_index]
            self.current_index = -1
            self.update_tree_list()
            self.temp_name_var.set(""); self.ppt_path.set(""); self.excel_path.set(""); self.tree_rules.delete(*self.tree_rules.get_children())

    def add_rule(self):
        c, t = self.ent_col.get(), self.ent_tag.get()
        if c and t:
            self.tree_rules.insert("", "end", values=(c, t))
            self.ent_col.delete(0, tk.END); self.ent_tag.delete(0, tk.END); self.save_to_mem()

    def delete_rule(self):
        for s in self.tree_rules.selection(): self.tree_rules.delete(s)
        self.save_to_mem()

    def save_config(self):
        self.save_to_mem()
        f = filedialog.asksaveasfilename(defaultextension=".json")
        if f: 
            with open(f, 'w', encoding='utf-8') as jf: json.dump(self.templates, jf, indent=4)
            messagebox.showinfo("成功", "存檔完成")

    def load_config(self):
        f = filedialog.askopenfilename()
        if f:
            with open(f, 'r', encoding='utf-8') as jf: self.templates = json.load(jf)
            self.update_tree_list()
            self.current_index = -1

    # --- 執行引擎 ---
    def run_batch_process(self):
        self.save_to_mem()
        active_templates = [t for t in self.templates if t.get("is_active", True)]
        
        if not active_templates:
            messagebox.showwarning("提示", "請至少勾選一個模板進行執行！")
            return

        self.log_area.delete("1.0", tk.END)
        self.log(f">>> [系統] 啟動批次任務，預計執行 {len(active_templates)} 組模板...")
        
        ppt_app = None
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            for t_data in active_templates:
                if not t_data["ppt"] or not t_data["excel"]:
                    self.log(f"[跳過] 「{t_data['name']}」路徑未設定。")
                    continue
                self.log(f"--- 正在處理: {t_data['name']} ---")
                self.process_core(ppt_app, t_data)
            
            messagebox.showinfo("完成", "勾選項目執行結束")
        except Exception as e:
            self.log(f"[致命錯誤] {e}")
        finally:
            if ppt_app: ppt_app.Quit()

    def process_core(self, ppt_app, data):
        try:
            # --- 1. 讀取並防呆處理 Excel 資料 ---
            # 讀取 Excel 並先移除完全空白的列
            df = pd.read_excel(data["excel"]).dropna(how='all').fillna("")
            
            # 【防呆 1】強制清除 Excel 欄位名稱前後的空白與換行符號
            df.columns = [str(c).strip() for c in df.columns]
            self.log(f"  [Debug] Excel 成功偵測到欄位: {list(df.columns)}")

            task_list = []
            for _, row in df.iterrows():
                # 【防呆 2】在匹配時，也把規則清單裡的 col 與 tag 去掉前後空白
                item = {}
                for col_raw, tag_raw in data["rules"]:
                    col_clean = str(col_raw).strip()
                    tag_clean = str(tag_raw).strip()
                    
                    if col_clean in df.columns:
                        item[tag_clean] = str(row[col_clean])
                    else:
                        # 如果規則設了但 Excel 找不到，輸出一行提醒
                        self.log(f"  [提醒] 規則中的欄位 '{col_clean}' 在 Excel 中找不到，請檢查名稱。")
                
                # 確保這一列至少有一個標籤有抓到資料，才加入任務
                if any(item.values()): 
                    task_list.append(item)
            
            if not task_list:
                self.log(f"  [警告] 模板「{data['name']}」無有效資料。請確認規則中的 Excel 欄位名稱是否正確。")
                return

            self.log(f"  [確認] 有效待處理資料: {len(task_list)} 筆")

            # --- 2. 開啟 PPT 並進行生成 ---
            pres = ppt_app.Presentations.Open(os.path.abspath(data["ppt"]))
            original_count = pres.Slides.Count
            
            for i, reps in enumerate(task_list):
                self.log(f"  [製作中] 第 {i+1}/{len(task_list)} 筆資料...")
                start_index = pres.Slides.Count + 1
                
                # 複製模板
                for s_idx in range(1, original_count + 1):
                    pasted = False
                    for retry in range(5):  # 剪貼簿重試機制
                        try:
                            pres.Slides(s_idx).Copy()
                            time.sleep(0.3)
                            pres.Slides.Paste(pres.Slides.Count + 1)
                            pasted = True
                            break 
                        except Exception:
                            self.log(f"    ! 剪貼簿衝突，正在重試第 {retry+1} 次...")
                            time.sleep(0.8)
                    
                    if not pasted:
                        raise Exception("系統剪貼簿遭佔用（可能是開啟了 LINE、剪貼簿管理員等），請關閉後重試。")

                # 精準替換內容
                for offset in range(original_count):
                    target_slide = pres.Slides(start_index + offset)
                    for shape in target_slide.Shapes:
                        if shape.HasTextFrame:
                            tr = shape.TextFrame.TextRange
                            for tag, val in reps.items():
                                tr.Replace(tag, val)
                        if shape.HasTable:
                            for row_t in shape.Table.Rows:
                                for cell in row_t.Cells:
                                    cr = cell.Shape.TextFrame.TextRange
                                    for tag, val in reps.items():
                                        cr.Replace(tag, val)

            # --- 3. 清理並存檔 ---
            self.log("  - 正在移除原始模板頁面...")
            for _ in range(original_count):
                pres.Slides(1).Delete()

            save_fn = f"Result_{data['name']}_{datetime.datetime.now().strftime('%H%M%S')}.pptx"
            pres.SaveAs(os.path.abspath(save_fn))
            pres.Close()
            self.log(f"  [成功] 檔案已產出至: {save_fn}\n")
            
        except Exception as e:
            self.log(f"  [失敗] 處理過程中發生錯誤: {e}\n")

    def log(self, msg):
        self.log_area.insert(tk.END, f"{msg}\n")
        self.log_area.see(tk.END); self.root.update()

if __name__ == "__main__":
    root = tk.Tk(); app = CertificateApp(root); root.mainloop()