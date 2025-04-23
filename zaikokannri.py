import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
import pandas as pd
import cv2
from pyzbar.pyzbar import decode
import qrcode
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

env_loaded = load_dotenv()
if not env_loaded:
    print(".envファイルの読み込みに失敗しました。デフォルト値を使用します。")

class CenteredAskString(simpledialog.Dialog):
    def __init__(self, parent, title, prompt):
        self.prompt = prompt
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text=self.prompt).pack(padx=10, pady=10)
        self.entry = tk.Entry(master)
        self.entry.pack(padx=10, pady=10)
        return self.entry

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f'{width}x{height}+{x}+{y}')

    def apply(self):
        self.result = self.entry.get().strip()

    def show(self):
        # マッピング後に中央配置を呼び出す
        self.after(0, self.center_window)
        self.wait_window()
        return self.result

def ask_centered_string(parent, title, prompt):
    dlg = CenteredAskString(parent, title, prompt)
    return dlg.result

def send_low_stock_email_no_oauth(low_stock_items, sender_email_default, sender_password_default, recipient_email):
    # 環境変数から認証情報取得（設定されていない場合はデフォルト値を利用）
    sender_email = os.getenv("GMAIL_USER", sender_email_default)
    sender_password = os.getenv("GMAIL_APP_PASSWORD", sender_password_default)
    
    subject = "在庫不足通知 (SMTP - App Password)"
    body = "以下の商品で在庫が不足しています:\n" + "\n".join([
        f"{item['name']} (在庫: {0 if item.get('quantity') is None else item.get('quantity')})"
        for item in low_stock_items
    ])

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=5)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print("在庫不足通知メールを送信しました。")
    except Exception as e:
        print("メール送信に失敗しました:", e)

def ask_integer_modal(parent, title, prompt, minvalue=1):
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.transient(parent)
    dialog.grab_set()       # モーダルにする
    dialog.focus_force()    # 最前面に表示

    # 画面中央に配置する計算
    dialog.update_idletasks()
    w = dialog.winfo_width()
    h = dialog.winfo_height()
    ws = dialog.winfo_screenwidth()
    hs = dialog.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    dialog.geometry(f'+{x}+{y}')

    tk.Label(dialog, text=prompt).pack(padx=10, pady=10)

    # 初期値は空白に設定
    var = tk.StringVar(value="")
    entry = tk.Entry(dialog, textvariable=var)
    entry.pack(padx=10, pady=10)
    entry.focus_force()

    result = []
    def on_ok(event=None):
        entered = var.get().strip()
        if entered == "":
            messagebox.showerror("入力エラー", "値を入力してください。", parent=dialog)
            return
        try:
            value = int(entered)
            if value < minvalue:
                messagebox.showerror("入力エラー", f"{minvalue}以上の整数を入力してください。", parent=dialog)
                return
            result.append(value)
        except ValueError:
            messagebox.showerror("入力エラー", "整数値を入力してください。", parent=dialog)
            return
        dialog.destroy()

    # ボタンを横並びに配置するためのフレーム
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=5)
    tk.Button(btn_frame, text="OK", command=on_ok).pack(side="left", padx=5)
    tk.Button(btn_frame, text="キャンセル", command=dialog.destroy).pack(side="left", padx=5)
    # EnterキーでOKを実行
    dialog.bind("<Return>", on_ok)

    parent.wait_window(dialog)
    return result[0] if result else None

class InventoryApp:
    EXCEL_FILE = r"C:\Users\ksuzuki4\Desktop\台帳.xlsx"
    LOW_STOCK_THRESHOLD = 5

    def __init__(self, root):
        self.root = root
        self.root.title("在庫管理アプリ")
        self.root.geometry("1000x400")
        
        # 認証情報およびメール設定を環境変数から取得（未設定の場合はデフォルト値を設定）
        self.admin_password = os.getenv("ADMIN_PASSWORD", "default_admin")
        self.sender_email = os.getenv("GMAIL_USER", "default_sender@example.com")
        self.sender_password = os.getenv("GMAIL_APP_PASSWORD", "default_app_password")
        self.recipient_email = os.getenv("RECIPIENT_EMAIL", "default_recipient@example.com")
        
        if not os.path.exists(self.EXCEL_FILE):
            messagebox.showerror("読み込みエラー", f"指定したExcelファイルが存在しません: {self.EXCEL_FILE}")
            self.root.destroy()
            return

        try:
            df = pd.read_excel(self.EXCEL_FILE)
            self.inventory_data = df.to_dict("records")
            for item in self.inventory_data:
                if "threshold" not in item or pd.isna(item["threshold"]):
                    item["threshold"] = 5
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"Excelファイルの読み込みに失敗しました: {e}")
            self.root.destroy()
            return

        self.cancel_qr = False

        # 追加: フィルタ用変数を初期化
        self.category_vars = {}
        self.location_vars = {}

        # --- レイアウト変更開始 ---
        # 左：在庫一覧(Treeview)用フレーム（スクロールバー追加）
        self.tree_frame = tk.Frame(root)
        self.tree_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.inventory_tree = ttk.Treeview(self.tree_frame, 
            columns=("id", "name", "category", "quantity", "location", "threshold"), 
            show="headings", height=15)
        self.inventory_tree.heading("id", text="ID")
        self.inventory_tree.heading("name", text="商品名")
        self.inventory_tree.heading("category", text="カテゴリ")
        self.inventory_tree.heading("quantity", text="数量")
        self.inventory_tree.heading("location", text="保管場所")
        self.inventory_tree.heading("threshold", text="閾値")
        self.inventory_tree.column("id", width=50, anchor="center")
        self.inventory_tree.column("name", width=150)
        self.inventory_tree.column("category", width=100)
        self.inventory_tree.column("quantity", width=50, anchor="center")
        self.inventory_tree.column("location", width=150)
        self.inventory_tree.column("threshold", width=50, anchor="center")
        self.inventory_tree.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.inventory_tree.yview)
        self.inventory_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # 右：フィルター用コンテナフレーム（全フィルター解除ボタンとフィルタ群）
        self.filter_container = tk.Frame(root)
        self.filter_container.grid(row=0, column=1, padx=10, pady=10, sticky="n")
        # フィルタボタン（全フィルター解除のみ）を配置
        filter_button_frame = tk.Frame(self.filter_container)
        filter_button_frame.pack(anchor="w", pady=(0,10))
        tk.Button(filter_button_frame, text="フィルタ解除", width=15, command=self.clear_filters)\
                .pack(side="left", padx=5)
        # フィルタ一覧を横並びに配置するためのフレーム
        filters_frame = tk.Frame(self.filter_container)
        filters_frame.pack(fill="x", expand=True)
        
        # カテゴリフィルタ（上端揃え）
        self.cat_filter_frame = tk.Frame(filters_frame)
        self.cat_filter_frame.pack(side="left", padx=5, anchor="n")
        tk.Label(self.cat_filter_frame, text="【カテゴリ】", font=("Helvetica", 10, "bold"))\
            .pack(anchor="w", pady=(0,5))
        self.update_category_checkboxes(in_frame_only=True)
        
        # 保管場所フィルタ（上端揃え）
        self.loc_filter_frame = tk.Frame(filters_frame)
        self.loc_filter_frame.pack(side="left", padx=5, anchor="n")
        tk.Label(self.loc_filter_frame, text="【保管場所】", font=("Helvetica", 10, "bold"))\
            .pack(anchor="w", pady=(0,5))
        self.update_location_checkboxes()
        # --- レイアウト変更終了 ---
        
        # 下部：機能ボタン配置（左詰め）
        self.button_frame = tk.Frame(root)
        self.button_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        self.create_buttons()
        self.update_inventory_display()
        
    def show_all_items(self):
        """全表示ボタン用：フィルターを無視してすべて表示"""
        for var in self.category_vars.values():
            var.set(0)
        for var in self.location_vars.values():
            var.set(0)
        self.update_inventory_display()

    def update_inventory_display(self):
        """Treeviewの内容をクリアし、フィルタに応じた在庫表示を更新"""
        self.filtered_inventory = []
        for row in self.inventory_tree.get_children():
            self.inventory_tree.delete(row)

        # 選択中のフィルター条件を取得
        selected_categories = [cat for cat, var in self.category_vars.items() if var.get() == 1]
        selected_locations = [loc for loc, var in self.location_vars.items() if var.get() == 1]

        for item in self.inventory_data:
            # itemのカテゴリと保管場所（空の場合は"未設定"）
            item_cat = str(item.get("category") or "未設定")
            item_loc = str(item.get("location") or "未設定")

            if selected_categories and (item_cat not in selected_categories):
                continue

            if selected_locations and (item_loc not in selected_locations):
                continue

            try:
                quantity = 0 if pd.isna(item['quantity']) else int(item['quantity'])
            except Exception:
                quantity = 0
            threshold = item.get("threshold")
            if threshold is None or pd.isna(threshold):
                threshold = "未設定"
            # 発注中なら商品名の前に【発注中】を表示
            name_to_show = item["name"]
            if item.get("order_pending", False):
                name_to_show = "【発注中】" + name_to_show

            self.filtered_inventory.append(item)
            self.inventory_tree.insert("", "end", values=(
                item["id"], name_to_show, item["category"], quantity, item_loc, threshold))

    def update_category_checkboxes(self, in_frame_only=False):
        # 既存のウィジェットをクリア
        for widget in self.cat_filter_frame.winfo_children():
            widget.destroy()
        # ラベルを表示
        tk.Label(self.cat_filter_frame, text="【カテゴリ】", font=("Helvetica", 10, "bold"))\
            .pack(anchor="w", pady=(0,5))
            
        # キャンバスとスクロールバーを作成（表示行数は高さで調整、ここでは150pxとして例示）
        canvas = tk.Canvas(self.cat_filter_frame, width=150, height=150)
        scrollbar = tk.Scrollbar(self.cat_filter_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.category_vars.clear()
        cats = {str(item.get("category") or "未設定") for item in self.inventory_data}
        for cat in sorted(cats):
            var = tk.IntVar(value=0)
            self.category_vars[cat] = var
            tk.Checkbutton(scrollable_frame, text=cat, variable=var, command=self.on_filter_change)\
                .pack(anchor="w")

    def update_location_checkboxes(self):
        # loc_filter_frame内のウィジェットをすべてクリア
        for widget in self.loc_filter_frame.winfo_children():
            widget.destroy()
        # ラベルを表示
        tk.Label(self.loc_filter_frame, text="【保管場所】", font=("Helvetica", 10, "bold"))\
            .pack(anchor="w", pady=(0,5))
            
        # キャンバスとスクロールバーを作成（表示行数は高さで調整、ここでは150pxとして例示）
        canvas = tk.Canvas(self.loc_filter_frame, width=150, height=150)
        scrollbar = tk.Scrollbar(self.loc_filter_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.location_vars.clear()
        locs = {str(item.get("location") or "未設定") for item in self.inventory_data}
        for loc in sorted(locs):
            var = tk.IntVar(value=0)
            self.location_vars[loc] = var
            tk.Checkbutton(scrollable_frame, text=loc, variable=var, command=self.on_filter_change)\
                .pack(anchor="w")

    def on_filter_change(self):
        """カテゴリまたは保管場所チェックボックス変更時の処理"""
        self.update_inventory_display()

    def clear_filters(self):
        """全フィルター解除"""
        for var in self.category_vars.values():
            var.set(0)
        for var in self.location_vars.values():
            var.set(0)
        self.update_inventory_display()

    def cancel_qr_button(self, cancel_window):
        """QRコード読み取り時のキャンセル処理"""
        self.cancel_qr = True
        cancel_window.destroy()

    def read_qr_code(self):
        """カメラからQRコードを読み取る。キャンセルボタンで中断可能"""
        cancel_window = tk.Toplevel(self.root)
        cancel_window.title("QRコード読み取り中")
        # cancel_window.attributes("-topmost", True)
        # cancel_window.grab_set()
        # cancel_window.focus_force()
        tk.Label(cancel_window, text="QRコード読み取り中です…").pack(padx=10, pady=10)
        tk.Button(cancel_window, text="キャンセル", command=lambda: self.cancel_qr_button(cancel_window)).pack(pady=10)

        cap = cv2.VideoCapture(1)
        qr_result = None
        self.cancel_qr = False

        while True:
            if self.cancel_qr:
                break

            ret, frame = cap.read()
            if not ret:
                break

            cv2.putText(frame, "EXIT Esc or q", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2, cv2.LINE_AA)

            for obj in decode(frame):
                qr_result = obj.data.decode("utf-8")
                messagebox.showinfo("QRコード読み取り", f"QRコードデータ: {qr_result}")
                self.cancel_qr = True
                break

            cv2.imshow("QRコード読み取り", frame)
            if cv2.waitKey(1) & 0xFF in [27, ord('q')]:
                self.cancel_qr = True
                break

        cap.release()
        cv2.destroyAllWindows()

        if cancel_window.winfo_exists():
            cancel_window.destroy()

        return qr_result

    def create_qr_code(self):
        selected = self.inventory_tree.selection()
        if not selected:
            messagebox.showwarning("QRコード生成", "生成する品目をリストから選択してください。")
            return
        item_values = self.inventory_tree.item(selected[0], "values")
        selected_id = item_values[0]
        selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(selected_id)), None)
        if not selected_item:
            messagebox.showerror("QRコード生成エラー", f"選択された品番が見つかりません: {selected_id}")
            return

        data = f"ID: {selected_item['id']}, 商品名: {selected_item['name']}, カテゴリ: {selected_item['category']}, 保管場所: {selected_item.get('location', '未設定')}"
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=2,  # 1セルのサイズ
            border=2     # 周囲の余白
        )
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        default_filename = f"{selected_item['id']}_{selected_item['name']}.png"
        save_path = filedialog.asksaveasfilename(initialfile=default_filename,
                                                 defaultextension=".png",
                                                 filetypes=[("PNG Files", "*.png")])
        if save_path:
            img.save(save_path)
            messagebox.showinfo("QRコード生成", f"QRコード画像を保存しました: {save_path}")

    def import_csv(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")])
        if not filepath:
            return
        try:
            # ファイル拡張子によって読み込み方法を切り替える
            if filepath.lower().endswith(('.xlsx', '.xls')):
                data = pd.read_excel(filepath)
            else:
                data = pd.read_csv(filepath)
            
            required = ['id', 'name', 'category', 'quantity', 'location', 'threshold', 'order_pending']
            missing = [col for col in required if col not in data.columns]
            if missing:
                messagebox.showerror("CSVエラー", f"次の列が不足しています: {', '.join(missing)}")
                return
            
            for _, row in data.iterrows():
                self.inventory_data.append({
                    "id": row['id'],
                    "name": row['name'],
                    "category": row['category'],
                    "quantity": row['quantity'],
                    "location": row['location'],
                    "threshold": row['threshold'],
                    # order_pending 列が存在するかチェックし、欠損値の場合は False を設定
                    "order_pending": row['order_pending'] if not pd.isna(row.get('order_pending', False)) else False
                })
            messagebox.showinfo("CSVインポート", "CSV/Excelファイルのインポートが成功しました！")
            self.update_inventory_display()
            self.update_category_checkboxes()
            self.update_location_checkboxes()
            self.save_inventory_to_excel()
        except Exception as e:
            messagebox.showerror("CSVインポートエラー", f"エラーが発生しました: {e}")

    def stock_in(self):
        choice = messagebox.askquestion("入庫方法選択", 
                                        "QRコードで入庫しますか？\n「いいえ」を選択すると、リスト選択またはID入力が可能です。")
        if choice == "yes":
            qr_data = self.read_qr_code()
            if not qr_data:
                return messagebox.showwarning("QRコードエラー", "QRコードの読み取りに失敗しました。")
            selected_item = next((item for item in self.inventory_data if str(item["id"]) in qr_data), None)
            if not selected_item:
                return messagebox.showerror("品番エラー", "QRコードに対応する品番が見つかりません。")
        else:
            selected = self.inventory_tree.selection()
            if selected:
                item_values = self.inventory_tree.item(selected[0], "values")
                selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(item_values[0])), None)
            else:
                # 直接手動入力に切り替える
                entered_id = ask_centered_string(self.root, "ID入力", "入庫する商品のIDを入力してください:")
                if not entered_id:
                    return
                selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(entered_id)), None)
                if not selected_item:
                    return messagebox.showerror("品番エラー", "入力されたIDに対応する品番が見つかりません。")

        add_qty = ask_integer_modal(self.root, "入庫", f"{selected_item['name']} の入庫数量を入力してください", minvalue=1)
        if add_qty is None:
            return

        ok = messagebox.askyesno("確認", f"{selected_item['name']} を {add_qty} 個 入庫します。よろしいですか？")
        if not ok:
            return

        try:
            current_qty = int(selected_item.get("quantity", 0))
            selected_item["quantity"] = current_qty + add_qty
            if selected_item.get("order_pending", False):
                selected_item["order_pending"] = False
        except Exception:
            selected_item["quantity"] = add_qty

        self.update_inventory_display()
        self.save_inventory_to_excel()
        messagebox.showinfo("入庫完了", f"{selected_item['name']} を {add_qty} 個 入庫しました。")
        self.record_log("入庫", selected_item, add_qty)

    def stock_out(self):
        choice = messagebox.askquestion("出庫方法選択", 
                                        "QRコードで出庫しますか？\n「いいえ」を選択すると、リスト選択またはID入力が可能です。")
        if choice == "yes":
            qr_data = self.read_qr_code()
            if not qr_data:
                return messagebox.showwarning("QRコードエラー", "QRコードの読み取りに失敗しました。")
            selected_item = next((item for item in self.inventory_data if str(item["id"]) in qr_data), None)
            if not selected_item:
                return messagebox.showerror("品番エラー", "QRコードに対応する品番が見つかりません。")
        else:
            selected = self.inventory_tree.selection()
            if selected:
                item_values = self.inventory_tree.item(selected[0], "values")
                selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(item_values[0])), None)
            else:
                # 直接手動入力に切り替える
                entered_id = ask_centered_string(self.root, "ID入力", "出庫する商品のIDを入力してください:")
                if not entered_id:
                    return
                selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(entered_id)), None)
                if not selected_item:
                    return messagebox.showerror("品番エラー", "入力されたIDに対応する品番が見つかりません。")

        remove_qty = simpledialog.askinteger("出庫", f"{selected_item['name']} の出庫数量を入力してください", parent=self.root, minvalue=1)
        if remove_qty is None:
            return

        try:
            current_qty = int(selected_item.get("quantity", 0))
        except Exception:
            current_qty = 0
        if remove_qty > current_qty:
            return messagebox.showerror("数量エラー", "出庫数量が在庫数量を超えています。")

        selected_item["quantity"] = current_qty - remove_qty
        self.update_inventory_display()
        self.save_inventory_to_excel()
        self.check_low_stock()
        messagebox.showinfo("出庫完了", f"{selected_item['name']} を {remove_qty} 個 出庫しました。")
        self.record_log("出庫", selected_item, -remove_qty)

    def save_inventory_to_excel(self):
        try:
            pd.DataFrame(self.inventory_data).to_excel(self.EXCEL_FILE, index=False)
            # messagebox.showinfo("Excel保存", f"在庫台帳がExcelファイルに保存されました: {self.EXCEL_FILE}")
        except Exception as e:
            messagebox.showerror("Excel保存エラー", f"Excel保存に失敗しました: {e}")

    def register_new_product(self):
        top = tk.Toplevel(self.root)
        top.title("新規品登録")
        top.geometry("350x300")

        tk.Label(top, text="商品ID:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        entry_id = tk.Entry(top, width=25)
        entry_id.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(top, text="商品名称:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry_name = tk.Entry(top, width=25)
        entry_name.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(top, text="カテゴリ:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        entry_category = tk.Entry(top, width=25)
        entry_category.grid(row=2, column=1, padx=10, pady=5)

        tk.Label(top, text="数量:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        entry_quantity = tk.Entry(top, width=25)
        entry_quantity.grid(row=3, column=1, padx=10, pady=5)

        tk.Label(top, text="保管場所:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        entry_location = tk.Entry(top, width=25)
        entry_location.grid(row=4, column=1, padx=10, pady=5)

        tk.Label(top, text="閾値:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        entry_threshold = tk.Entry(top, width=25)
        entry_threshold.grid(row=5, column=1, padx=10, pady=5)

        def submit():
            product_id = entry_id.get().strip()
            name = entry_name.get().strip()
            category = entry_category.get().strip()
            quantity_str = entry_quantity.get().strip()
            location = entry_location.get().strip()
            threshold_str = entry_threshold.get().strip()
            if not all([product_id, name, category, quantity_str, threshold_str]):
                messagebox.showwarning("入力エラー", "すべての必須項目（ID、名称、カテゴリ、数量、閾値）を入力してください。")
                return
            for prod in self.inventory_data:
                if str(prod.get("id")).strip() == product_id:
                    messagebox.showwarning("入力エラー", "すでに同じIDが存在します。")
                    return
            try:
                quantity = int(quantity_str)
                threshold = int(threshold_str)
                if quantity < 0 or threshold < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("数量エラー", "数量と閾値は0以上の整数で入力してください。")
                return
            new_product = {
                "id": product_id,
                "name": name,
                "category": category,
                "quantity": quantity,
                "location": location,
                "threshold": threshold
            }
            self.inventory_data.append(new_product)
            self.update_inventory_display()
            self.update_category_checkboxes()
            self.update_location_checkboxes()
            self.save_inventory_to_excel()
            top.destroy()

        tk.Button(top, text="登録", command=submit).grid(row=6, column=0, padx=10, pady=15)
        tk.Button(top, text="キャンセル", command=top.destroy).grid(row=6, column=1, padx=10, pady=15)

    def open_inventory_input(self):
        """台帳入力ボタン押下時に、サブ機能（新規品登録、CSVインポート、QRコード生成）のウィンドウを表示"""
        win = tk.Toplevel(self.root)
        win.title("台帳入力")
        win.geometry("300x200")

        tk.Button(win, text="新規品番登録", width=20, command=self.register_new_product).pack(pady=10)
        tk.Button(win, text="CSVインポート", width=20, command=self.import_csv).pack(pady=10)
        tk.Button(win, text="QRコード生成", width=20, command=self.create_qr_code).pack(pady=10)
        tk.Button(win, text="閉じる", width=20, command=win.destroy).pack(pady=10)

    def create_buttons(self):
        """メイン画面下部に各機能ボタン（入庫、出庫、発注、台帳入力、設定、終了）を横並びに配置"""
        btn_specs = [
            ("入庫", self.stock_in),
            ("出庫", self.stock_out),
            ("発注", self.order_product),
            ("台帳入力", self.open_inventory_input),
            ("設定", self.open_settings),
            ("終了", self.root.destroy)
        ]
        for text, command in btn_specs:
            btn = tk.Button(self.button_frame, text=text, command=command, width=15)
            btn.pack(side="left", padx=5, pady=5)

    def order_product(self):
        """発注ボタン押下時の処理。対象商品を選択し、発注中フラグを立てる。"""
        selected = self.inventory_tree.selection()
        if selected:
            item_values = self.inventory_tree.item(selected[0], "values")
            selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(item_values[0])), None)
        else:
            use_manual = messagebox.askyesno("ID入力確認", 
                                "リストに選択がありません。\nIDを手動で入力しますか？\n「いいえ」を選択すると、再度リストから選択できます。")
            if use_manual:
                entered_id = ask_centered_string(self.root, "ID入力", "発注する商品のIDを入力してください:")
                if not entered_id:
                    return
                selected_item = next((item for item in self.inventory_data if str(item["id"]) == str(entered_id)), None)
                if not selected_item:
                    return messagebox.showerror("品番エラー", "入力されたIDに対応する品番が見つかりません。")
            else:
                return

        ok = messagebox.askyesno("発注確認", f"{selected_item['name']} を発注します。よろしいですか？")
        if not ok:
            return

        selected_item["order_pending"] = True
        self.update_inventory_display()
        messagebox.showinfo("発注完了", f"{selected_item['name']} は発注中です。")

    def open_settings(self):
        pwd = simpledialog.askstring("管理者認証", "管理者パスワードを入力してください", parent=self.root, show="*")
        if pwd != self.admin_password:
            messagebox.showerror("認証エラー", "パスワードが間違っています。")
            return

        settings_win = tk.Toplevel(self.root)
        settings_win.title("メール＆台帳設定")
        settings_win.geometry("400x300")

        tk.Label(settings_win, text="送信元メールアドレス:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        entry_sender = tk.Entry(settings_win, width=30)
        entry_sender.grid(row=0, column=1, padx=5, pady=5)
        entry_sender.insert(0, self.sender_email)

        tk.Label(settings_win, text="送信元パスワード:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry_sender_pw = tk.Entry(settings_win, width=30, show="*")
        entry_sender_pw.grid(row=1, column=1, padx=5, pady=5)
        entry_sender_pw.insert(0, self.sender_password)

        tk.Label(settings_win, text="通知先メールアドレス:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        entry_recipient = tk.Entry(settings_win, width=30)
        entry_recipient.grid(row=2, column=1, padx=5, pady=5)
        entry_recipient.insert(0, self.recipient_email)

        tk.Label(settings_win, text="台帳ファイル:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        entry_excel = tk.Entry(settings_win, width=30)
        entry_excel.grid(row=3, column=1, padx=5, pady=5)
        entry_excel.insert(0, self.EXCEL_FILE)

        def choose_excel_file():
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
            if file_path:
                entry_excel.delete(0, tk.END)
                entry_excel.insert(0, file_path)

        tk.Button(settings_win, text="参照", command=choose_excel_file).grid(row=3, column=2, padx=5, pady=5)

        def save_settings():
            self.sender_email = entry_sender.get().strip()
            self.sender_password = entry_sender_pw.get().strip()
            self.recipient_email = entry_recipient.get().strip()
            self.EXCEL_FILE = entry_excel.get().strip()
            messagebox.showinfo("設定完了", "メールおよび台帳設定を更新しました。")
            settings_win.destroy()

        tk.Button(settings_win, text="保存", command=save_settings).grid(row=4, column=0, padx=10, pady=15)
        tk.Button(settings_win, text="キャンセル", command=settings_win.destroy).grid(row=4, column=1, padx=10, pady=15)

    def check_low_stock(self):
        low_stock_items = []
        for item in self.inventory_data:
            qty = item.get("quantity", 0)
            if pd.isna(qty):
                qty = 0
            else:
                qty = int(qty)
            # order_pending が True の場合は既に発注中なので通知対象外とする
            if qty <= self.LOW_STOCK_THRESHOLD and not item.get("order_pending", False):
                low_stock_items.append(item)
        if low_stock_items:
            items_str = "\n".join([
                f"{item['name']} (在庫: {0 if pd.isna(item.get('quantity', 0)) else int(item.get('quantity', 0))})"
                for item in low_stock_items
            ])
            messagebox.showwarning("在庫注意", f"以下の商品で在庫数量が少なくなっています:\n{items_str}")
            send_low_stock_email_no_oauth(low_stock_items, self.sender_email, self.sender_password, self.recipient_email)

    def send_low_stock_email(self, low_stock_items):
        sender_email = os.getenv("GMAIL_USER", self.sender_email)
        sender_password = os.getenv("GMAIL_APP_PASSWORD", self.sender_password)
        recipient_email = self.recipient_email

        subject = "在庫不足通知 (SMTP - App Password)"
        body = "以下の商品で在庫が不足しています:\n" + "\n".join([
            f"{item['name']} (在庫: {0 if pd.isna(item.get('quantity', 0)) else int(item.get('quantity', 0))})"
            for item in low_stock_items
        ])

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587, timeout=5)
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
            server.quit()
            print("在庫不足通知メールを送信しました。")
        except Exception as e:
            print("メール送信に失敗しました:", e)

    def record_log(self, action, item, quantity):
        log_message = f"{action}: {item['name']} (ID: {item['id']}) - 数量: {quantity}"
        print(log_message)

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()
