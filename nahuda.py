"""
=========================================================
名札PDF生成ツール (nahuda.py) - メンテナンスガイド
=========================================================

【主なメンテナンス方法・設定変更の手順】

1. レイアウトや座標（印字位置）の調整
   - create_pdf() 関数内の変数で各要素の座標を計算しています（A4用紙の左下が原点です）。
   - 全体の間隔調整: `offset_x` (横の間隔), `offset_y` (縦の間隔) の数値を変更してください。
   - 2つ目の役職の高さ基準: `base_y_2_val = 270 * mm` の数値を上下させることで調整可能です。
   - 名前や学部学年の個別調整: `base_x`, `base_y`, `year_y`, `pos_y` 等の計算式内の数値を変更します。

2. フォントの変更・サイズ調整
   - フォントファイル: "C:/Windows/Fonts/HGRKK.TTC" (HG創英角ﾎﾟｯﾌﾟ体) を読み込んでいます。
     別のフォントにする場合は、ttfonts.TTFont() のパスとフォント名を修正してください。
   - サイズ変更: `page.setFont("HGRKK", 26)` などの第2引数の数値(26など)を変更してください。

3. EXE化（実行ファイルのビルド）手順
   - Python環境で PyInstaller を用いてビルドします。
   - コマンドプロンプト等で以下のコマンドを実行してください:
     > pyinstaller --onefile --windowed nahuda.py
   - 生成された `dist/nahuda.exe` を配布用に使用します。

4. 新しい項目（入力フィールド）の追加
   - CSVヘッダーや、ManualInputWindow（手動入力画面）に入力欄を追加する場合は、以下の3点を同期して修正します。
     (1) GUIフォームへのラベル・Entryの追加
     (2) Treeview（リスト表示）のカラム設定の追加
     (3) add_entry(), generate_pdf() メソッド内の辞書データのキーの追加
=========================================================
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.units import mm
import pandas as pd
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def create_pdf(data, img_path, save_path):
    try:
        page = canvas.Canvas(save_path, pagesize=portrait(A4))
        # フォント登録はここで行うか、グローバルで行うか。毎回やってもエラーにはならないはずだが非効率かも
        # 念のためtry-exceptで囲むか、登録済みかチェックする
        try:
           pdfmetrics.registerFont(TTFont("HGRKK", "C:/Windows/Fonts/HGRKK.TTC"))
        except:
           pass # 既に登録されている場合など

        offset_x = 85* mm
        offset_y = -55.5* mm

        # 2個目の役職の位置調整用変数 (base_y_2)
        # ここを変更することで、2個目の役職のY座標の基準位置を調整できます
        base_y_2_val = 270 * mm 

        for idx, row in data.iterrows():
            mod = (idx % 10) + 1
            name = str(row["名前"])
            year = str(row["学部学年"])
            pos = str(row["役職"]) if "役職" in row and not pd.isnull(row["役職"]) else ''
            pos2 = str(row["役職2"]) if "役職2" in row and not pd.isnull(row["役職2"]) else ''
            
            furigana = ""
            if "フリガナ" in row and not pd.isnull(row["フリガナ"]):
                furigana = str(row["フリガナ"])

            base_x = 57.5*mm if mod % 2 == 1 else 57.5*mm+offset_x
            # base_x = 57*mm if mod % 2 == 1 else 57*mm+offset_x

            base_y = 255*mm + ((mod-1)//2)*offset_y
            right_x = 89*mm if mod % 2 == 1 else 90*mm+offset_x
            string_x = 25*mm if mod % 2 == 1 else 25*mm+offset_x
            year_y = 280*mm + ((mod-1)//2)*offset_y
            pos_y = 274*mm + ((mod-1)//2)*offset_y
            
            # 2個目の役職のY座標計算
            pos_y_2 = base_y_2_val + ((mod-1)//2)*offset_y

            furigana_y = base_y + 8*mm  # Slightly above the name

            if mod == 1:
                page.drawImage(img_path, 0*mm, 0*mm, 210*mm, 297*mm)
            
            if furigana:
                page.setFont("HGRKK", 10)  # Smaller font for Furigana
                page.drawCentredString(base_x, furigana_y, furigana)

            page.setFont("HGRKK", 26)
            page.drawCentredString(base_x, base_y, name)
            page.setFont("HGRKK", 13)
            page.drawRightString(right_x, year_y, year)
            page.setFont("HGRKK", 12)
            page.drawString(string_x, pos_y, pos)
            
            # 2個目の役職を描画
            if pos2:
                page.drawString(string_x, pos_y_2, pos2)

            if mod == 10 or idx == len(data) - 1:
                page.showPage()

        page.save()
        print("PDF生成が完了しました:", save_path)
        return True
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return False

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("名札PDF生成ツール")
        self.root.geometry("400x300")

        # モード選択ボタン
        tk.Label(root, text="モードを選択してください", font=("Helvetica", 16)).pack(pady=20)
        
        tk.Button(root, text="CSVファイルから読み込み", command=self.csv_mode, width=30, height=2).pack(pady=10)
        tk.Button(root, text="手動入力", command=self.manual_mode, width=30, height=2).pack(pady=10)

    def select_background_and_savepath(self):
        # 背景画像選択
        img_path = filedialog.askopenfilename(
            title='背景画像を選択', filetypes=[('PNG Files', '*.png'), ('All files', '*.*')]
        )
        if not img_path:
            return None, None
            
        # 保存先選択
        save_path = filedialog.asksaveasfilename(
            title='保存先を選択', defaultextension='.pdf', filetypes=[('PDF Files', '*.pdf')]
        )
        if not save_path:
            return None, None
            
        return img_path, save_path

    def csv_mode(self):
        # CSVファイル選択
        csv_path = filedialog.askopenfilename(
            title='CSVファイルを選択', filetypes=[('CSV Files', '*.csv')]
        )
        if not csv_path:
            return

        try:
            data = pd.read_csv(csv_path)
        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました:\n{e}")
            return

        img_path, save_path = self.select_background_and_savepath()
        if not img_path or not save_path:
            return

        if create_pdf(data, img_path, save_path):
             messagebox.showinfo("完了", "PDF生成が完了しました")

    def manual_mode(self):
        ManualInputWindow(self.root)

class ManualInputWindow:
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("手動入力モード")
        self.window.geometry("700x550") # 少し広げる (600x500 -> 700x550)
        
        self.data_list = []

        # 入力フォーム
        frame_input = tk.Frame(self.window)
        frame_input.pack(pady=10)

        tk.Label(frame_input, text="名前:").grid(row=0, column=0, padx=5)
        self.entry_name = tk.Entry(frame_input)
        self.entry_name.grid(row=0, column=1, padx=5)

        tk.Label(frame_input, text="フリガナ:").grid(row=1, column=0, padx=5)
        self.entry_furigana = tk.Entry(frame_input)
        self.entry_furigana.grid(row=1, column=1, padx=5)

        tk.Label(frame_input, text="学部学年:").grid(row=2, column=0, padx=5)
        self.entry_year = tk.Entry(frame_input)
        self.entry_year.grid(row=2, column=1, padx=5)

        tk.Label(frame_input, text="役職:").grid(row=3, column=0, padx=5)
        self.entry_pos = tk.Entry(frame_input)
        self.entry_pos.grid(row=3, column=1, padx=5)

        tk.Label(frame_input, text="役職2:").grid(row=4, column=0, padx=5)
        self.entry_pos2 = tk.Entry(frame_input)
        self.entry_pos2.grid(row=4, column=1, padx=5)

        tk.Button(frame_input, text="追加", command=self.add_entry).grid(row=5, column=0, columnspan=2, pady=10)

        # リスト表示
        self.tree = ttk.Treeview(self.window, columns=("Name", "Furigana", "Year", "Pos", "Pos2"), show="headings")
        self.tree.heading("Name", text="名前")
        self.tree.heading("Furigana", text="フリガナ")
        self.tree.heading("Year", text="学部学年")
        self.tree.heading("Pos", text="役職")
        self.tree.heading("Pos2", text="役職2")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # アクションボタン
        frame_action = tk.Frame(self.window)
        frame_action.pack(pady=10)
        
        tk.Button(frame_action, text="削除", command=self.delete_entry).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_action, text="PDF生成", command=self.generate_pdf).pack(side=tk.LEFT, padx=5)

    def add_entry(self):
        name = self.entry_name.get()
        furigana = self.entry_furigana.get()
        year = self.entry_year.get()
        pos = self.entry_pos.get()
        pos2 = self.entry_pos2.get()

        if not name:
             messagebox.showwarning("警告", "名前は必須です")
             return

        self.data_list.append({"名前": name, "フリガナ": furigana, "学部学年": year, "役職": pos, "役職2": pos2})
        self.tree.insert("", "end", values=(name, furigana, year, pos, pos2))
        
        # 入力欄クリア
        self.entry_name.delete(0, tk.END)
        self.entry_furigana.delete(0, tk.END)
        self.entry_year.delete(0, tk.END)
        self.entry_pos.delete(0, tk.END)
        self.entry_pos2.delete(0, tk.END)
        self.entry_name.focus_set()

    def delete_entry(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        for item in selected_item:
            # Treeviewから削除
            idx = self.tree.index(item)
            self.tree.delete(item)
            # データリストから削除 (インデックスがずれるので注意が必要だが、選択削除は1つずつか、後ろからやるのが安全)
            # 今回は簡易的に再構築するか、ID管理する方が良いが、リスト同期でいく
            del self.data_list[idx]

    def generate_pdf(self):
        if not self.data_list:
            messagebox.showwarning("警告", "データがありません")
            return

        # ここで背景画像と保存先を選択させる（Appクラスのメソッドを借用するか再実装）
        # 簡単のため再実装
        img_path = filedialog.askopenfilename(
            title='背景画像を選択', filetypes=[('PNG Files', '*.png'), ('All files', '*.*')]
        )
        if not img_path:
            return
            
        save_path = filedialog.asksaveasfilename(
            title='保存先を選択', defaultextension='.pdf', filetypes=[('PDF Files', '*.pdf')]
        )
        if not save_path:
            return

        # データフレーム作成
        df = pd.DataFrame(self.data_list)
        
        if create_pdf(df, img_path, save_path):
             messagebox.showinfo("完了", "PDF生成が完了しました")


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
 
 