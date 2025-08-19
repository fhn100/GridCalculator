import tkinter as tk
from tkinter import ttk
import pandas as pd
from data_processor import COLUMN_NAME_MAP

class TableManager:
    """表格管理器，用于处理表格的创建和更新"""
    
    def __init__(self, app):
        self.app = app
        self.treeviews = {}
        
    def create_dynamic_table(self, table_type, df):
        if table_type == "stock_summary":
            parent_frame = self.app.stock_summary_controls['table_frame']
        else:
            parent_frame = getattr(self.app, f"{table_type}_frame")
        
        for item in parent_frame.winfo_children():
            item.destroy()

        frame = ttk.Frame(parent_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = list(df.columns)
        display_columns = [COLUMN_NAME_MAP.get(col, col) for col in columns]
        
        tree = ttk.Treeview(frame, columns=columns, show='headings')
        self.treeviews[table_type] = tree
        
        for i, (col, disp_col) in enumerate(zip(columns, display_columns)):
            tree.heading(col, text=disp_col, command=lambda _col=col: self.app.treeview_sort_column(table_type, _col, False))
            if col in ['account_name']:
                tree.column(col, width=180, anchor='w')
            elif col in ['stock_code']:
                tree.column(col, width=100, anchor='w')
            elif col in ['stock_name']: # 新增股票名称列宽
                tree.column(col, width=120, anchor='w')
            elif 'profit' in col.lower() or 'amount' in col.lower() or 'moneychg' in col.lower():
                tree.column(col, width=120, anchor='e')
            elif 'count' in col.lower() or 'quantity' in col.lower():
                tree.column(col, width=80, anchor='center')
            elif 'month' in col.lower():
                tree.column(col, width=100, anchor='center')
            elif 'datetime' in col.lower():
                tree.column(col, width=150, anchor='w')
            else:
                tree.column(col, width=130, anchor='w')

        v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
    def populate_table(self, table_type, df):
        if df is None or df.empty:
            return

        # 特殊处理股票明细表的列顺序
        if table_type == "stock_detail" and df is not None and not df.empty:
            # 定义期望的列顺序
            preferred_order = ['account_name', 'month', 'stock_code', 'stock_name', 'total_profit', 'trade_pair_count']
            # 获取实际存在的列
            available_columns = [col for col in preferred_order if col in df.columns]
            # 添加其他可能存在的列
            other_columns = [col for col in df.columns if col not in available_columns]
            # 合并列顺序
            new_column_order = available_columns + other_columns
            # 重新排列DataFrame列
            df = df[new_column_order]
        
        # 特殊处理股票汇总表的列顺序
        if table_type == "stock_summary" and df is not None and not df.empty:
            # 定义期望的列顺序，确保股票名称在股票代码之后
            preferred_order = ['account_name', 'month', 'stock_code', 'stock_name', 'stock_total_profit']
            # 获取实际存在的列
            available_columns = [col for col in preferred_order if col in df.columns]
            # 添加其他可能存在的列
            other_columns = [col for col in df.columns if col not in available_columns]
            # 合并列顺序
            new_column_order = available_columns + other_columns
            # 重新排列DataFrame列
            df = df[new_column_order]

        if table_type not in self.treeviews or not self.treeviews[table_type].winfo_exists():
            self.create_dynamic_table(table_type, df)
        
        tree = self.treeviews.get(table_type)
        if not tree:
            self.app.log_message(f"无法找到或创建 {table_type} 的表格。")
            return

        for item in tree.get_children():
            tree.delete(item)

        for index, row in df.iterrows():
            values = [row[col] for col in df.columns]
            tree.insert("", "end", values=values)
            
    def clear_tables(self):
        for table_type in ["account_month", "stock_summary", "stock_detail"]:
            tree = self.treeviews.get(table_type)
            if tree:
                for item in tree.get_children():
                    tree.delete(item)