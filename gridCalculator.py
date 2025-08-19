import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
from datetime import datetime

from data_processor import (
    get_current_month_range, 
    analyze_trades_from_data,
    COLUMN_NAME_MAP
)
from api_client import APIClient
from excel_exporter import save_results_to_excel
from table_manager import TableManager

class GridProfitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网格交易收益分析工具")
        self.root.geometry("1250x850")

        self.account_month_df = None
        self.stock_summary_df = None
        self.stock_detail_df = None
        self.details_df = None
        self.api_controls = {}
        self.stock_summary_controls = {}
        
        # 初始化表格管理器
        self.table_manager = TableManager(self)

        self.create_widgets()

    def create_widgets(self):
        # --- API 接口区域 ---
        api_frame = tk.LabelFrame(self.root, text="API接口获取数据")
        api_frame.pack(pady=10, padx=10, fill=tk.X)

        # API 输入控件
        api_inputs_frame = tk.Frame(api_frame)
        api_inputs_frame.pack(fill=tk.X, padx=5, pady=5)

        # 第一行
        row1 = tk.Frame(api_inputs_frame)
        row1.pack(fill=tk.X, pady=2)
        tk.Label(row1, text="用户账号:", width=10, anchor='w').pack(side=tk.LEFT)
        self.api_controls['user_id_var'] = tk.StringVar(value='')
        tk.Entry(row1, textvariable=self.api_controls['user_id_var'], width=20).pack(side=tk.LEFT, padx=(5, 10))
        
        tk.Label(row1, text="股票账户:", width=10, anchor='w').pack(side=tk.LEFT)
        self.api_controls['fund_key_var'] = tk.StringVar(value='')
        tk.Entry(row1, textvariable=self.api_controls['fund_key_var'], width=20).pack(side=tk.LEFT, padx=(5, 10))

        tk.Label(row1, text="开始日期:", width=10, anchor='w').pack(side=tk.LEFT)
        # 修改开始日期默认值为当月第一天
        first_day, _ = get_current_month_range()
        self.api_controls['start_date_var'] = tk.StringVar(value=first_day)
        tk.Entry(row1, textvariable=self.api_controls['start_date_var'], width=10).pack(side=tk.LEFT, padx=(5, 10))

        tk.Label(row1, text="结束日期:", width=10, anchor='w').pack(side=tk.LEFT)
        # 修改结束日期默认值为当月最后一天
        _, last_day = get_current_month_range()
        self.api_controls['end_date_var'] = tk.StringVar(value=last_day)
        tk.Entry(row1, textvariable=self.api_controls['end_date_var'], width=10).pack(side=tk.LEFT, padx=(5, 0))

        # 第二行
        row2 = tk.Frame(api_inputs_frame)
        row2.pack(fill=tk.X, pady=2)
        tk.Label(row2, text="Cookie:", width=10, anchor='w').pack(side=tk.LEFT)
        self.api_controls['cookie_var'] = tk.StringVar(value='')
        tk.Entry(row2, textvariable=self.api_controls['cookie_var'], width=50).pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)

        # API 操作按钮
        api_button_frame = tk.Frame(api_frame)
        api_button_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Button(api_button_frame, text="从接口获取数据", command=self.start_api_analysis).pack(side=tk.LEFT)

        # --- 通用操作按钮区域 ---
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        self.clear_button = tk.Button(button_frame, text="清空结果", command=self.clear_results)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # --- Notebook (标签页) ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # 标签页 1: 账户月度汇总
        self.account_month_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.account_month_frame, text="账户月度汇总")

        # 标签页 2: 股票汇总 (新增筛选和排序功能)
        self.stock_summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stock_summary_frame, text="股票汇总")
        self.create_stock_summary_with_controls()

        # 标签页 3: 股票明细
        self.stock_detail_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stock_detail_frame, text="股票明细")

        # 标签页 4: 交易匹配明细 (文本)
        self.details_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.details_frame, text="交易匹配明细")
        self.details_text = scrolledtext.ScrolledText(self.details_frame, wrap=tk.NONE, state=tk.DISABLED, font=("Consolas", 9))
        v_scrollbar_d = tk.Scrollbar(self.details_frame, orient=tk.VERTICAL, command=self.details_text.yview)
        h_scrollbar_d = tk.Scrollbar(self.details_frame, orient=tk.HORIZONTAL, command=self.details_text.xview)
        self.details_text.configure(yscrollcommand=v_scrollbar_d.set, xscrollcommand=h_scrollbar_d.set)
        self.details_text.grid(row=0, column=0, sticky='nsew')
        v_scrollbar_d.grid(row=0, column=1, sticky='ns')
        h_scrollbar_d.grid(row=1, column=0, sticky='ew')
        self.details_frame.grid_rowconfigure(0, weight=1)
        self.details_frame.grid_columnconfigure(0, weight=1)

        # 标签页 5: 运行日志
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="运行日志")
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def create_stock_summary_with_controls(self):
        """为股票汇总标签页创建筛选和排序控件"""
        control_frame = ttk.Frame(self.stock_summary_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(control_frame, text="账户:").pack(side=tk.LEFT, padx=(0, 5))
        self.stock_summary_controls['account_var'] = tk.StringVar(value="全部")
        self.stock_summary_controls['account_combo'] = ttk.Combobox(control_frame, textvariable=self.stock_summary_controls['account_var'], width=15, state="readonly")
        self.stock_summary_controls['account_combo'].pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(control_frame, text="月份:").pack(side=tk.LEFT, padx=(0, 5))
        self.stock_summary_controls['month_var'] = tk.StringVar(value="全部")
        self.stock_summary_controls['month_combo'] = ttk.Combobox(control_frame, textvariable=self.stock_summary_controls['month_var'], width=10, state="readonly")
        self.stock_summary_controls['month_combo'].pack(side=tk.LEFT, padx=(0, 10))
        
        self.stock_summary_controls['profit_sort_var'] = tk.BooleanVar(value=True)
        self.stock_summary_controls['profit_sort_check'] = ttk.Checkbutton(
            control_frame, 
            text="按收益降序", 
            variable=self.stock_summary_controls['profit_sort_var']
        )
        self.stock_summary_controls['profit_sort_check'].pack(side=tk.LEFT, padx=(0, 10))
        
        self.stock_summary_controls['apply_button'] = ttk.Button(
            control_frame, 
            text="应用筛选", 
            command=self.apply_stock_summary_filter
        )
        self.stock_summary_controls['apply_button'].pack(side=tk.LEFT)
        
        self.stock_summary_controls['table_frame'] = ttk.Frame(self.stock_summary_frame)
        self.stock_summary_controls['table_frame'].pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def update_stock_summary_controls(self):
        if self.stock_summary_df is None or self.stock_summary_df.empty:
            return
            
        accounts = ["全部"] + sorted(self.stock_summary_df['account_name'].unique())
        self.stock_summary_controls['account_combo']['values'] = accounts
        if self.stock_summary_controls['account_var'].get() not in accounts:
            self.stock_summary_controls['account_var'].set("全部")
            
        months = ["全部"] + sorted(self.stock_summary_df['month'].astype(str).unique())
        self.stock_summary_controls['month_combo']['values'] = months
        if self.stock_summary_controls['month_var'].get() not in months:
            self.stock_summary_controls['month_var'].set("全部")

    def apply_stock_summary_filter(self):
        if self.stock_summary_df is None or self.stock_summary_df.empty:
            return
            
        df_filtered = self.stock_summary_df.copy()
        
        selected_account = self.stock_summary_controls['account_var'].get()
        if selected_account != "全部":
            df_filtered = df_filtered[df_filtered['account_name'] == selected_account]
            
        selected_month = self.stock_summary_controls['month_var'].get()
        if selected_month != "全部":
            df_filtered = df_filtered[df_filtered['month'].astype(str) == selected_month]
            
        sort_ascending = not self.stock_summary_controls['profit_sort_var'].get()
        try:
            df_filtered = df_filtered.sort_values(
                by=['account_name', 'month', 'stock_total_profit'],
                ascending=[True, True, sort_ascending]
            )
        except Exception as e:
            self.log_message(f"股票汇总表筛选排序时出错: {e}")
        
        self.table_manager.populate_table("stock_summary", df_filtered)

    def treeview_sort_column(self, table_type, col, reverse):
        try:
            tree = self.table_manager.treeviews.get(table_type)
            if not tree:
                return
            data = [(tree.set(k, col), k) for k in tree.get_children('')]
            try:
                data.sort(key=lambda t: float(t[0]), reverse=reverse)
            except ValueError:
                data.sort(reverse=reverse)

            for index, (val, k) in enumerate(data):
                tree.move(k, '', index)

            tree.heading(col, command=lambda _col=col: self.treeview_sort_column(table_type, _col, not reverse))
        except Exception as e:
            self.log_message(f"排序时出错: {e}")

    def start_api_analysis(self):
        user_id = self.api_controls['user_id_var'].get().strip()
        fund_key = self.api_controls['fund_key_var'].get().strip()
        cookie = self.api_controls['cookie_var'].get().strip()
        start_date = self.api_controls['start_date_var'].get().strip()
        end_date = self.api_controls['end_date_var'].get().strip()

        if not all([user_id, fund_key, cookie, start_date, end_date]):
            messagebox.showwarning("警告", "请填写所有API接口参数。")
            return

        self.log_message("开始从接口获取数据...")
        self.clear_results()
        self.clear_button.config(state=tk.DISABLED)

        import threading
        thread = threading.Thread(target=self.run_api_analysis, args=(user_id, fund_key, cookie, start_date, end_date))
        thread.daemon = True
        thread.start()

    def run_api_analysis(self, user_id, fund_key, cookie, start_date, end_date):
        try:
            log_messages = ["正在通过API获取交易数据..."]
            client = APIClient(user_id, fund_key, cookie, start_date, end_date)
            
            # 1. 获取交易历史
            trade_response = client.get_stock_history()
            
            if trade_response.status_code == 200:
                log_messages.append("交易数据API请求成功。")
                trade_data = trade_response.json()
                
                if trade_data.get('error_code') != '0':
                    error_msg = trade_data.get('error_msg', '未知API错误')
                    raise Exception(f"交易数据API返回错误: {error_msg}")

                raw_trades = None
                if 'ex_data' in trade_data and 'list' in trade_data['ex_data']:
                    raw_trades = trade_data['ex_data']['list']
                elif 'data' in trade_data and 'list' in trade_data['data']:
                    raw_trades = trade_data['data']['list']
                
                if not raw_trades:
                    raise Exception("API返回交易数据中未找到交易记录列表。")

                log_messages.append(f"从API获取到 {len(raw_trades)} 条交易记录。")
                
                # 2. 获取股票持仓信息（用于股票名称）
                log_messages.append("正在通过API获取股票持仓信息...")
                position_response = client._get_stock_position()
                stock_name_map = {}
                if position_response.status_code == 200:
                    log_messages.append("股票持仓信息API请求成功。")
                    position_data = position_response.json()
                    if position_data.get('error_code') == '0':
                        positions = position_data.get('ex_data', {}).get('position', [])
                        for pos in positions:
                            code = pos.get('code')
                            name = pos.get('name')
                            if code and name:
                                stock_name_map[code] = name
                        log_messages.append(f"获取到 {len(stock_name_map)} 支股票的名称。")
                    else:
                        error_msg = position_data.get('error_msg', '未知API错误')
                        log_messages.append(f"股票持仓信息API返回错误: {error_msg}")
                else:
                    log_messages.append(f"股票持仓信息API请求失败，状态码: {position_response.status_code}")
                
                # 3. 调用分析函数 (传递 stock_name_map)
                account_month_df, stock_summary_df, stock_detail_df, details_df, log_messages = analyze_trades_from_data(raw_trades, log_messages, stock_name_map)
                # 4. 传递 details_df 而不是 details_text
                self.root.after(0, self.display_results, account_month_df, stock_summary_df, stock_detail_df, details_df, log_messages, stock_name_map)
            else:
                raise Exception(f"交易数据API请求失败，状态码: {trade_response.status_code}")
                
        except Exception as e:
            self.root.after(0, self.log_message, f"API获取数据或分析出错: {e}")
        finally:
            self.root.after(0, lambda: self.clear_button.config(state=tk.NORMAL))

    def generate_details_text(self, details_df):
        """
        根据 details_df 生成用于显示的文本。"""
        result_text = ""
        result_text += "="*150 + "\n"
        result_text += "详细的交易匹配记录 (收益 = 卖出moneychg + 买入moneychg)\n"
        result_text += "="*150 + "\n"
        if details_df is not None and not details_df.empty:
            # 为了文本显示，交换买/卖时间的含义 (如果需要，取决于 details_df 的结构)
            # 假设 details_df 已经是最终格式
            details_for_text = details_df.copy()
            
            # 构建表头，考虑是否有 stock_name
            if 'stock_name' in details_for_text.columns:
                header = f"{'账户名称':<25} {'月份':<10} {'股票名称':<15} {'股票代码':<10} {'卖出时间':<20} {'买入时间':<20} {'匹配数量':<8} {'买入金额变化':<15} {'卖出金额变化':<15} {'收益':<12}\n"
                separator_len = 170
            else:
                header = f"{'账户名称':<25} {'月份':<10} {'股票代码':<10} {'卖出时间':<20} {'买入时间':<20} {'匹配数量':<8} {'买入金额变化':<15} {'卖出金额变化':<15} {'收益':<12}\n"
                separator_len = 150
                
            result_text += header
            result_text += "-" * separator_len + "\n"
            
            # 确保排序列存在
            sort_cols = ['account_name', 'month', 'stock_code']
            if 'stock_name' in details_for_text.columns:
                sort_cols.insert(2, 'stock_name')
                
            for _, row in details_for_text.sort_values(sort_cols).iterrows():
                # 将Period类型的month转换为字符串
                month_str = str(row['month']) if pd.notna(row['month']) else ""
                
                if 'stock_name' in row:
                     result_text += f"{row['account_name']:<25} {month_str:<10} {row['stock_name']:<15} {row['stock_code']:<10} " \
                                   f"{row['sell_datetime'].strftime('%Y-%m-%d %H:%M:%S'):<20} {row['buy_datetime'].strftime('%Y-%m-%d %H:%M:%S'):<20} " \
                                   f"{row['matched_quantity']:<8} {row['buy_moneychg']:<15.2f} {row['sell_moneychg']:<15.2f} {row['profit']:<12.2f}\n"
                else:
                     result_text += f"{row['account_name']:<25} {month_str:<10} {row['stock_code']:<10} " \
                                   f"{row['sell_datetime'].strftime('%Y-%m-%d %H:%M:%S'):<20} {row['buy_datetime'].strftime('%Y-%m-%d %H:%M:%S'):<20} " \
                                   f"{row['matched_quantity']:<8} {row['buy_moneychg']:<15.2f} {row['sell_moneychg']:<15.2f} {row['profit']:<12.2f}\n"
        else:
            result_text += "无匹配记录。\n"
        return result_text

    def display_results(self, account_month_df, stock_summary_df, stock_detail_df, details_df, log_messages, stock_name_map):
        """
        展示分析结果。
        新增 stock_name_map 参数，虽然主要在 analyze_trades_from_data 中用了，
        但保留参数以备将来可能需要。
        现在接收 details_df 而不是 details_text。
        """
        self.account_month_df = account_month_df
        self.stock_summary_df = stock_summary_df
        self.stock_detail_df = stock_detail_df
        # 存储 details_df
        self.details_df = details_df 
        
        if self.account_month_df is not None and not self.account_month_df.empty:
            self.table_manager.populate_table("account_month", self.account_month_df)
        else:
            self.log_message("账户月度汇总数据为空。")

        if self.stock_summary_df is not None and not self.stock_summary_df.empty:
            self.update_stock_summary_controls()
            self.table_manager.populate_table("stock_summary", self.stock_summary_df)
        else:
            self.log_message("股票汇总数据为空。")

        if self.stock_detail_df is not None and not self.stock_detail_df.empty:
            self.table_manager.populate_table("stock_detail", self.stock_detail_df)
        else:
            self.log_message("股票明细数据为空。")

        # 重新生成 details_text 用于显示
        details_text = self.generate_details_text(self.details_df)
        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        self.details_text.insert(tk.END, details_text)
        self.details_text.config(state=tk.DISABLED)

        for msg in log_messages:
            self.log_message(msg)
            
        self.log_message("分析完成。")

        # 保存结果到Excel (现在传递 details_df)
        if any(df is not None and not df.empty for df in [account_month_df, stock_summary_df, stock_detail_df, details_df]):
            try:
                output_file = '网格交易收益分析结果.xlsx'
                # 调用修改后的 save function，传递 details_df
                success, save_msg = save_results_to_excel(account_month_df, stock_summary_df, stock_detail_df, details_df, output_file)
                self.log_message(save_msg)
            except Exception as e:
                self.log_message(f"保存Excel时出错3: {e}")

    def clear_results(self):
        self.table_manager.clear_tables()

        self.details_text.config(state=tk.NORMAL)
        self.details_text.delete(1.0, tk.END)
        self.details_text.config(state=tk.DISABLED)
        
        # 同时清空运行日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        self.account_month_df = None
        self.stock_summary_df = None
        self.stock_detail_df = None
        self.details_df = None # 清空 details_df
        
        if 'account_var' in self.stock_summary_controls:
            self.stock_summary_controls['account_var'].set("全部")
        if 'month_var' in self.stock_summary_controls:
            self.stock_summary_controls['month_var'].set("全部")
        if 'profit_sort_var' in self.stock_summary_controls:
            self.stock_summary_controls['profit_sort_var'].set(True)
            
        # 不再在日志中显示"结果已清空"消息，因为日志本身已被清空

    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

# --- 主程序入口 ---
if __name__ == "__main__":
    root = tk.Tk()
    app = GridProfitApp(root)
    root.mainloop()