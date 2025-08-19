import pandas as pd
import json
from datetime import datetime
import calendar

# --- 列名映射 ---
COLUMN_NAME_MAP = {
    'account_name': '账户名称',
    'month': '月份',
    'stock_code': '股票代码',
    'stock_name': '股票名称',
    'total_profit': '总收益',
    'monthly_total_profit': '月度总收益',
    'stock_total_profit': '股票总收益',
    'trade_pair_count': '交易对数',
    'sell_datetime': '卖出时间',
    'buy_datetime': '买入时间',
    'matched_quantity': '匹配数量',
    'buy_moneychg': '买入金额变化',
    'sell_moneychg': '卖出金额变化',
    'profit': '收益'
}

def get_current_month_range():
    """获取当月第一天和最后一天的日期字符串"""
    today = datetime.today()
    first_day = datetime(today.year, today.month, 1)
    last_day = datetime(today.year, today.month, calendar.monthrange(today.year, today.month)[1])
    return first_day.strftime('%Y%m%d'), last_day.strftime('%Y%m%d')

def parse_trade_data_from_content(content):
    """
    从文件内容字符串中解析交易记录，提取 ex_data.list 数组。
    """
    try:
        # 尝试将整个内容解析为JSON
        data = json.loads(content)
        
        # 检查是否有 ex_data 和 list
        # 同时检查接口返回的数据结构 (data.list)
        if 'ex_data' in data and 'list' in data['ex_data']:
            trades = data['ex_data']['list']
            return trades, None
        elif 'data' in data and 'list' in data['data']:
            # 接口返回的数据结构
            trades = data['data']['list']
            return trades, None
        else:
            return None, "错误：JSON中未找到 'ex_data.list' 或 'data.list'。"
            
    except json.JSONDecodeError as e:
        return None, f"JSON解析错误: {e}"
    except Exception as e:
        return None, f"解析内容时发生未知错误: {e}"

def preprocess_trades(trades):
    """
    预处理交易数据，转换格式，计算必要的字段。
    """
    if not trades:
        return pd.DataFrame(), "警告：未解析到任何交易记录。"

    df = pd.DataFrame(trades)
    
    # --- 数据预处理和字段标准化 ---
    # 1. 确保关键字段存在，用默认值填充缺失
    key_fields = ['account_name', 'stock_code', 'transDateTime', 'moneychg', 'trans_count', 'op']
    for field in key_fields:
        if field not in df.columns:
            df[field] = None # 或根据情况设置其他默认值
    
    # 2. 处理 transDateTime 格式 (文件是 'YYYYMMDDHHMMSS', 接口可能是 'YYYY-MM-DD HH:MM:SS')
    # 尝试第一种格式
    df['trans_datetime'] = pd.to_datetime(df['transDateTime'], format='%Y%m%d%H%M%S', errors='coerce')
    # 如果失败，尝试第二种格式
    if df['trans_datetime'].isna().any():
        df['trans_datetime'] = pd.to_datetime(df['transDateTime'], errors='coerce') # 更宽松的解析
    
    # 如果仍有转换失败的，记录并处理
    if df['trans_datetime'].isna().any():
        print("警告：部分 transDateTime 无法解析，这些记录将被忽略。")
        df = df.dropna(subset=['trans_datetime'])

    df['trans_date'] = df['trans_datetime'].dt.date
    df['month'] = df['trans_datetime'].dt.to_period('M')
    
    # 3. 数值字段转换
    df['moneychg'] = pd.to_numeric(df['moneychg'], errors='coerce').fillna(0)
    df['trans_count'] = pd.to_numeric(df['trans_count'], errors='coerce').fillna(0)
    df['op'] = pd.to_numeric(df['op'], errors='coerce').fillna(0).astype(int)
    
    # 4. 根据 op 字段确定交易类型 (1:买入, 2:卖出)
    df['trade_type'] = df['op'].map({1: '买入', 2: '卖出'})
    # 过滤出买入和卖出的记录
    df = df[df['trade_type'].isin(['买入', '卖出'])].copy()
    
    # 5. 计算数量的绝对值，方便后续处理
    df['quantity'] = df['trans_count'].abs()
    
    # 6. 确保字符串字段是字符串类型
    df['account_name'] = df['account_name'].astype(str)
    df['stock_code'] = df['stock_code'].astype(str)
    
    return df, None

def calculate_grid_profit_for_group(group_df):
    """
    为一个特定的 (账户, 股票, 月份) 组计算网格收益。
    匹配规则：为每个卖出记录找到时间在其之前且最近的买入记录。
    收益 = 卖出记录的 moneychg + 买入记录的 moneychg
    """
    # 分离买入和卖出记录
    buys = group_df[group_df['trade_type'] == '买入'].copy()
    sells = group_df[group_df['trade_type'] == '卖出'].copy()
    
    # 按时间排序
    buys = buys.sort_values('trans_datetime').reset_index(drop=True)
    sells = sells.sort_values('trans_datetime').reset_index(drop=True)
    
    # 为买入记录创建一个副本用于跟踪剩余可匹配数量
    buy_inventory = buys.copy()
    buy_inventory['remaining_quantity'] = buy_inventory['quantity']
    
    matched_trades = []
    total_profit = 0.0
    
    # 遍历每一个卖出记录
    for _, sell_record in sells.iterrows():
        sell_quantity = sell_record['quantity']
        sell_moneychg = sell_record['moneychg'] # 正数 (资金流入)
        sell_time = sell_record['trans_datetime']
        
        # 为当前卖出记录找到匹配的买入记录
        # 条件：买入时间 < 卖出时间，买入数量 > 0
        potential_buys = buy_inventory[
            (buy_inventory['trans_datetime'] < sell_time) & 
            (buy_inventory['remaining_quantity'] > 0)
        ].copy()
        
        # 按时间倒序排序，最近的在前面
        potential_buys = potential_buys.sort_values('trans_datetime', ascending=False)
        
        # 只要还有未匹配的卖出数量，并且还有潜在的买入记录可供匹配
        while sell_quantity > 0 and not potential_buys.empty:
            # 取出最近的买入记录
            buy_record = potential_buys.iloc[0]
            buy_index_in_inventory = buy_record.name # 这是 buy_inventory 的索引
            buy_quantity_available = buy_record['remaining_quantity']
            
            # 确定本次交易匹配的数量
            matched_quantity = min(sell_quantity, buy_quantity_available)
            
            # 计算匹配部分的买入金额变化
            matched_buy_moneychg = (buy_record['moneychg'] / buy_record['quantity']) * matched_quantity
            
            # 计算匹配部分的卖出金额变化
            matched_sell_moneychg = (sell_moneychg / sell_record['quantity']) * matched_quantity
            
            # 计算此匹配对的收益
            profit = matched_sell_moneychg + matched_buy_moneychg # moneychg对于买入是负数
            
            # 注意：为了与显示逻辑一致，这里交换了 buy_datetime 和 sell_datetime 的含义
            matched_trades.append({
                'sell_datetime': buy_record['trans_datetime'], # 买入时间
                'buy_datetime': sell_record['trans_datetime'], # 卖出时间
                'stock_code': sell_record['stock_code'],
                'matched_quantity': matched_quantity,
                'buy_moneychg': matched_buy_moneychg, # 负数
                'sell_moneychg': matched_sell_moneychg, # 正数
                'profit': profit
            })
            
            total_profit += profit
            
            # 更新买入记录的剩余数量
            new_buy_quantity = buy_quantity_available - matched_quantity
            buy_inventory.at[buy_index_in_inventory, 'remaining_quantity'] = new_buy_quantity
            
            # 更新待匹配的卖出数量
            sell_quantity -= matched_quantity
            
            # 更新潜在买入列表（移除已完全匹配的记录）
            if new_buy_quantity <= 0:
                potential_buys = potential_buys.iloc[1:] # 移除第一个（已完全匹配的）
            else:
                # 更新第一个记录的剩余数量
                potential_buys.iloc[0, potential_buys.columns.get_loc('remaining_quantity')] = new_buy_quantity
                
    return total_profit, matched_trades

def analyze_trades_from_data(trades_data, log_messages, stock_name_map=None):
    """从已解析的交易数据列表进行分析。"""
    # 初始化可能返回的 DataFrame
    summary_df = pd.DataFrame()
    all_matched_details = []
    account_month_summary = pd.DataFrame()
    stock_summary = pd.DataFrame()
    stock_detail_summary = pd.DataFrame()
    details_df = pd.DataFrame()

    try:
        if not trades_data:
            log_messages.append("未解析到任何有效的交易记录。")
            # 注意：返回值数量需要与函数签名匹配
            return account_month_summary, stock_summary, stock_detail_summary, details_df, log_messages

        log_messages.append(f"解析到 {len(trades_data)} 条原始记录。")
        log_messages.append("正在预处理交易数据...")
        df, error_msg = preprocess_trades(trades_data)
        
        if error_msg:
            log_messages.append(error_msg)
            
        if df.empty:
            log_messages.append("预处理后无有效交易记录。")
            # 返回空的 DataFrame 和日志
            return account_month_summary, stock_summary, stock_detail_summary, details_df, log_messages

        log_messages.append(f"预处理后得到 {len(df)} 条有效交易记录。")

        # --- 新增：核心分组和收益计算逻辑 ---
        # 1. 按 account_name, stock_code, month 分组
        log_messages.append("正在进行交易匹配和收益计算...")
        grouped = df.groupby(['account_name', 'stock_code', 'month'], group_keys=False)

        summary_data = []
        # 2. 对每个组调用 calculate_grid_profit_for_group
        for name, group in grouped:
            account_name, stock_code, month = name
            # 调用收益计算函数
            total_profit, matched_trades = calculate_grid_profit_for_group(group)
            trade_pair_count = len(matched_trades)

            # 3. 收集汇总信息
            summary_data.append({
                'account_name': account_name,
                'stock_code': stock_code,
                'month': month,
                'total_profit': total_profit,
                'trade_pair_count': trade_pair_count
            })

            # 4. 收集所有匹配的明细
            for detail in matched_trades:
                # 为每个明细记录添加账户和股票信息
                detail['account_name'] = account_name
                detail['stock_code'] = stock_code
                # profit 已在 calculate_grid_profit_for_group 中计算
                all_matched_details.append(detail)

        # 5. 创建最终的汇总 DataFrame summary_df
        summary_df = pd.DataFrame(summary_data)
        if not summary_df.empty:
            summary_df['total_profit'] = summary_df['total_profit'].round(2)
        log_messages.append("交易匹配和收益计算完成。")
        # --- 新增逻辑结束 ---

        # --- 原有后续处理逻辑 ---
        # 1. 账户月度汇总
        if not summary_df.empty:
            account_month_summary = summary_df.groupby(['account_name', 'month'])['total_profit'].sum().reset_index()
            account_month_summary.rename(columns={'total_profit': 'monthly_total_profit'}, inplace=True)
            account_month_summary['monthly_total_profit'] = account_month_summary['monthly_total_profit'].round(2)
        
        # --- 新增：添加股票名称 ---
        if stock_name_map:
            log_messages.append("正在添加股票名称...")
            def get_stock_name(code):
                return stock_name_map.get(str(code), str(code)) # 如果找不到名称，则使用代码
            
            # 为所有包含 'stock_code' 列的 DataFrame 添加 'stock_name' 列
            for df_temp in [summary_df, account_month_summary, stock_summary, stock_detail_summary, details_df]:
                if df_temp is not None and 'stock_code' in df_temp.columns:
                    df_temp['stock_name'] = df_temp['stock_code'].apply(get_stock_name)
            
            # 特别处理 all_matched_details（它是列表，转换为DataFrame后添加）
            if all_matched_details:
                details_df_for_name = pd.DataFrame(all_matched_details)
                if 'stock_code' in details_df_for_name.columns:
                    details_df_for_name['stock_name'] = details_df_for_name['stock_code'].apply(get_stock_name)
                    # 更新 all_matched_details 列表中的字典（如果需要，或者直接用 details_df_for_name）
                    # 这里选择更新 details_df_for_name 并在最后赋值给 details_df
                    details_df = details_df_for_name
                else:
                    details_df = pd.DataFrame(all_matched_details)
            else:
                details_df = pd.DataFrame(all_matched_details)

            log_messages.append("股票名称添加完成。")
        else:
            # 如果没有 stock_name_map，也要从 all_matched_details 创建 details_df
            details_df = pd.DataFrame(all_matched_details)

        # 2. 新增：股票汇总 (按账户、月份、股票)
        if not summary_df.empty:
            # 先添加 stock_name 列（如果存在）
            columns_to_include = ['account_name', 'month', 'stock_code', 'total_profit']
            # 如果 stock_name 存在，则也包含它
            if 'stock_name' in summary_df.columns:
                columns_to_include.insert(3, 'stock_name')
    
            stock_summary = summary_df[columns_to_include].copy()
            # 注意：如果 stock_name_map 存在，stock_name 列已在 summary_df 中添加
            stock_summary.rename(columns={'total_profit': 'stock_total_profit'}, inplace=True)
            stock_summary['stock_total_profit'] = stock_summary['stock_total_profit'].round(2)
            
            # 过滤掉总收益为0的股票
            stock_summary = stock_summary[stock_summary['stock_total_profit'] != 0]
        
        # 3. 股票明细 (保持原样，包含交易对数)
        if not summary_df.empty:
            stock_detail_summary = summary_df.copy()
            # 注意：如果 stock_name_map 存在，stock_name 列已在 summary_df 中添加
            stock_detail_summary['total_profit'] = stock_detail_summary['total_profit'].round(2)
            
            # 过滤掉总收益为0的股票
            stock_detail_summary = stock_detail_summary[stock_detail_summary['total_profit'] != 0]
        
        # details_df 的最终处理已在上面 stock_name_map 分支中处理
        # 确保 datetime 列是 datetime 类型 (如果 details_df 不为空)
        if not details_df.empty:
            # 确保 datetime 列是 datetime 类型
            details_df['sell_datetime'] = pd.to_datetime(details_df['sell_datetime'])
            details_df['buy_datetime'] = pd.to_datetime(details_df['buy_datetime'])
            details_df['account_name'] = details_df['account_name'].astype(str)
            # 重新计算 month，因为 details_df 现在来自 matched details
            details_df['month'] = details_df['sell_datetime'].dt.to_period('M').astype(str)
            details_df['profit'] = details_df['profit'].round(2)
        
        return account_month_summary, stock_summary, stock_detail_summary, details_df, log_messages

    except Exception as e:
        error_msg = f"分析过程中发生未知错误: {e}"
        log_messages.append(error_msg)
        # 即使出错也返回空的DataFrame和日志
        return account_month_summary, stock_summary, stock_detail_summary, details_df, log_messages

def analyze_trades_from_file(file_path, log_messages):
    """
    从文件路径读取并分析交易数据。
    """
    try:
        log_messages.append("正在解析交易数据...")
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        raw_trades, error_msg = parse_trade_data_from_content(content)
        if error_msg:
            log_messages.append(error_msg)
            # 注意：返回值数量需要与函数签名匹配
            return None, None, None, None, log_messages

        # 文件分析通常不获取股票名称，传递空字典
        return analyze_trades_from_data(raw_trades, log_messages, stock_name_map={})

    except FileNotFoundError:
        error_msg = f"错误：找不到文件 {file_path}"
        log_messages.append(error_msg)
        return None, None, None, None, log_messages
    except Exception as e:
        error_msg = f"从文件读取数据时发生错误: {e}"
        log_messages.append(error_msg)
        return None, None, None, None, log_messages