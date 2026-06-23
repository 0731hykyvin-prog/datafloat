#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading

# ========== 数据处理函数 ==========
def process_data(input_folder, person_db_path, output_folder, similarity_threshold, time_window, log_callback):
    try:
        log_callback("="*60)
        log_callback("开始处理视频布控数据...")
        log_callback("="*60)
        
        column_names = ['目标类型', '预警任务类型', '预警任务名称', '预警任务目标', 
                        '告警设备', '告警时间', '预警任务原因', '证件类型', '证件号码', '相似度']
        
        root_path = Path(input_folder)
        csv_files = list(root_path.rglob('*.csv'))
        
        if not csv_files:
            log_callback("❌ 未找到CSV文件")
            return None
        
        log_callback(f"找到 {len(csv_files)} 个CSV文件")
        
        def clean_excel_string(value):
            if isinstance(value, str):
                if value.startswith('="') and value.endswith('"'):
                    return value[2:-1]
            return value
        
        all_data = []
        
        for i, csv_file in enumerate(csv_files):
            log_callback(f"  [{i+1}/{len(csv_files)}] 读取: {csv_file.name}")
            
            try:
                df = pd.read_csv(csv_file, encoding='utf-8', skiprows=5,
                                names=column_names, on_bad_lines='skip')
                df = df.dropna(how='all')
                
                if len(df) > 0:
                    for col in df.columns:
                        df[col] = df[col].apply(clean_excel_string)
                    all_data.append(df)
                    log_callback(f"      ✓ 成功，{len(df)} 行")
            except Exception as e:
                log_callback(f"      ✗ 失败: {e}")
                continue
        
        if not all_data:
            log_callback("❌ 没有成功读取任何文件")
            return None
        
        log_callback(f"\n合并数据中...")
        result = pd.concat(all_data, ignore_index=True)
        result['相似度'] = pd.to_numeric(result['相似度'], errors='coerce')
        result['告警时间'] = pd.to_datetime(result['告警时间'], errors='coerce')
        
        log_callback(f"筛选相似度 >= {similarity_threshold}...")
        result = result.dropna(subset=['相似度'])
        result_filtered = result[result['相似度'] >= similarity_threshold].copy()
        log_callback(f"  筛选后: {len(result_filtered)} 条记录")
        
        log_callback(f"去重处理（时间窗口 {time_window} 秒）...")
        result_filtered = result_filtered.dropna(subset=['告警时间'])
        
        if len(result_filtered) > 0:
            def dedupe_group(group):
                if len(group) <= 1:
                    return group
                group = group.sort_values('告警时间').copy()
                group['_temp_group'] = 0
                current_group = 0
                last_time = None
                for idx in group.index:
                    current_time = group.loc[idx, '告警时间']
                    if last_time is None or (current_time - last_time).total_seconds() > time_window:
                        current_group += 1
                    group.loc[idx, '_temp_group'] = current_group
                    last_time = current_time
                
                def select_row(subgroup):
                    if len(subgroup) == 1:
                        return subgroup.iloc[[0]]
                    non_empty_device = subgroup[subgroup['告警设备'].notna() & (subgroup['告警设备'] != '')]
                    if len(non_empty_device) > 0:
                        return non_empty_device.iloc[[0]]
                    else:
                        return subgroup.iloc[[0]]
                
                result_group = group.groupby('_temp_group').apply(select_row)
                if isinstance(result_group, pd.DataFrame):
                    return result_group
                else:
                    result_group = result_group.reset_index(drop=True)
                    if isinstance(result_group, pd.Series):
                        result_group = pd.DataFrame(result_group.tolist(), columns=group.columns)
                    return result_group
            
            result_deduped_list = []
            for id_num, group in result_filtered.groupby('证件号码'):
                deduped_group = dedupe_group(group)
                result_deduped_list.append(deduped_group)
            
            if result_deduped_list:
                result_deduped = pd.concat(result_deduped_list, ignore_index=True)
            else:
                result_deduped = pd.DataFrame(columns=result_filtered.columns)
            
            if '_temp_group' in result_deduped.columns:
                result_deduped = result_deduped.drop(columns=['_temp_group'])
        else:
            result_deduped = result_filtered
        
        log_callback(f"  去重后: {len(result_deduped)} 条记录")
        
        log_callback("添加深夜标签...")
        result_deduped['告警时间'] = pd.to_datetime(result_deduped['告警时间'], errors='coerce')
        result_deduped['小时'] = result_deduped['告警时间'].dt.hour
        result_deduped['深夜出行'] = result_deduped['小时'].apply(lambda x: 1 if (x >= 23 or x <= 5) else 0)
        night_count = result_deduped['深夜出行'].sum()
        log_callback(f"  深夜出行记录: {night_count} 条")
        
        log_callback("统计深夜出行次数...")
        night_stats = result_deduped[result_deduped['深夜出行'] == 1].groupby('证件号码').size().reset_index(name='深夜出行次数')
        total_stats = result_deduped.groupby('证件号码').size().reset_index(name='总出行次数')
        stats_result = total_stats.merge(night_stats, on='证件号码', how='left')
        stats_result['深夜出行次数'] = stats_result['深夜出行次数'].fillna(0).astype(int)
        stats_result['深夜出行占比'] = (stats_result['深夜出行次数'] / stats_result['总出行次数'] * 100).round(2)
        
        log_callback("匹配人员库姓名...")
        if os.path.exists(person_db_path):
            try:
                # 支持 .xlsx 格式
                if person_db_path.endswith('.xlsx'):
                    personnel_df = pd.read_excel(person_db_path, engine='openpyxl')
                else:
                    personnel_df = pd.read_excel(person_db_path, engine='xlrd')
                
                if '公民身份号码' in personnel_df.columns and '姓名' in personnel_df.columns:
                    name_map = dict(zip(personnel_df['公民身份号码'].astype(str), personnel_df['姓名']))
                    stats_result['姓名'] = stats_result['证件号码'].astype(str).map(name_map)
                    stats_result['姓名'] = stats_result['姓名'].fillna('未匹配')
                    matched = (stats_result['姓名'] != '未匹配').sum()
                    log_callback(f"  ✓ 成功匹配 {matched}/{len(stats_result)} 人")
                else:
                    log_callback("  ⚠️ 人员库列名不匹配，使用'未匹配'")
                    stats_result['姓名'] = '未匹配'
            except Exception as e:
                log_callback(f"  ⚠️ 读取人员库失败: {e}")
                stats_result['姓名'] = '未匹配'
        else:
            log_callback("  ⚠️ 人员库文件不存在，姓名列将为空")
            stats_result['姓名'] = '未匹配'
        
        final_result = stats_result[['姓名', '证件号码', '深夜出行次数', '总出行次数', '深夜出行占比']].copy()
        final_result = final_result.sort_values('深夜出行次数', ascending=False)
        final_result.insert(0, '序号', range(1, len(final_result) + 1))
        
        os.makedirs(output_folder, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(output_folder, f"视频布控分析报告_{timestamp}.xlsx")
        final_result.to_excel(output_file, index=False)
        
        log_callback("\n" + "="*60)
        log_callback("✅ 处理完成！")
        log_callback("="*60)
        log_callback(f"📊 统计结果:")
        log_callback(f"   总人数: {len(final_result)}")
        log_callback(f"   深夜出行总次数: {final_result['深夜出行次数'].sum()}")
        log_callback(f"   深夜出行人数: {(final_result['深夜出行次数'] > 0).sum()}")
        log_callback(f"\n📁 结果保存至: {output_file}")
        
        return output_file
        
    except Exception as e:
        log_callback(f"❌ 处理失败: {str(e)}")
        import traceback
        log_callback(traceback.format_exc())
        return None


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("视频布控数据分析工具 v1.0")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
        # 变量
        self.input_folder = tk.StringVar()
        self.person_db = tk.StringVar()
        self.output_folder = tk.StringVar(value="C:/Users/Administrator/Desktop/分析结果")
        self.similarity = tk.DoubleVar(value=0.9)
        self.time_window = tk.IntVar(value=5)
        
        self.setup_ui()
    
    def setup_ui(self):
        # 标题
        title = tk.Label(self.root, text="视频布控数据分析工具 v1.0", 
                         font=("Microsoft YaHei", 16, "bold"), bg='#f0f0f0', fg="#333")
        title.pack(pady=10)
        
        # 主框架
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 输入区域
        input_frame = tk.LabelFrame(main_frame, text="输入设置", 
                                    font=("Microsoft YaHei", 11, "bold"),
                                    bg='#f0f0f0', fg='#333')
        input_frame.pack(fill=tk.X, pady=5)
        
        # CSV文件夹
        row1 = tk.Frame(input_frame, bg='#f0f0f0')
        row1.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row1, text="CSV文件夹:", width=12, anchor='w',
                font=("Microsoft YaHei", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        tk.Entry(row1, textvariable=self.input_folder, 
                font=("Microsoft YaHei", 10), width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(row1, text="浏览", command=self.browse_input,
                 font=("Microsoft YaHei", 10), bg="#4CAF50", fg="white", padx=10).pack(side=tk.RIGHT)
        
        # 人员库文件
        row2 = tk.Frame(input_frame, bg='#f0f0f0')
        row2.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row2, text="人员库文件:", width=12, anchor='w',
                font=("Microsoft YaHei", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        tk.Entry(row2, textvariable=self.person_db, 
                font=("Microsoft YaHei", 10), width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(row2, text="浏览", command=self.browse_person,
                 font=("Microsoft YaHei", 10), bg="#4CAF50", fg="white", padx=10).pack(side=tk.RIGHT)
        
        # 输出路径
        row3 = tk.Frame(input_frame, bg='#f0f0f0')
        row3.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row3, text="输出路径:", width=12, anchor='w',
                font=("Microsoft YaHei", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        tk.Entry(row3, textvariable=self.output_folder, 
                font=("Microsoft YaHei", 10), width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(row3, text="浏览", command=self.browse_output,
                 font=("Microsoft YaHei", 10), bg="#4CAF50", fg="white", padx=10).pack(side=tk.RIGHT)
        
        # 参数区域
        param_frame = tk.LabelFrame(main_frame, text="参数设置", 
                                    font=("Microsoft YaHei", 11, "bold"),
                                    bg='#f0f0f0', fg='#333')
        param_frame.pack(fill=tk.X, pady=5)
        
        row4 = tk.Frame(param_frame, bg='#f0f0f0')
        row4.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row4, text="相似度阈值:", width=12, anchor='w',
                font=("Microsoft YaHei", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        scale = tk.Scale(row4, from_=0.5, to=1.0, resolution=0.01, orient=tk.HORIZONTAL, 
                         variable=self.similarity, length=250, bg='#f0f0f0')
        scale.pack(side=tk.LEFT, padx=5)
        tk.Label(row4, textvariable=self.similarity, width=6,
                font=("Microsoft YaHei", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        
        row5 = tk.Frame(param_frame, bg='#f0f0f0')
        row5.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row5, text="时间窗口(秒):", width=12, anchor='w',
                font=("Microsoft YaHei", 10), bg='#f0f0f0').pack(side=tk.LEFT)
        tk.Spinbox(row5, from_=1, to=30, textvariable=self.time_window, 
                  font=("Microsoft YaHei", 10), width=10).pack(side=tk.LEFT, padx=5)
        
        # 按钮
        btn_frame = tk.Frame(main_frame, bg='#f0f0f0')
        btn_frame.pack(pady=10)
        self.run_btn = tk.Button(btn_frame, text="🚀 开始分析", command=self.run_analysis, 
                                  bg="#2196F3", fg="white", font=("Microsoft YaHei", 12, "bold"), 
                                  padx=30, pady=5, relief='flat', cursor='hand2')
        self.run_btn.pack(side=tk.LEFT, padx=10)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(pady=5, fill=tk.X)
        
        # 日志区域
        log_frame = tk.LabelFrame(main_frame, text="处理日志", 
                                  font=("Microsoft YaHei", 10, "bold"),
                                  bg='#f0f0f0', fg='#333')
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, 
                                                   font=("Consolas", 9),
                                                   bg='#2d2d2d', fg='#f8f8f2',
                                                   insertbackground='white')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 状态栏
        self.status_bar = tk.Label(self.root, text="✅ 就绪 - 请选择CSV文件夹", 
                                   font=("Microsoft YaHei", 9), bg='#e0e0e0', 
                                   fg="#666", relief="sunken", anchor='w')
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def browse_input(self):
        folder = filedialog.askdirectory(title="选择CSV文件夹")
        if folder:
            self.input_folder.set(folder)
            self.status_bar.config(text=f"✅ 已选择: {folder}")
    
    def browse_person(self):
        file = filedialog.askopenfilename(title="选择人员库文件", 
                                           filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if file:
            self.person_db.set(file)
            self.status_bar.config(text=f"✅ 已选择人员库: {os.path.basename(file)}")
    
    def browse_output(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_folder.set(folder)
            self.status_bar.config(text=f"✅ 输出路径: {folder}")
    
    def run_analysis(self):
        if not self.input_folder.get():
            messagebox.showwarning("警告", "请选择CSV文件夹")
            return
        
        if not self.person_db.get():
            reply = messagebox.askyesno("确认", "未选择人员库文件，将无法匹配姓名。是否继续？")
            if not reply:
                return
        
        if not self.output_folder.get():
            messagebox.showwarning("警告", "请选择输出路径")
            return
        
        self.run_btn.config(state=tk.DISABLED, text="⏳ 处理中...")
        self.progress.start()
        self.log_text.delete(1.0, tk.END)
        self.status_bar.config(text="⏳ 正在处理...")
        
        def run():
            result = process_data(
                self.input_folder.get(),
                self.person_db.get(),
                self.output_folder.get(),
                self.similarity.get(),
                self.time_window.get(),
                self.log
            )
            
            self.root.after(0, lambda: self.on_finished(result))
        
        threading.Thread(target=run, daemon=True).start()
    
    def on_finished(self, result_file):
        self.progress.stop()
        self.run_btn.config(state=tk.NORMAL, text="🚀 开始分析")
        if result_file:
            self.status_bar.config(text=f"✅ 完成！结果保存至: {result_file}")
            messagebox.showinfo("成功", f"分析完成！\n\n结果保存至:\n{result_file}")
        else:
            self.status_bar.config(text="❌ 处理失败，请查看日志")
    
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = App()
    app.run()