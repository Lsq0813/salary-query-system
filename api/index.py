# 纯Python工资查询系统 - 无外部依赖
from flask import Flask, render_template, request, redirect, url_for, session, flash
import os
import json
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

app = Flask(__name__, template_folder=os.path.join(os.path.dirname(__file__), '../templates'))
app.secret_key = 'your_secret_key_here'  # 请修改为更复杂的密钥

# 数据存储路径（使用Vercel可写的/tmp目录）
DATA_DIR = '/tmp/salary_data'
EMPLOYEES_FILE = os.path.join(DATA_DIR, 'employees.json')
SALARY_FILE = os.path.join(DATA_DIR, 'salary_records.json')

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# 初始化数据文件
def init_data_files():
    if not os.path.exists(EMPLOYEES_FILE):
        with open(EMPLOYEES_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f, ensure_ascii=False)
    if not os.path.exists(SALARY_FILE):
        with open(SALARY_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f, ensure_ascii=False)

init_data_files()

def format_number(value):
    try:
        if isinstance(value, str):
            value = value.replace(',', '').strip()
        num = float(value)
        if num == int(num):
            return str(int(num))
        else:
            return "%.2f" % num
    except:
        return str(value)

# 员工数据操作
def load_employees():
    with open(EMPLOYEES_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_employees(employees):
    with open(EMPLOYEES_FILE, 'w', encoding='utf-8') as f:
        json.dump(employees, f, ensure_ascii=False, indent=2)

# 工资记录操作
def load_salary_records():
    with open(SALARY_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_salary_records(records):
    with open(SALARY_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

# 纯Python读取Excel (仅支持xlsx格式，使用标准库)
def read_xlsx(file_path):
    """使用标准库读取xlsx文件，不需要openpyxl"""
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 找到工作表文件
            sheet_files = [f for f in zf.namelist() if f.startswith('xl/worksheets/sheet') and f.endswith('.xml')]
            if not sheet_files:
                raise Exception("未找到工作表")
            
            # 读取第一个工作表
            sheet_file = sheet_files[0]
            with zf.open(sheet_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
            
            # 命名空间
            ns = {
                'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }
            
            # 获取所有单元格数据
            cells = {}
            for row in root.findall('.//main:row', ns):
                row_num = int(row.get('r', '1'))
                cells[row_num] = {}
                for cell in row.findall('.//main:c', ns):
                    col = cell.get('r', 'A')[0]  # 获取列名 (A, B, C...)
                    cell_value = ''
                    
                    # 检查单元格类型
                    cell_type = cell.get('t', 'n')
                    
                    if cell_type == 's':
                        # 共享字符串
                        if cell.find('.//main:v', ns) is not None:
                            s_idx = int(cell.find('.//main:v', ns).text)
                            # 读取共享字符串表
                            with zf.open('xl/sharedStrings.xml') as ssf:
                                s_tree = ET.parse(ssf)
                                s_root = s_tree.getroot()
                                if s_root.findall('.//main:t', ns):
                                    cell_value = s_root.findall('.//main:t', ns)[s_idx].text or ''
                    elif cell_type == 'n':
                        # 数值
                        if cell.find('.//main:v', ns) is not None:
                            cell_value = cell.find('.//main:v', ns).text or ''
                    else:
                        # 其他类型直接读取
                        if cell.find('.//main:v', ns) is not None:
                            cell_value = cell.find('.//main:v', ns).text or ''
                    
                    cells[row_num][col] = cell_value
            
            # 转换为表格格式
            if not cells:
                return [], []
            
            # 获取最大行列
            max_row = max(cells.keys())
            max_col = ord(max(max(row.keys()) for row in cells.values()))
            
            # 读取表头
            headers = []
            if 1 in cells:
                for col_idx in range(ord('A'), max_col + 1):
                    col = chr(col_idx)
                    headers.append(cells[1].get(col, '').strip())
            
            # 读取数据
            data = []
            for row_num in range(2, max_row + 1):
                if row_num in cells:
                    row_data = {}
                    for col_idx, header in enumerate(headers):
                        col = chr(ord('A') + col_idx)
                        row_data[header] = cells[row_num].get(col, '')
                    data.append(row_data)
            
            return headers, data
    except Exception as e:
        raise Exception("读取Excel文件失败: " + str(e))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/query', methods=['POST'])
def query_salary():
    name = request.form.get('name', '').strip()
    card_last6 = request.form.get('card_last6', '').strip()
    
    if not name or not card_last6:
        flash('请输入姓名和银行卡号后6位')
        return redirect(url_for('index'))
    
    # 加载员工数据
    employees = load_employees()
    
    # 查找员工
    employee = None
    for emp in employees:
        if emp['name'] == name and emp['card_last6'] == card_last6:
            employee = emp
            break
    
    if not employee:
        flash('未找到该员工信息，请检查姓名和银行卡号后6位是否正确')
        return redirect(url_for('index'))
    
    # 加载工资记录
    records = load_salary_records()
    
    # 过滤该员工的工资记录
    salary_list = []
    for record in records:
        if record['employee_id'] == employee['id']:
            salary_list.append({
                'month': record['month'],
                'data': record['salary_data']
            })
    
    # 按月份倒序排序
    salary_list.sort(key=lambda x: x['month'], reverse=True)
    
    return render_template('result.html', 
                           employee_name=name,
                           salary_list=salary_list)

@app.route('/admin')
def admin():
    if 'admin_logged_in' not in session:
        return redirect(url_for('login'))
    return render_template('admin.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        # 简单的管理员验证
        if username == 'admin' and password == 'admin123':
            session['admin_logged_in'] = True
            session['admin_username'] = username
            return redirect(url_for('admin'))
        else:
            flash('用户名或密码错误')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('admin_logged_in', None)
    session.pop('admin_username', None)
    return redirect(url_for('index'))

@app.route('/upload', methods=['POST'])
def upload_salary():
    if 'admin_logged_in' not in session:
        return redirect(url_for('login'))
    
    if 'file' not in request.files:
        flash('请选择文件')
        return redirect(url_for('admin'))
    
    file = request.files['file']
    month = request.form.get('month', '').strip()
    
    if file.filename == '':
        flash('请选择文件')
        return redirect(url_for('admin'))
    
    if not month:
        flash('请输入月份（格式如：2024-01）')
        return redirect(url_for('admin'))
    
    if not file.filename.endswith('.xlsx'):
        flash('请上传Excel文件（仅支持.xlsx格式）')
        return redirect(url_for('admin'))
    
    try:
        # 保存上传的文件
        filepath = os.path.join(DATA_DIR, f'salary_{month}.xlsx')
        file.save(filepath)
        
        # 使用纯Python读取Excel
        headers, data = read_xlsx(filepath)
        
        required_cols = ['姓名', '银行卡号']
        for col in required_cols:
            if col not in headers:
                flash(f'Excel缺少必需的列：{col}')
                return redirect(url_for('admin'))
        
        # 加载现有数据
        employees = load_employees()
        records = load_salary_records()
        
        # 生成员工ID映射
        employee_id_map = {}
        for emp in employees:
            employee_id_map["%s_%s" % (emp['name'], emp['card_last6'])] = emp['id']
        
        # 生成新ID
        def get_new_employee_id():
            if not employees:
                return 1
            return max(emp['id'] for emp in employees) + 1
        
        # 处理工资数据
        processed_count = 0
        
        # 删除该月份的旧记录
        records = [r for r in records if r['month'] != month]
        
        for row in data:
            name = str(row.get('姓名', '')).strip()
            card_number = str(row.get('银行卡号', '')).strip()
            card_last6 = card_number[-6:] if len(card_number) >= 6 else card_number
            
            if not name or name == 'nan':
                continue
            
            key = "%s_%s" % (name, card_last6)
            
            # 获取或创建员工ID
            if key not in employee_id_map:
                employee_id = get_new_employee_id()
                new_employee = {
                    'id': employee_id,
                    'name': name,
                    'card_last6': card_last6
                }
                employees.append(new_employee)
                employee_id_map[key] = employee_id
            else:
                employee_id = employee_id_map[key]
            
            # 构建工资数据
            salary_dict = {}
            for col in headers:
                if col not in ['姓名', '银行卡号']:
                    value = row.get(col)
                    if value is not None and value != '':
                        salary_dict[col] = format_number(value)
            
            # 添加工资记录
            new_record = {
                'id': len(records) + 1,
                'employee_id': employee_id,
                'month': month,
                'salary_data': salary_dict
            }
            records.append(new_record)
            
            processed_count += 1
        
        # 保存数据
        save_employees(employees)
        save_salary_records(records)
        
        flash(f'成功上传 {processed_count} 条工资记录，月份：{month}')
        
    except Exception as e:
        flash(f'上传失败：{str(e)}')
    
    return redirect(url_for('admin'))

# Vercel Serverless Function入口
def handler(event, context):
    return app

# 腾讯云函数入口（保留兼容）
def main_handler(event, context):
    return app(event, context)

if __name__ == '__main__':
    import os
    # 从环境变量获取端口
    port = int(os.environ.get('PORT', 9000))
    app.run(debug=False, host='0.0.0.0', port=port)
