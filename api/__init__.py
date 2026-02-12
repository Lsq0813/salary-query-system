# 纯Python工资查询系统 - 无外部依赖
from flask import Flask, render_template, request, redirect, url_for, session, flash
import os
import json
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

app = Flask(__name__, 
           template_folder=os.path.join(os.path.dirname(__file__), '../templates'),
           static_folder=os.path.join(os.path.dirname(__file__), '../static'))
# 从环境变量获取密钥，确保安全性
app.secret_key = os.environ.get('SECRET_KEY', 'your_secret_key_here_change_in_production')

# 数据存储路径（支持 Railway 持久化卷）
# 在 Railway 上使用 /data 卷，本地使用项目目录
DATA_DIR = os.environ.get('RAILWAY_VOLUME_MOUNT_PATH', os.path.join(os.path.dirname(__file__), 'data'))
EMPLOYEES_FILE = os.path.join(DATA_DIR, 'employees.json')
SALARY_FILE = os.path.join(DATA_DIR, 'salary_records.json')

# 确保数据目录存在
if not os.path.exists(DATA_DIR):
    try:
        os.makedirs(DATA_DIR)
        print(f"创建数据目录: {DATA_DIR}")
    except Exception as e:
        print(f"创建目录失败: {e}，使用当前目录")
        # 如果创建目录失败，使用当前目录作为备选
        DATA_DIR = os.getcwd()
        EMPLOYEES_FILE = os.path.join(DATA_DIR, 'employees.json')
        SALARY_FILE = os.path.join(DATA_DIR, 'salary_records.json')

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
    try:
        if not os.path.exists(EMPLOYEES_FILE):
            return []
        with open(EMPLOYEES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"加载员工数据失败: {str(e)}, 路径: {EMPLOYEES_FILE}")
        return []

def save_employees(employees):
    try:
        # 确保目录存在
        os.makedirs(os.path.dirname(EMPLOYEES_FILE), exist_ok=True)
        with open(EMPLOYEES_FILE, 'w', encoding='utf-8') as f:
            json.dump(employees, f, ensure_ascii=False, indent=2)
        print(f"保存员工数据成功: {len(employees)} 条记录, 路径: {EMPLOYEES_FILE}")
        return True
    except Exception as e:
        print(f"保存员工数据失败: {str(e)}, 路径: {EMPLOYEES_FILE}")
        return False

# 工资记录操作
def load_salary_records():
    try:
        if not os.path.exists(SALARY_FILE):
            return []
        with open(SALARY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"加载工资记录失败: {str(e)}, 路径: {SALARY_FILE}")
        return []

def save_salary_records(records):
    try:
        # 确保目录存在
        os.makedirs(os.path.dirname(SALARY_FILE), exist_ok=True)
        with open(SALARY_FILE, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        print(f"保存工资记录成功: {len(records)} 条记录, 路径: {SALARY_FILE}")
        return True
    except Exception as e:
        print(f"保存工资记录失败: {str(e)}, 路径: {SALARY_FILE}")
        return False

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
            
            # 辅助函数：从单元格引用中提取列名
            def get_column_name(cell_ref):
                # 提取列名部分（如从'A1'提取'A'，从'AA12'提取'AA'）
                col_name = ''
                for char in cell_ref:
                    if char.isalpha():
                        col_name += char
                    else:
                        break
                return col_name
            
            # 获取所有单元格数据
            cells = {}
            for row in root.findall('.//main:row', ns):
                row_num = int(row.get('r', '1'))
                cells[row_num] = {}
                for cell in row.findall('.//main:c', ns):
                    cell_ref = cell.get('r', 'A1')
                    col = get_column_name(cell_ref)  # 获取完整列名 (A, B, AA, AB...)
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
            
            # 获取最大行
            max_row = max(cells.keys())
            
            # 收集所有实际存在的列名并排序
            all_columns = set()
            for row in cells.values():
                all_columns.update(row.keys())
            
            # 按Excel列顺序排序（A, B, ..., Z, AA, AB, ...）
            def sort_columns(cols):
                def col_key(col):
                    # 将列名转换为数字以便排序
                    num = 0
                    for char in col:
                        num = num * 26 + (ord(char.upper()) - ord('A') + 1)
                    return num
                return sorted(cols, key=col_key)
            
            sorted_columns = sort_columns(all_columns)
            
            # 读取表头
            headers = []
            if 1 in cells:
                for col in sorted_columns:
                    headers.append(cells[1].get(col, '').strip())
            
            # 读取数据
            data = []
            for row_num in range(2, max_row + 1):
                if row_num in cells:
                    row_data = {}
                    for col_idx, header in enumerate(headers):
                        col = sorted_columns[col_idx]
                        row_data[header] = cells[row_num].get(col, '')
                    data.append(row_data)
            
            return headers, data
    except Exception as e:
        raise Exception("读取Excel文件失败: " + str(e))

@app.route('/')
def index():
    # 获取所有已上传的月份（去重并排序）
    records = load_salary_records()
    available_months = sorted(set(record['month'] for record in records), reverse=True)
    return render_template('index.html', available_months=available_months)

@app.route('/query', methods=['POST'])
def query_salary():
    name = request.form.get('name', '').strip()
    card_last6 = request.form.get('card_last6', '').strip()
    month = request.form.get('month', '').strip()
    
    if not name or not card_last6 or not month:
        flash('请输入姓名、银行卡号后6位和选择月份')
        return redirect(url_for('index'))
    
    # 显示查询信息和数据存储路径
    print(f"查询请求 - 姓名: {name}, 卡号后6位: {card_last6}, 月份: {month}")
    print(f"数据存储路径: {DATA_DIR}")
    print(f"员工数据文件: {EMPLOYEES_FILE}")
    print(f"工资记录文件: {SALARY_FILE}")
    
    # 加载员工数据
    employees = load_employees()
    print(f"加载到的员工数: {len(employees)}")
    
    # 查找员工
    employee = None
    for emp in employees:
        print(f"员工: {emp['name']}, 卡号后6位: {emp['card_last6']}")
        if emp['name'] == name and emp['card_last6'] == card_last6:
            employee = emp
            print(f"找到匹配员工: {employee}")
            break
    
    if not employee:
        flash('未找到该员工信息，请检查姓名和银行卡号后6位是否正确')
        return redirect(url_for('index'))
    
    # 加载工资记录
    records = load_salary_records()
    print(f"加载到的工资记录数: {len(records)}")
    
    # 打印所有工资记录
    for i, record in enumerate(records):
        print(f"工资记录 {i+1}: 员工ID: {record['employee_id']}, 月份: {record['month']}")
    
    # 过滤该员工的特定月份工资记录
    employee_salary_data = None
    for record in records:
        if record['employee_id'] == employee['id'] and record['month'] == month:
            employee_salary_data = record['salary_data']
            print(f"找到匹配的工资记录: {record}")
            break
    
    if not employee_salary_data:
        flash(f'未找到该员工{month}月份的工资记录')
        print(f"未找到工资记录 - 员工ID: {employee['id']}, 月份: {month}")
        return redirect(url_for('index'))
    
    # 获取该月份所有员工的工资项目并集（统一显示所有项目）
    # 使用第一个记录的列顺序作为基准顺序
    base_columns = []
    all_salary_items = set()
    for record in records:
        if record['month'] == month:
            if not base_columns and 'salary_columns' in record:
                base_columns = record['salary_columns']
            all_salary_items.update(record['salary_data'].keys())
    
    # 过滤掉不需要显示的项目
    filtered_items = {item for item in all_salary_items if item not in ['序号']}
    
    # 构建最终的项目顺序：先按原始列顺序，再添加其他项目
    salary_items = []
    # 1. 先添加原始列顺序中的项目
    for col in base_columns:
        if col in filtered_items and col not in ['序号']:
            salary_items.append(col)
            filtered_items.discard(col)
    # 2. 添加剩余的项目（按原始顺序）
    salary_items.extend(sorted(filtered_items))
    
    # 构建完整的工资数据，缺失的项目显示为 "-"
    complete_salary_data = {}
    for item in salary_items:
        complete_salary_data[item] = employee_salary_data.get(item, '-')
    
    print(f"查询成功 - 找到工资记录，包含 {len(complete_salary_data)} 个项目")
    print(f"工资项目: {list(complete_salary_data.keys())}")
    return render_template('result.html', 
                           employee_name=name,
                           month=month,
                           salary_data=complete_salary_data)

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
    month = request.form.get('month_value', '').strip()
    
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
            
            # 构建工资数据 - 保存所有项目（包括值为0或空的），保持原始列顺序
            salary_dict = {}
            salary_columns = []  # 保存列顺序
            for col in headers:
                if col not in ['姓名', '银行卡号']:
                    salary_columns.append(col)
                    value = row.get(col)
                    # 即使值为空或0，也保存项目名称，显示为 "-" 或 "0"
                    if value is not None and value != '':
                        salary_dict[col] = format_number(value)
                    else:
                        salary_dict[col] = '-'
            
            # 添加工资记录
            new_record = {
                'id': len(records) + 1,
                'employee_id': employee_id,
                'month': month,
                'salary_data': salary_dict,
                'salary_columns': salary_columns  # 保存列顺序
            }
            records.append(new_record)
            
            processed_count += 1
        
        # 保存数据
        try:
            # 显示当前数据存储路径
            print(f"当前数据存储路径: {DATA_DIR}")
            print(f"员工数据文件: {EMPLOYEES_FILE}")
            print(f"工资记录文件: {SALARY_FILE}")
            
            # 保存员工数据
            emp_save_success = save_employees(employees)
            # 保存工资记录
            record_save_success = save_salary_records(records)
            
            if emp_save_success and record_save_success:
                # 验证保存结果
                saved_employees = load_employees()
                saved_records = load_salary_records()
                
                print(f"验证结果 - 员工数: {len(saved_employees)}, 工资记录数: {len(saved_records)}")
                
                flash(f'成功上传 {processed_count} 条工资记录，月份：{month}')
                flash(f'数据存储路径: {DATA_DIR}')
            else:
                flash(f'上传失败：数据保存过程中出现错误')
                flash(f'数据存储路径: {DATA_DIR}')
        except Exception as e:
            flash(f'保存数据失败：{str(e)}')
            print(f"上传过程异常: {str(e)}")
        
    except Exception as e:
        flash(f'上传失败：{str(e)}')
    
    return redirect(url_for('admin'))