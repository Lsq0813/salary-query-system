# 纯Python工资查询系统 - 主入口文件
# 此文件仅作为入口，所有逻辑在 api/__init__.py 中实现
from api import app

if __name__ == '__main__':
    import os
    # 从环境变量获取端口，默认使用9000
    port = int(os.environ.get('PORT', 9000))
    app.run(debug=False, host='0.0.0.0', port=port)
