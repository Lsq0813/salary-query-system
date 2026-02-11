# Vercel Serverless Function入口
from . import app

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