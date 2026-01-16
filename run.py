#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
VoteSite 独立运行启动脚本
"""
import os
import sys

# 确保使用当前目录
if __name__ == '__main__':
    # 添加当前目录到 Python 路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    # 导入并运行应用
    from app import app, db, User
    from werkzeug.security import generate_password_hash
    import logging
    
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger(__name__)
    
    # 初始化数据库
    with app.app_context():
        db.create_all()
        # 创建管理员账号（如果不存在）
        if not User.query.filter_by(username='admin').first():
            admin = User(
                username='admin',
                password_hash=generate_password_hash('admin123'),
                is_admin=True
            )
            db.session.add(admin)
            db.session.commit()
            logger.info("管理员账号已创建: admin / admin123")
        else:
            logger.info("管理员账号已存在")
    
    # 获取配置
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'False').lower() == 'true'
    admin_key = os.environ.get('ADMIN_GATE_KEY', 'wzkjgz')
    
    logger.info("=" * 60)
    logger.info("VoteSite 投票系统")
    logger.info("=" * 60)
    logger.info(f"服务器地址: http://{host}:{port}")
    logger.info(f"管理员入口: http://{host}:{port}/admin_login?k={admin_key}")
    logger.info(f"调试模式: {'开启' if debug else '关闭'}")
    logger.info("=" * 60)
    
    # 运行应用
    app.run(host=host, port=port, debug=debug, use_reloader=False)

