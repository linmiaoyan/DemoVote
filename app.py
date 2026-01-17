from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, current_user
from datetime import datetime, timedelta
import qrcode
from io import BytesIO
import pandas as pd
import os
from werkzeug.security import generate_password_hash, check_password_hash
import secrets
from PIL import Image, ImageDraw, ImageFont
import math
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfutils import ImageReader
from reportlab.lib.pagesizes import A4, portrait
import time
import threading
import queue
import atexit
from sqlalchemy.orm import scoped_session, sessionmaker
import logging
import socket
from dotenv import load_dotenv

# 加载 .env 文件
load_dotenv()

logger = logging.getLogger(__name__)
import json

# 获取项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, 'instance')
os.makedirs(INSTANCE_DIR, exist_ok=True)

# 配置：优先从 .env 文件读取，如果没有则使用环境变量或默认值
PUBLIC_HOST = os.getenv('PUBLIC_HOST', '')  # 如果为空，将动态获取
SECRET_KEY = os.getenv('SECRET_KEY', secrets.token_hex(32))
ADMIN_GATE_KEY = os.getenv('ADMIN_GATE_KEY', 'wzkjgz')
DATABASE_PATH = os.path.join(INSTANCE_DIR, 'votes.db')
HOST = os.getenv('HOST', '0.0.0.0')
PORT = int(os.getenv('PORT', 5005))
DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'

app = Flask(__name__)
app.config['SECRET_KEY'] = SECRET_KEY
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DATABASE_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['ADMIN_GATE_KEY'] = ADMIN_GATE_KEY

# 设置时区为北京时间
def get_current_time():
    return datetime.utcnow() + timedelta(hours=8)

def get_local_ip():
    """获取本机IP地址"""
    try:
        # 创建一个UDP socket来获取本机IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        # 不实际发送数据，只是用来获取本机IP
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            # 备用方法：通过hostname获取
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
            return ip
        except Exception:
            return '127.0.0.1'

def get_public_host():
    """动态获取PUBLIC_HOST，如果未设置则从请求中获取"""
    if PUBLIC_HOST:
        return PUBLIC_HOST
    # 从请求中动态获取（如果可用）
    try:
        from flask import has_request_context, request as req
        if has_request_context() and req:
            return f"{req.scheme}://{req.host}/"
    except:
        pass
    # 默认值（用于生成二维码时）
    return f"http://localhost:{PORT}/"

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'index'

# 数据模型
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(120), nullable=False)
    is_admin = db.Column(db.Boolean, default=False)
    qr_code = db.Column(db.String(200), unique=True)
    votes = db.relationship('Vote', backref='user', lazy=True)
    subjective_answers = db.relationship(
        'SubjectiveAnswer',
        backref='user',
        lazy=True,
        cascade="all, delete-orphan"
    )

class Survey(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    type = db.Column(db.String(20), nullable=False)  # 'single_choice' 或 'table'
    introduction = db.Column(db.Text, nullable=True)  # 保留：问卷简介
    subjective_question_prompt = db.Column(db.Text, nullable=True) # 新增：主观题说明文字
    created_at = db.Column(db.DateTime, default=get_current_time)
    is_active = db.Column(db.Boolean, default=True)
    option_limits = db.Column(db.JSON, nullable=True)  # 新增：选项限制，格式为 {"A": 7, "B": 7, ...}
    table_option_count = db.Column(db.Integer, default=3)  # 新增：表格问卷选项数量，默认3
    enable_quick_fill = db.Column(db.Boolean, default=True)  # 新增：是否启用快填功能，默认启用
    questions = db.relationship('Question', backref='survey', lazy=True)
    qr_codes = db.relationship('QRCode', backref='survey', lazy=True)

class Question(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    survey_id = db.Column(db.Integer, db.ForeignKey('survey.id'), nullable=False)
    content = db.Column(db.Text, nullable=False)
    option_count = db.Column(db.Integer, nullable=True)  # 单选题的选项数量
    created_at = db.Column(db.DateTime, default=get_current_time)
    votes = db.relationship('Vote', backref='question', lazy=True)

class TableRespondent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    survey_id = db.Column(db.Integer, db.ForeignKey('survey.id'), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    created_at = db.Column(db.DateTime, default=get_current_time)
    survey = db.relationship('Survey', backref='table_respondents', lazy=True)

class QRCode(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    survey_id = db.Column(db.Integer, db.ForeignKey('survey.id'), nullable=False)
    token = db.Column(db.String(200), unique=True, nullable=False)
    is_used = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=get_current_time)

class Vote(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('question.id'), nullable=False)
    table_respondent_id = db.Column(db.Integer, db.ForeignKey('table_respondent.id'), nullable=True)
    score = db.Column(db.Text, nullable=False)  # 改为Text以支持长文本回答
    created_at = db.Column(db.DateTime, default=get_current_time)
    table_respondent = db.relationship('TableRespondent', backref='votes', lazy=True)

class SubjectiveAnswer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    survey_id = db.Column(db.Integer, db.ForeignKey('survey.id'), nullable=False)
    content = db.Column(db.Text, nullable=True) # 主观回答内容，可以为空
    created_at = db.Column(db.DateTime, default=get_current_time)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

# 路由
@app.route('/')
def index():
    if current_user.is_authenticated and current_user.is_admin:
        surveys = Survey.query.filter_by(is_active=True).all()
        return render_template('admin.html', surveys=surveys)
    surveys = Survey.query.filter_by(is_active=True).all()
    return render_template('index.html', surveys=surveys)

@app.route('/admin_login', methods=['GET'])
def admin_login():
    # 通过 GET 参数密钥校验，仅持有密钥者可进入后台，无需 POST 账号密码
    provided_key = request.args.get('k', '')
    if not provided_key or provided_key != app.config['ADMIN_GATE_KEY']:
        flash('非法访问', 'danger')
        return redirect(url_for('thank_you'))
    session['is_admin'] = True
    return redirect(url_for('admin'))

def ensure_admin_session():
    if not session.get('is_admin'):
        return redirect(url_for('thank_you'))
    return None


@app.route('/admin')
def admin():
    guard = ensure_admin_session()
    if guard:
        return guard
    surveys = Survey.query.filter_by(is_active=True).all()
    
    # 计算每个问卷的数据条数
    survey_stats = []
    for survey in surveys:
        # 计算投票数据条数
        if survey.type == 'single_choice':
            vote_count = Vote.query.join(Question).filter(Question.survey_id == survey.id).count()
        elif survey.type == 'table':
            vote_count = Vote.query.join(Question).join(TableRespondent).filter(
                Question.survey_id == survey.id,
                TableRespondent.survey_id == survey.id
            ).count()
        else:
            vote_count = 0
        
        # 计算主观题回答数
        subjective_count = SubjectiveAnswer.query.filter_by(survey_id=survey.id).count()
        
        # 总数据条数
        total_count = vote_count + subjective_count
        
        survey_stats.append({
            'survey': survey,
            'vote_count': vote_count,
            'subjective_count': subjective_count,
            'total_count': total_count
        })
    
    return render_template('admin.html', survey_stats=survey_stats)

@app.route('/admin/create_survey', methods=['POST'])
def create_survey():
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey_type = request.form.get('survey_type')
    survey_name = request.form.get('survey_name')
    survey_introduction = request.form.get('survey_introduction')
    subjective_question_prompt = request.form.get('subjective_question_prompt')
    table_option_count = int(request.form.get('table_option_count', 3))
    
    if not survey_name or not survey_type:
        flash('请填写问卷名称并选择类型', 'danger')
        return redirect(url_for('admin'))
    
    enable_quick_fill = request.form.get('enable_quick_fill') == 'on'
    survey = Survey(
        name=survey_name, 
        type=survey_type, 
        introduction=survey_introduction, 
        subjective_question_prompt=subjective_question_prompt,
        table_option_count=table_option_count if survey_type == 'table' else None,
        enable_quick_fill=enable_quick_fill
    )
    db.session.add(survey)
    db.session.commit()

    if survey_type == 'single_choice':
        return redirect(url_for('create_single_choice_questions', survey_id=survey.id))
    else:
        return redirect(url_for('create_table_questions', survey_id=survey.id))

@app.route('/admin/create_single_choice_questions/<int:survey_id>', methods=['GET', 'POST'])
def create_single_choice_questions(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_single':
            content = request.form.get('content')
            option_count = int(request.form.get('option_count', 4))
            
            if content:
                question = Question(
                    survey_id=survey_id,
                    content=content,
                    option_count=option_count
                )
                db.session.add(question)
                db.session.commit()
                flash('问题添加成功', 'success')
            else:
                flash('问题内容不能为空', 'danger')
        elif action == 'import_list':
            question_list_text = request.form.get('question_list')
            option_count_batch = int(request.form.get('option_count_batch', 4))
            
            if question_list_text:
                questions_content = [q.strip() for q in question_list_text.split('\n') if q.strip()]
                for content in questions_content:
                    question = Question(
                        survey_id=survey.id,
                        content=content,
                        option_count=option_count_batch
                    )
                    db.session.add(question)
                db.session.commit()
                flash(f'{len(questions_content)}个问题已成功导入', 'success')
            else:
                flash('导入列表不能为空', 'danger')
    
    questions = Question.query.filter_by(survey_id=survey_id).all()
    return render_template('create_single_choice.html', survey=survey, questions=questions)

@app.route('/admin/create_table_questions/<int:survey_id>', methods=['GET', 'POST'])
def create_table_questions(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    if request.method == 'POST':
        content = request.form.get('content')
        
        if content:
            question = Question(
                survey_id=survey_id,
                content=content,
                option_count=None
            )
            db.session.add(question)
            db.session.commit()
            flash('问题添加成功', 'success')
    
    questions = Question.query.filter_by(survey_id=survey_id).all()
    return render_template('create_table.html', survey=survey, questions=questions)

@app.route('/admin/manage_table_respondents/<int:survey_id>', methods=['GET', 'POST'])
def manage_table_respondents(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard

    survey = Survey.query.get_or_404(survey_id)

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_single':
            name = request.form.get('name')
            if name:
                respondent = TableRespondent(survey_id=survey_id, name=name)
                db.session.add(respondent)
                db.session.commit()
                flash('人名添加成功', 'success')
            else:
                flash('人名不能为空', 'danger')
        elif action == 'import_list':
            name_list_text = request.form.get('name_list')
            if name_list_text:
                names = [n.strip() for n in name_list_text.split('\n') if n.strip()]
                for name in names:
                    respondent = TableRespondent(survey_id=survey_id, name=name)
                    db.session.add(respondent)
                db.session.commit()
                flash(f'{len(names)}个人名已成功导入', 'success')
            else:
                flash('导入列表不能为空', 'danger')

    respondents = TableRespondent.query.filter_by(survey_id=survey_id).all()
    return render_template('manage_table_respondents.html', survey=survey, respondents=respondents)

@app.route('/admin/generate_qr/<int:survey_id>', methods=['POST'])
def generate_qr(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    num_users = int(request.form.get('num_users', 0))
    
    if num_users <= 0:
        flash('请输入有效的用户数量', 'danger')
        return redirect(url_for('admin'))
    
    # 生成二维码
    qr_codes = []
    for _ in range(num_users):
        token = secrets.token_urlsafe(16)
        qr_code = QRCode(survey_id=survey_id, token=token)
        db.session.add(qr_code)
        qr_codes.append(token)
    
    db.session.commit()
    
    # 生成二维码图片
    qr_images = []
    # 动态获取主机地址
    public_host = get_public_host()
    for token in qr_codes:
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(f"{public_host}login/{token}")
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        # 添加问卷名称
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("msyh.ttf", 20)
        except IOError:
            try:
                font = ImageFont.truetype("simhei.ttf", 20) # 备用字体
            except IOError:
                font = ImageFont.load_default()
                from flask import current_app
                current_app.logger.warning("无法加载中文字体 (msyh.ttf, simhei.ttf)。问卷名称可能无法正确显示或显示为方框。")

        # 在二维码下方添加问卷名称
        text_width = draw.textlength(survey.name, font=font)
        img_width = img.size[0]
        # Adjust y_pos for text to be slightly above the bottom border
        draw.text(((img_width - text_width) // 2, img.size[1] - 30), 
                 survey.name, font=font, fill='black')
        
        qr_images.append(img)
    
    # 创建PDF文件
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    
    # 定义每页的二维码布局
    cols = 4
    rows = 4 # 每页固定显示4x4个二维码
    
    # 计算每个二维码可用的最大正方形尺寸，并考虑页面边距
    margin = 20 # 页面边距
    available_width = A4[0] - 2 * margin
    available_height = A4[1] - 2 * margin
    
    cell_width = available_width / cols
    cell_height = available_height / rows
    qr_size_on_page = min(cell_width, cell_height) # 确保二维码是正方形
    
    for page in range(math.ceil(len(qr_images) / (cols * rows))):
        start_idx = page * (cols * rows)
        end_idx = min((page + 1) * (cols * rows), len(qr_images))
        page_qr_images = qr_images[start_idx:end_idx]
        
        for idx, img in enumerate(page_qr_images):
            row_in_page = idx // cols
            col_in_page = idx % cols
            
            # 计算二维码在页面上的位置，并居中
            x_pos = margin + col_in_page * cell_width + (cell_width - qr_size_on_page) / 2
            y_pos = A4[1] - margin - (row_in_page + 1) * cell_height + (cell_height - qr_size_on_page) / 2
            
            # 将PIL图像转换为PDF可用的格式，并通过ImageReader传递
            img_buffer = BytesIO()
            img.save(img_buffer, format='PNG')
            img_reader = ImageReader(img_buffer)
            
            # 在PDF中放置二维码，保持正方形比例
            c.drawImage(img_reader, 
                       x_pos, 
                       y_pos,
                       width=qr_size_on_page,
                       height=qr_size_on_page)
        
        c.showPage()
    
    c.save()
    pdf_buffer.seek(0)
    
    return send_file(
        pdf_buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=f'qr_codes_{survey.name}.pdf'
    )

@app.route('/login/<token>')
def login_with_qr(token):
    qr = QRCode.query.filter_by(token=token).first()
    if not qr:
        flash('无效的二维码', 'danger')
        return redirect(url_for('thank_you'))
    
    # 查找或创建用户
    user = User.query.filter_by(qr_code=token).first()
    if not user:
        # 创建新用户
        user = User(
            username=f"user_{token[:8]}",
            password_hash=generate_password_hash(token),
            qr_code=token
        )
        db.session.add(user)
        db.session.commit()
    
    login_user(user)
    return redirect(url_for('vote', survey_id=qr.survey_id))

@app.route('/preview/<int:survey_id>')
def preview_survey(survey_id):
    """预览问卷（不需要验证二维码，仅用于管理员预览）"""
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    questions = Question.query.filter_by(survey_id=survey_id).order_by(Question.id).all()
    
    # 如果是表格问卷，获取所有受访者
    respondents = []
    if survey.type == 'table':
        respondents = TableRespondent.query.filter_by(survey_id=survey_id).order_by(TableRespondent.id).all()
    
    table_option_count = survey.table_option_count if survey.type == 'table' else None
    
    return render_template(
        'vote.html',
        survey=survey,
        questions=questions,
        respondents=respondents,
        subjective_question_prompt=survey.subjective_question_prompt,
        table_option_count=table_option_count,
        enable_quick_fill=survey.enable_quick_fill,
        is_preview=True
    )

@app.route('/vote/<int:survey_id>')
@login_required
def vote(survey_id):
    survey = Survey.query.get_or_404(survey_id)
    questions = Question.query.filter_by(survey_id=survey_id).all()
    
    respondents = []
    if survey.type == 'table':
        respondents = TableRespondent.query.filter_by(survey_id=survey_id).all()
        
    table_option_count = survey.table_option_count if survey.type == 'table' else None
    
    return render_template(
        'vote.html',
        survey=survey,
        questions=questions,
        respondents=respondents,
        subjective_question_prompt=survey.subjective_question_prompt,
        table_option_count=table_option_count,
        enable_quick_fill=survey.enable_quick_fill
    )

@app.route('/admin/set_option_limits/<int:survey_id>', methods=['POST'])
def set_option_limits(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    # 获取选项限制
    option_limits = {}
    for option in 'ABCDE':
        limit = request.form.get(f'limit_{option}')
        if limit and limit.strip():
            try:
                limit_value = int(limit)
                if limit_value > 0:
                    option_limits[option] = limit_value
            except ValueError:
                flash(f'选项 {option} 的限制值必须是正整数', 'danger')
                return redirect(url_for('create_single_choice_questions', survey_id=survey_id))
    
    # 更新问卷的选项限制
    survey.option_limits = option_limits
    db.session.commit()
    
    flash('选项限制设置已保存', 'success')
    return redirect(url_for('create_single_choice_questions', survey_id=survey_id))



submit_queue = queue.Queue()

def db_worker():
    with app.app_context():
        while True:
            try:
                func, args, kwargs = submit_queue.get()
                func(*args, **kwargs)
                submit_queue.task_done()
            except Exception as e:
                logger.error(f"数据库写入失败: {e}", exc_info=True)
                submit_queue.task_done()  # 确保即使出错也标记任务完成

threading.Thread(target=db_worker, daemon=True).start()

def save_vote_to_db(vote_data, retry_count=0):
    """保存投票到数据库，带重试机制
    
    Args:
        vote_data: 投票数据字典
        retry_count: 当前重试次数（默认最大重试3次）
    """
    MAX_RETRIES = 3
    Session = scoped_session(sessionmaker(bind=db.engine))
    session = Session()
    try:
        survey_id = vote_data['survey_id']
        user_id = vote_data['user_id']
        survey = session.get(Survey, survey_id)
        if not survey:
            logger.error(f"问卷不存在: survey_id={survey_id}")
            return
        
        # 删除旧投票
        question_ids = [q.id for q in session.query(Question).filter_by(survey_id=survey_id).all()]
        session.query(Vote).filter(Vote.user_id == user_id, Vote.question_id.in_(question_ids)).delete(synchronize_session='fetch')
        session.query(SubjectiveAnswer).filter_by(user_id=user_id, survey_id=survey_id).delete(synchronize_session='fetch')
        # 插入新投票
        if survey.type == 'single_choice':
            for q_id, score in vote_data['single_choice_votes']:
                vote = Vote(user_id=user_id, question_id=q_id, score=score)
                session.add(vote)
        elif survey.type == 'table':
            for q_id, respondent_id, score in vote_data['table_votes']:
                vote = Vote(user_id=user_id, question_id=q_id, table_respondent_id=respondent_id, score=score)
                session.add(vote)
        if vote_data.get('subjective_answer'):
            subjective_answer = SubjectiveAnswer(user_id=user_id, survey_id=survey_id, content=vote_data['subjective_answer'])
            session.add(subjective_answer)
        
        # 提交事务
        session.commit()
        if retry_count > 0:
            logger.info(f"投票数据成功写入（经过 {retry_count} 次重试）: user_id={user_id}, survey_id={survey_id}")
    except Exception as e:
        logger.error(f"数据库写入异常: user_id={vote_data['user_id']}, survey_id={vote_data['survey_id']}, 重试次数={retry_count}, 错误: {e}", exc_info=True)
        session.rollback()
        
        # 如果未超过最大重试次数，则重新入队
        if retry_count < MAX_RETRIES:
            try:
                submit_queue.put_nowait((save_vote_to_db, (vote_data, retry_count + 1), {}))
                time.sleep(0.5 * (retry_count + 1))  # 指数退避
            except queue.Full:
                logger.error(f"队列已满，无法重试: user_id={vote_data['user_id']}, survey_id={vote_data['survey_id']}")
        else:
            logger.error(f"达到最大重试次数，放弃写入: user_id={vote_data['user_id']}, survey_id={vote_data['survey_id']}")
    finally:
        session.close()

@app.route('/submit_vote/<int:survey_id>', methods=['POST'])
@login_required
def submit_vote(survey_id):
    survey = Survey.query.get_or_404(survey_id)
    
    # 校验逻辑
    if survey.type == 'single_choice':
        questions = Question.query.filter_by(survey_id=survey_id).all()
        for question in questions:
            if f'question_{question.id}' not in request.form or not request.form[f'question_{question.id}']:
                flash('请完成所有问题后再进行提交', 'danger')
                return redirect(url_for('vote', survey_id=survey_id))
        if survey.option_limits:
            option_counts = {}
            for question_id, score in request.form.items():
                if question_id.startswith('question_'):
                    option = score
                    option_counts[option] = option_counts.get(option, 0) + 1
            for option, limit in survey.option_limits.items():
                if option_counts.get(option, 0) > limit:
                    flash(f'选项 {option} 的选择次数超过了限制 ({limit}次)', 'danger')
                    return redirect(url_for('vote', survey_id=survey_id))
    elif survey.type == 'table':
        questions = Question.query.filter_by(survey_id=survey_id).all()
        respondents = TableRespondent.query.filter_by(survey_id=survey_id).all()
        for question in questions:
            for respondent in respondents:
                if f'vote_{question.id}_{respondent.id}' not in request.form or not request.form[f'vote_{question.id}_{respondent.id}']:
                    flash('请完成所有问题后再进行提交', 'danger')
                    return redirect(url_for('vote', survey_id=survey_id))
    
    # 打包投票数据
    vote_data = {
        'survey_id': survey_id,
        'user_id': current_user.id,
        'single_choice_votes': [],
        'table_votes': [],
        'subjective_answer': None
    }
    
    if survey.type == 'single_choice':
        for question_id, score in request.form.items():
            if question_id.startswith('question_'):
                q_id = int(question_id.split('_')[1])
                vote_data['single_choice_votes'].append((q_id, score))
    elif survey.type == 'table':
        for key, score in request.form.items():
            if key.startswith('vote_'):
                parts = key.split('_')
                q_id = int(parts[1])
                respondent_id = int(parts[2])
                vote_data['table_votes'].append((q_id, respondent_id, score))
    
    if survey.subjective_question_prompt:
        subjective_answer_content = request.form.get('subjective_answer', '').strip()
        if subjective_answer_content:
            vote_data['subjective_answer'] = subjective_answer_content
    
    # 将投票数据入队等待写入数据库
    submit_queue.put((save_vote_to_db, (vote_data,), {}))
    flash('您的投票已提交成功！', 'success')
    return redirect(url_for('thank_you'))

@app.route('/thank_you')
def thank_you():
    return render_template('thank_you.html')

@app.route('/admin/results/<int:survey_id>')
def view_results(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    # 获取投票数据
    votes_data = []
    if survey.type == 'single_choice':
        votes = Vote.query.join(Question).filter(Question.survey_id == survey_id).order_by(Vote.created_at.desc()).all()
        for vote in votes:
            votes_data.append({
                'user': vote.user.username,
                'question': vote.question.content,
                'option': vote.score,
                'time': vote.created_at
            })
    elif survey.type == 'table':
        votes = Vote.query.join(Question).join(TableRespondent).filter(
            Question.survey_id == survey_id, 
            TableRespondent.survey_id == survey_id
        ).order_by(Vote.created_at.desc()).all()
        
        for vote in votes:
            votes_data.append({
                'user': vote.user.username,
                'question': vote.question.content,
                'respondent': vote.table_respondent.name if vote.table_respondent else None,
                'option': vote.score,
                'time': vote.created_at
            })
    
    # 获取主观题回答
    subjective_answers = SubjectiveAnswer.query.filter_by(survey_id=survey_id).order_by(SubjectiveAnswer.created_at.desc()).all()
    subjective_data = []
    for ans in subjective_answers:
        subjective_data.append({
            'user': ans.user.username,
            'content': ans.content,
            'time': ans.created_at
        })
    
    # 统计数据
    total_votes = len(votes_data)
    unique_users = len(set(v['user'] for v in votes_data))
    unique_respondents = len(set(v.get('respondent') for v in votes_data if v.get('respondent'))) if survey.type == 'table' else 0
    total_questions = len(survey.questions)
    total_subjective_answers = len(subjective_data)
    
    return render_template('view_results.html', 
                         survey=survey,
                         votes_data=votes_data,
                         subjective_answers=subjective_data,
                         total_votes=total_votes,
                         unique_users=unique_users,
                         unique_respondents=unique_respondents,
                         total_questions=total_questions,
                         total_subjective_answers=total_subjective_answers)

@app.route('/admin/download_results/<int:survey_id>')
def download_results(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    # 创建原始数据
    data = []
    if survey.type == 'single_choice':
        votes = Vote.query.join(Question).filter(Question.survey_id == survey_id).all()
        for vote in votes:
            data.append({
                '用户': vote.user.username,
                '问题': vote.question.content.replace(' ', '-'),  # 替换空格为连字符
                '选项': vote.score,
                '时间': vote.created_at
            })
    elif survey.type == 'table':
        votes = Vote.query.join(Question).join(TableRespondent).filter(
            Question.survey_id == survey_id, 
            TableRespondent.survey_id == survey_id
        ).all()
        
        for vote in votes:
            data.append({
                '用户': vote.user.username,
                '问题': vote.question.content.replace(' ', '-'),  # 替换空格为连字符
                '人名': vote.table_respondent.name,
                '选项': vote.score,
                '时间': vote.created_at
            })
    
    # 包含主观题回答
    subjective_answers = SubjectiveAnswer.query.filter_by(survey_id=survey_id).all()
    if subjective_answers:
        for ans in subjective_answers:
            data.append({
                '用户': ans.user.username,
                '问题': (survey.subjective_question_prompt if survey.subjective_question_prompt else "主观题回答").replace(' ', '-'),
                '人名': None, # 为主观题回答添加人名，设置为None
                '选项': ans.content,
                '时间': ans.created_at
            })

    # 定义DataFrame的列名，以确保所有类型的数据都有正确的列
    columns = ['用户', '问题', '选项', '时间']
    if survey.type == 'table':
        columns.insert(2, '人名') # 在'问题'和'选项'之间插入'人名'

    # 创建DataFrame
    df = pd.DataFrame(data, columns=columns)
    
    # 创建Excel文件
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. 原始数据（按用户排序）
        df.to_excel(writer, sheet_name='原始数据', index=False)
        
        # 2. 按问题排列的数据
        if survey.type == 'single_choice':
            # 对于单选题，按问题和选项排序
            sort_cols = []
            if '问题' in df.columns: sort_cols.append('问题')
            if '选项' in df.columns: sort_cols.append('选项')
            df_sorted = df.sort_values(sort_cols) if sort_cols else df
            df_sorted.to_excel(writer, sheet_name='按问题排列', index=False)
            
            # 3. 统计结果 - 新的格式
            # 获取所有唯一的问题和选项
            questions = df['问题'].unique()
            options = df['选项'].unique()
            
            # 创建统计结果DataFrame
            stats_data = []
            for question in questions:
                row_data = {'问题': question}
                # 计算每个选项的出现次数
                for option in options:
                    count = len(df[(df['问题'] == question) & (df['选项'] == option)])
                    row_data[f'{option}'] = count
                stats_data.append(row_data)
            
            stats_df = pd.DataFrame(stats_data)
            stats_df.to_excel(writer, sheet_name='统计结果', index=False)
            
        elif survey.type == 'table':
            # 对于表格题，按问题、人名和选项排序
            sort_cols = []
            if '问题' in df.columns: sort_cols.append('问题')
            if '人名' in df.columns: sort_cols.append('人名')
            if '选项' in df.columns: sort_cols.append('选项')
            df_sorted = df.sort_values(sort_cols) if sort_cols else df
            df_sorted.to_excel(writer, sheet_name='按问题排列', index=False)
            
            # 3. 统计结果 - 新的格式
            # 获取所有唯一的问题、人名和选项
            questions = df['问题'].unique()
            respondents = df['人名'].unique()
            options = list('ABCDE')[:survey.table_option_count]
            
            # 创建统计结果DataFrame
            stats_data = []
            for question in questions:
                for respondent in respondents:
                    row_data = {'问题': question, '人名': respondent}
                    # 计算每个选项的出现次数
                    for option in options:
                        count = len(df[(df['问题'] == question) & \
                                     (df['人名'] == respondent) & \
                                     (df['选项'] == option)])
                        row_data[f'{option}'] = count
                    stats_data.append(row_data)
            
            stats_df = pd.DataFrame(stats_data)
            stats_df.to_excel(writer, sheet_name='统计结果', index=False)
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'vote_results_{survey.name}.xlsx'
    )

@app.route('/admin/delete_results/<int:survey_id>', methods=['POST'])
def delete_results(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    try:
        # 删除所有投票数据
        if survey.type == 'single_choice':
            votes = Vote.query.join(Question).filter(Question.survey_id == survey_id).all()
        elif survey.type == 'table':
            votes = Vote.query.join(Question).join(TableRespondent).filter(
                Question.survey_id == survey_id,
                TableRespondent.survey_id == survey_id
            ).all()
        
        for vote in votes:
            db.session.delete(vote)
        
        # 删除主观题回答
        subjective_answers = SubjectiveAnswer.query.filter_by(survey_id=survey_id).all()
        for answer in subjective_answers:
            db.session.delete(answer)
        
        db.session.commit()
        flash(f'已成功删除问卷 "{survey.name}" 的所有投票数据', 'success')
    except Exception as e:
        db.session.rollback()
        logger.error(f"删除投票数据失败: {e}")
        flash('删除投票数据失败，请重试', 'danger')
    
    return redirect(url_for('view_results', survey_id=survey_id))

@app.route('/admin/copy_survey/<int:survey_id>', methods=['POST'])
def copy_survey(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    original_survey = Survey.query.get_or_404(survey_id)
    
    try:
        # 创建新问卷
        new_survey = Survey(
            name=f"{original_survey.name} (副本)",
            type=original_survey.type,
            introduction=original_survey.introduction,
            subjective_question_prompt=original_survey.subjective_question_prompt,
            option_limits=original_survey.option_limits.copy() if original_survey.option_limits else None,
            table_option_count=original_survey.table_option_count,
            enable_quick_fill=original_survey.enable_quick_fill
        )
        db.session.add(new_survey)
        db.session.flush()  # 获取新问卷的ID
        
        # 复制所有问题
        original_questions = Question.query.filter_by(survey_id=survey_id).all()
        for orig_question in original_questions:
            new_question = Question(
                survey_id=new_survey.id,
                content=orig_question.content,
                option_count=orig_question.option_count
            )
            db.session.add(new_question)
        
        # 如果是表格问卷，复制所有人名
        if original_survey.type == 'table':
            original_respondents = TableRespondent.query.filter_by(survey_id=survey_id).all()
            for orig_respondent in original_respondents:
                new_respondent = TableRespondent(
                    survey_id=new_survey.id,
                    name=orig_respondent.name
                )
                db.session.add(new_respondent)
        
        db.session.commit()
        flash(f'问卷已复制为新问卷："{new_survey.name}"', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'复制问卷失败: {e}', 'danger')
    
    return redirect(url_for('admin'))

@app.route('/admin/delete_survey/<int:survey_id>', methods=['POST'])
def delete_survey(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    
    survey = Survey.query.get_or_404(survey_id)
    
    try:
        # 删除与问卷相关的所有投票记录
        Vote.query.filter(Vote.question.has(survey_id=survey_id)).delete(synchronize_session='fetch')
        # 删除与问卷相关的所有问题
        Question.query.filter_by(survey_id=survey_id).delete(synchronize_session='fetch')
        # 删除与问卷相关的所有人名（如果问卷是表格类型）
        TableRespondent.query.filter_by(survey_id=survey_id).delete(synchronize_session='fetch')
        # 删除与问卷相关的所有二维码
        QRCode.query.filter_by(survey_id=survey_id).delete(synchronize_session='fetch')
        
        # 最后删除问卷本身
        db.session.delete(survey)
        db.session.commit()
        flash(f'问卷 "{survey.name}" 及其所有相关数据已删除', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'删除问卷失败: {e}', 'danger')
        
    return redirect(url_for('admin'))

@app.route('/admin/edit_survey_title/<int:survey_id>', methods=['POST'])
def edit_survey_title(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    survey = Survey.query.get_or_404(survey_id)
    new_title = request.form.get('new_title')
    if new_title:
        survey.name = new_title
        db.session.commit()
        flash('问卷标题已更新', 'success')
    else:
        flash('标题不能为空', 'danger')
    return redirect(url_for('admin'))

@app.route('/admin/edit_survey/<int:survey_id>')
def edit_survey(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    survey = Survey.query.get_or_404(survey_id)
    questions = Question.query.filter_by(survey_id=survey_id).order_by(Question.id).all()
    respondents = []
    if survey.type == 'table':
        respondents = TableRespondent.query.filter_by(survey_id=survey_id).order_by(TableRespondent.id).all()
    return render_template('edit_survey.html', survey=survey, questions=questions, respondents=respondents)

@app.route('/admin/update_survey_info/<int:survey_id>', methods=['POST'])
def update_survey_info(survey_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    survey = Survey.query.get_or_404(survey_id)
    
    survey.name = request.form.get('survey_name', '').strip()
    survey.introduction = request.form.get('survey_introduction', '').strip() or None
    survey.subjective_question_prompt = request.form.get('subjective_question_prompt', '').strip() or None
    survey.enable_quick_fill = request.form.get('enable_quick_fill') == 'on'
    
    if survey.type == 'table':
        table_option_count = int(request.form.get('table_option_count', 3))
        survey.table_option_count = table_option_count
    
    if not survey.name:
        flash('问卷名称不能为空', 'danger')
        return redirect(url_for('edit_survey', survey_id=survey_id))
    
    db.session.commit()
    flash('问卷基本信息已更新', 'success')
    return redirect(url_for('edit_survey', survey_id=survey_id))

@app.route('/admin/update_question/<int:question_id>', methods=['POST'])
def update_question(question_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    question = Question.query.get_or_404(question_id)
    survey = question.survey
    
    content = request.form.get('content', '').strip()
    if not content:
        flash('问题内容不能为空', 'danger')
        return redirect(url_for('edit_survey', survey_id=survey.id))
    
    question.content = content
    
    # 如果是单选题，更新选项数量
    if survey.type == 'single_choice':
        option_count = int(request.form.get('option_count', 4))
        question.option_count = option_count
    
    db.session.commit()
    flash('问题已更新', 'success')
    return redirect(url_for('edit_survey', survey_id=survey.id))

@app.route('/admin/delete_question/<int:question_id>', methods=['POST'])
def delete_question(question_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    question = Question.query.get_or_404(question_id)
    survey_id = question.survey_id
    
    # 删除与该问题相关的所有投票记录
    Vote.query.filter_by(question_id=question_id).delete()
    
    # 删除问题
    db.session.delete(question)
    db.session.commit()
    
    flash('问题已删除', 'success')
    return redirect(url_for('edit_survey', survey_id=survey_id))

@app.route('/admin/delete_respondent/<int:respondent_id>', methods=['POST'])
def delete_respondent(respondent_id):
    guard = ensure_admin_session()
    if guard:
        return guard
    respondent = TableRespondent.query.get_or_404(respondent_id)
    survey_id = respondent.survey_id
    
    # 删除与该人名相关的所有投票记录
    Vote.query.filter_by(table_respondent_id=respondent_id).delete()
    
    # 删除人名
    db.session.delete(respondent)
    db.session.commit()
    
    flash('人名已删除', 'success')
    return redirect(url_for('edit_survey', survey_id=survey_id))

@app.route('/admin/batch_delete_questions', methods=['POST'])
def batch_delete_questions():
    guard = ensure_admin_session()
    if guard:
        return guard
    
    question_ids = request.form.getlist('question_ids')
    survey_id = request.form.get('survey_id')
    
    if not question_ids:
        flash('请选择要删除的问题', 'warning')
        return redirect(url_for('edit_survey', survey_id=survey_id) if survey_id else url_for('admin'))
    
    if not survey_id:
        # 如果没有提供survey_id，从第一个问题获取
        first_question = Question.query.get(int(question_ids[0]))
        if not first_question:
            flash('问题不存在', 'danger')
            return redirect(url_for('admin'))
        survey_id = first_question.survey_id
    
    try:
        survey_id = int(survey_id)
        
        # 验证所有问题都属于同一个问卷
        question_id_list = [int(qid) for qid in question_ids]
        questions = Question.query.filter(Question.id.in_(question_id_list)).all()
        
        # 验证问题数量和survey_id
        if len(questions) != len(question_id_list):
            flash('部分问题不存在', 'danger')
            return redirect(url_for('edit_survey', survey_id=survey_id))
        
        for question in questions:
            if question.survey_id != survey_id:
                flash('不能批量删除不同问卷的问题', 'danger')
                return redirect(url_for('edit_survey', survey_id=survey_id))
        
        # 删除与这些问题相关的所有投票记录
        Vote.query.filter(Vote.question_id.in_(question_id_list)).delete(synchronize_session='fetch')
        
        # 删除选中的问题
        for question in questions:
            db.session.delete(question)
        
        db.session.commit()
        flash(f'已成功删除 {len(questions)} 个问题', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'批量删除失败: {e}', 'danger')
        survey_id = survey_id if 'survey_id' in locals() else None
    
    return redirect(url_for('edit_survey', survey_id=survey_id) if survey_id else url_for('admin'))


if __name__ == '__main__':
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    with app.app_context():
        db.create_all()
        
        # 数据库迁移：添加缺失的列
        try:
            from sqlalchemy import inspect, text
            inspector = inspect(db.engine)
            columns = [col['name'] for col in inspector.get_columns('survey')]
            
            # 检查并添加 enable_quick_fill 列（如果不存在）
            if 'enable_quick_fill' not in columns:
                logger.info("检测到数据库需要迁移：添加 enable_quick_fill 列")
                # SQLite 中 BOOLEAN 存储为 INTEGER (0 或 1)
                db.session.execute(text('ALTER TABLE survey ADD COLUMN enable_quick_fill INTEGER DEFAULT 1'))
                db.session.commit()
                logger.info("数据库迁移完成：已添加 enable_quick_fill 列")
        except Exception as e:
            logger.warning(f"数据库迁移检查失败（可能是新数据库）: {e}")
            db.session.rollback()
        
        # 创建管理员账号
        if not User.query.filter_by(username='admin').first():
            admin = User(
                username='admin',
                password_hash=generate_password_hash('admin123'),
                is_admin=True
            )
            db.session.add(admin)
            db.session.commit()
            logger.info("管理员账号已创建: admin / admin123")
    
    # 获取实际IP地址用于显示
    display_host = get_local_ip() if HOST == '0.0.0.0' else HOST
    
    logger.info(f"启动服务器: http://{HOST}:{PORT}")
    logger.info(f"数据库路径: {DATABASE_PATH}")
    admin_url = f"http://{display_host}:{PORT}/admin_login?k={ADMIN_GATE_KEY}"
    logger.info(f"管理员入口: {admin_url}")
    
    # 在控制台醒目输出管理员入口地址
    print("\n" + "="*60)
    print("="*60)
    print(f"  管理员登录入口地址:")
    print(f"  {admin_url}")
    print("="*60)
    print("="*60 + "\n")
    
    app.run(host=HOST, port=PORT, debug=DEBUG, use_reloader=False)