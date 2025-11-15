# models.py
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    roll_no = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(150))
    email = db.Column(db.String(200))
    password = db.Column(db.String(200))

class Question(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    text = db.Column(db.Text, nullable=False)
    opt_a = db.Column(db.String(300))
    opt_b = db.Column(db.String(300))
    opt_c = db.Column(db.String(300))
    opt_d = db.Column(db.String(300))
    correct = db.Column(db.String(5))   # 'A'/'B'/'C'/'D'
    per_question_time = db.Column(db.Integer, default=30)  # seconds
    order_index = db.Column(db.Integer, default=0)

class Attempt(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.Integer, db.ForeignKey('student.id'))
    started_at = db.Column(db.DateTime, default=datetime.utcnow)
    finished_at = db.Column(db.DateTime, nullable=True)
    score = db.Column(db.Integer, nullable=True)

class Answer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey('attempt.id'))
    question_id = db.Column(db.Integer, db.ForeignKey('question.id'))
    selected = db.Column(db.String(5))
    answered_at = db.Column(db.DateTime, default=datetime.utcnow)
