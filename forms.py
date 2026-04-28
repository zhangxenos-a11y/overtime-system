from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SubmitField, SelectField, TextAreaField, FloatField, DateField, BooleanField
from wtforms.validators import DataRequired, Length, EqualTo, ValidationError
from models import User, Department


class LoginForm(FlaskForm):
    username = StringField('用户名', validators=[DataRequired(), Length(2, 80)])
    password = PasswordField('密码', validators=[DataRequired()])
    submit = SubmitField('登录')


class RegisterForm(FlaskForm):
    username = StringField('用户名', validators=[DataRequired(), Length(2, 80)])
    password = PasswordField('密码', validators=[DataRequired(), Length(6, 128)])
    password2 = PasswordField('确认密码', validators=[DataRequired(), EqualTo('password')])
    department_id = SelectField('部门', coerce=int, validators=[DataRequired()])
    submit = SubmitField('注册')

    def validate_username(self, username):
        user = User.query.filter_by(username=username.data).first()
        if user:
            raise ValidationError('用户名已存在')


class UserEditForm(FlaskForm):
    username = StringField('用户名', validators=[DataRequired(), Length(2, 80)])
    role = SelectField('角色', choices=[('admin', '管理员'), ('manager', '部门负责人'), ('teacher', '教师')])
    department_id = SelectField('部门', coerce=int)
    submit = SubmitField('保存')


class UserCreateForm(FlaskForm):
    username = StringField('用户名', validators=[DataRequired(), Length(2, 80)])
    role = SelectField('角色', choices=[('admin', '管理员'), ('manager', '部门负责人'), ('teacher', '教师')])
    department_id = SelectField('部门', coerce=int)
    submit = SubmitField('保存')

    def validate_username(self, username):
        user = User.query.filter_by(username=username.data).first()
        if user:
            raise ValidationError('用户名已存在')


class DepartmentForm(FlaskForm):
    name = StringField('部门名称', validators=[DataRequired(), Length(1, 50)])
    submit = SubmitField('保存')


class OvertimeForm(FlaskForm):
    content = TextAreaField('加班内容', validators=[DataRequired()])
    date = DateField('加班日期', validators=[DataRequired()])
    hours = FloatField('加班时长（小时）', validators=[DataRequired()])
    is_workday = BooleanField('是否为工作日')
    memo = TextAreaField('备忘录')
    submit = SubmitField('保存')


class OvertimeFilterForm(FlaskForm):
    department_id = SelectField('部门', coerce=int, validators=[])
    start_date = DateField('开始日期')
    end_date = DateField('结束日期')
    user_id = SelectField('人员', coerce=int, validators=[])
    submit = SubmitField('筛选')
    reset = SubmitField('重置')
