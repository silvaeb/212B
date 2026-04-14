import os


class BaseConfig:
    SECRET_KEY = os.getenv('SECRET_KEY', 'change-me')
    SQLALCHEMY_DATABASE_URI = os.getenv('DATABASE_URL', 'sqlite:///sistema_pdrlog.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SESSION_COOKIE_SECURE = False
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
    EXCEL_PATH = os.getenv('EXCEL_PATH', 'Extra PDRLOG.xlsx')
    DEBUG = False


class DevelopmentConfig(BaseConfig):
    DEBUG = True
    SESSION_COOKIE_SECURE = False


class ProductionConfig(BaseConfig):
    DEBUG = False
    SESSION_COOKIE_SECURE = True


def get_config_class():
    app_env = (os.getenv('APP_ENV') or os.getenv('FLASK_ENV') or 'development').strip().lower()
    if app_env == 'production':
        return ProductionConfig
    return DevelopmentConfig
