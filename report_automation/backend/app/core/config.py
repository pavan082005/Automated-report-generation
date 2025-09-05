from pydantic import BaseSettings


class Settings(BaseSettings):
    # SMTP settings
    SMTP_HOST: str = "smtp.gmail.com"
    SMTP_PORT: int = 587
    SMTP_USER: str = "your_email@gmail.com"
    SMTP_PASS: str = "your_app_password"
    FROM_EMAIL: str = "your_email@gmail.com"

    class Config:
        env_file = ".env"   # loads variables from .env if present


# global settings object
settings = Settings()
