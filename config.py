import os

SQLALCHEMY_DATABASE_URI = os.getenv('DATABASE_URL', 'postgresql://postgres:2323@localhost:5432/federal_police_feedback')
SQLALCHEMY_TRACK_MODIFICATIONS = False