import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
CLIENT_SECRET_FILENAME = 'client_secret.json'
MOVING_AVERAGE_WINDOW = 2
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, CLIENT_SECRET_FILENAME)

VISITORS_COL = 'Visitors'
MOVING_AVERAGE_COL = 'Moving Average'
