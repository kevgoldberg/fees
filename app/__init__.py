from flask import Flask

app = Flask(__name__)

# Import routes to register view functions
from .routes import *
