from flask import Flask
from pathlib import Path

def create_app():
    # Determine templates folder (fees/fees/templates)
    project_dir = Path(__file__).resolve().parents[2]
    template_dir = project_dir / 'templates'
    app = Flask(__name__, template_folder=str(template_dir))

    with app.app_context():
        from .routes import bp as main_blueprint
        app.register_blueprint(main_blueprint)

    return app