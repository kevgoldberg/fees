from flask import Flask
from pathlib import Path


def create_app():
    """Create and configure the Flask application."""
    # Determine templates folder relative to this package (fees/templates)
    project_dir = Path(__file__).resolve().parent.parent
    template_dir = project_dir / "templates"
    app = Flask(__name__, template_folder=str(template_dir))

    with app.app_context():
        from .routes import bp as main_blueprint
        app.register_blueprint(main_blueprint)

    return app
