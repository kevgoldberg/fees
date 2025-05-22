from flask import Flask

app = Flask(__name__)

# Import routes to register view functions if available.  The tests import
# ``app.workbook_processor`` without providing a routes module, so guard the
# import to avoid ``ModuleNotFoundError`` during testing.
try:  # pragma: no cover - routes are optional during unit tests
    from .routes import *  # noqa: F401,F403
except ImportError:
    # ``routes`` is optional in some contexts (e.g. unit tests)
    pass
