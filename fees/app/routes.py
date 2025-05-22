from flask import Blueprint, request, render_template, send_file, abort
from pathlib import Path
from .workbook_processor import process_workbook

bp = Blueprint('main', __name__)

@bp.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Handle two file uploads
        query_file = request.files.get('query_file')
        report_file = request.files.get('report_file')
        week = int(request.form.get('week', 0) or 0)
        # Validate inputs
        if not query_file or not query_file.filename:
            abort(400, 'Query file is required')
        if not report_file or not report_file.filename:
            abort(400, 'Report file is required')
        # Create raw directories at project root
        project_root = Path(__file__).resolve().parents[3]
        qdir = project_root / 'Query'; qdir.mkdir(exist_ok=True)
        rdir = project_root / 'Report'; rdir.mkdir(exist_ok=True)
        # Save raw files
        qpath = qdir / query_file.filename
        query_file.save(str(qpath))
        rpath = rdir / report_file.filename
        report_file.save(str(rpath))
        # Process the query workbook
        result = process_workbook(open(str(qpath), 'rb'), week)
        return send_file(
            result,
            as_attachment=True,
            download_name='processed.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return render_template('index.html')