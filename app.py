"""Web UI that turns a list of people into Moodle import files (CSV, PDF and XLSX)."""
import os
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, abort, jsonify, render_template, request, send_from_directory

import moodle_import as mi

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

DEFAULT_PASSWORD = os.getenv("DEFAULT_PASSWORD", "ChangeMe123!")
CSV_DELIMITER = os.getenv("CSV_DELIMITER", ";").strip() or ";"
COURSES_FILE = BASE_DIR / os.getenv("COURSES_FILE", "courses.txt")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


def output_folders():
    return {kind: mi.resolve_folder(os.getenv(f"{kind.upper()}_DIR", f"output/{kind}"))
            for kind in ("pdf", "csv", "xlsx")}


def load_courses():
    """Read courses.txt: one course per line, lines starting with # are comments."""
    if not COURSES_FILE.exists():
        return []
    lines = COURSES_FILE.read_text(encoding="utf-8").splitlines()
    return [l.strip() for l in lines if l.strip() and not l.startswith("#")]


def save_courses(courses):
    """Rewrite courses.txt keeping the comment lines."""
    comments = []
    if COURSES_FILE.exists():
        comments = [l for l in COURSES_FILE.read_text(encoding="utf-8").splitlines() if l.startswith("#")]
    COURSES_FILE.write_text("\n".join(comments + courses) + "\n", encoding="utf-8")


@app.get("/")
def index():
    return render_template("index.html", courses=load_courses())


@app.get("/courses")
def courses_page():
    return render_template("courses.html", courses=load_courses())


@app.route("/api/courses", methods=["POST", "DELETE"])
def change_courses():
    course = mi.clean((request.get_json(silent=True) or {}).get("course"))
    if not course:
        return jsonify(error="Enter the course name."), 400

    courses = load_courses()
    if request.method == "POST":
        if course.casefold() in (c.casefold() for c in courses):
            return jsonify(error=f'The course "{course}" is already in the list.'), 409
        courses.append(course)
    else:
        courses = [c for c in courses if c != course]
    save_courses(courses)
    return jsonify(courses=courses)


@app.post("/api/read-sheet")
def read_sheet():
    file = request.files.get("file")
    if not file or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="Upload the registration sheet as .xlsx"), 400
    try:
        return jsonify(rows=mi.read_registration_sheet(file.stream))
    except Exception as exc:  # spreadsheet outside the expected layout
        return jsonify(error=f"Could not read the spreadsheet: {exc}"), 400


@app.post("/api/generate")
def generate():
    rows = (request.get_json(silent=True) or {}).get("rows", [])
    people = mi.build_people(rows, load_courses())
    if not people:
        return jsonify(error="No users filled in."), 400

    errors = [{"row": i, "errors": p.errors} for i, p in enumerate(people, start=1) if p.errors]
    if errors:
        return jsonify(error="Fix the highlighted rows before generating the files.", rows=errors), 422

    try:
        saved = mi.save_files(people, output_folders(), DEFAULT_PASSWORD, CSV_DELIMITER)
    except OSError as exc:
        return jsonify(error=f"Failed to save the files: {exc}"), 500

    return jsonify(
        total=len(people),
        delimiter=CSV_DELIMITER,
        files={kind: {"name": path.name, "path": str(path)} for kind, path in saved.items()},
    )


@app.get("/download/<kind>/<path:name>")
def download(kind, name):
    if kind not in ("pdf", "csv", "xlsx") or Path(name).name != name or not name.endswith(f".{kind}"):
        abort(404)
    return send_from_directory(output_folders()[kind].resolve(), name, as_attachment=True)


if __name__ == "__main__":
    app.run(host=os.getenv("HOST", "127.0.0.1"), port=int(os.getenv("PORT", "8533")),
            debug=os.getenv("FLASK_DEBUG") == "1")
