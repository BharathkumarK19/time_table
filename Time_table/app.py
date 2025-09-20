from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify
import random, os, subprocess, traceback

# Import the necessary functions and global variables from your module
from timetable_generator import assign_subjects_for_faculty, export_all, empty_table_for_shift, FREE_DAY_SETTINGS, normalize_token, ensure_div_table


app = Flask(__name__)
app.config["RESULT_FOLDER"] = os.path.join(os.getcwd(), "results")

# Memory store
ftables = {}
dtables = {}

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/builder", methods=["GET"])
def builder():
    return render_template("builder.html")

@app.route("/generate", methods=["POST"])
def generate():
    print(">>> /generate called (POST)")
    global ftables, dtables
    ftables, dtables = {}, {}
    random.seed(7)

    try:
        data = request.get_json(force=True)
        faculties = data.get("faculties", [])
        if not faculties:
            return jsonify(ok=False, error="No faculties provided"), 400

        # Collect and set universal headers
        university = data.get("university", "")
        department = data.get("department", "")
        academic = data.get("academic", "")

        # Populate FREE_DAY_SETTINGS from JSON before scheduling
        FREE_DAY_SETTINGS.clear()
        for f in faculties:
            for subj in f.get("Subjects", []):
                sem = subj.get("Semester")
                div = subj.get("Division")
                holidays = subj.get("Holidays", [])
                
                if sem and div and holidays:
                    key = (str(sem), normalize_token(div))
                    FREE_DAY_SETTINGS[key] = holidays

        # Initialize and apply holidays to all timetables
        for f in faculties:
            fname = f["Name"]
            fshift = f["Shift"]
            if fname not in ftables:
                ftables[fname] = empty_table_for_shift(fshift)

            for subj in f.get("Subjects", []):
                sem = subj.get("Semester")
                div = subj.get("Division")
                dshift = subj.get("Div_Shift", "8-3")
                dtbl, _ = ensure_div_table(dtables, sem, div, dshift)
        
        from timetable_generator import apply_free_day_markings_from_inputs
        apply_free_day_markings_from_inputs(dtables, faculties)

        # Assign subjects
        for f in faculties:
            assign_subjects_for_faculty(f, ftables, dtables)
        
        # Export results
        os.makedirs(app.config["RESULT_FOLDER"], exist_ok=True)
        prev_cwd = os.getcwd()
        try:
            os.chdir(app.config["RESULT_FOLDER"])
            export_all(ftables, dtables, faculties,
                       university=university,
                       department=department,
                       academic=academic)
        finally:
            os.chdir(prev_cwd)

        return jsonify(ok=True, message="Generated", redirect=url_for("success"))

    except Exception as e:
        tb = traceback.format_exc()
        print("ERROR in /generate:", e)
        print(tb)
        return jsonify(ok=False, error=str(e), traceback=tb.splitlines()[-12:]), 500


@app.route("/success")
def success():
    files = []
    if os.path.exists(app.config["RESULT_FOLDER"]):
        files = sorted(os.listdir(app.config["RESULT_FOLDER"]))
    return render_template("success.html", files=files)

@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(app.config["RESULT_FOLDER"], filename, as_attachment=True)

if __name__ == "__main__":
  

    try:
        app.run(host="0.0.0.0", port=5000, debug=True, threaded=True)
    finally:
            pass