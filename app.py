from flask import Flask, render_template, request, send_file, redirect, session
import os
import glob
from Report_modules.axis_performance import generate_report as axis_performance
from Report_modules.axis_npa_cc_dpr import generate_report as axis_dpr
from Report_modules.axis_bucket import generate_report as axis_bucket
from Report_modules.bob_dpr import generate_report as generate_bob
from Report_modules.axis_whatsapp import generate_axis_message
from Report_modules.cmd_void_cleaner import generate_cmd_clean_report
from Scanner_modules.document_scanner import generate_document_pdf
from Report_modules.separate_portfolio import generate_separate_portfolio
from Report_modules.payment_dump_clean import generate_payment_dump
from datetime import timedelta



app = Flask(__name__)
app.secret_key = "12345"
app.permanent_session_lifetime = timedelta(minutes=15)

# ======================
# LOGIN
# ======================
USERNAME = "Ankit"
PASSWORD = "Creditas@2026"

REPORTS = {
    "axis_performance": "Axis NPA CC Performance",
    "axis_dpr": "Axis NPA CC DPR",
    "axis_bucket": "Axis Bucket DPR",
    "bob_dpr": "BOB Bucket DPR",
    "axis_collated": "Axis Bucket Collated Allocation",
    "axis_whatsapp": "Axis NPA CC Whatsapp message",
    "payment_dump_clean": "Payment Dump Clean",
    "cmd_clean": "Number Cleaner Report",
    "document_scanner":"Document Scanner",
    "separate_portfolio": "Separate Portfolio"
}

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

def cleanup_files():

    for file in glob.glob("uploads/*"):

        try:
            os.remove(file)

        except:
            pass


def cleanup_output():

    for file in glob.glob("output/*"):

        try:
            os.remove(file)

        except:
            pass


os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

cleanup_files()
cleanup_output()


# ======================
# LOGIN
# ======================
@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        if username == USERNAME and password == PASSWORD:
            session.permanent = True
            session["user"] = username
            return redirect("/")

        return render_template("login.html", error="Wrong Credentials")

    return render_template("login.html")


# ======================
# LOGOUT
# ======================
@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect("/login")


# ======================
# HOME
# ======================
@app.route("/")
def home():

    if "user" not in session:
        return redirect("/login")

    return render_template("home.html", reports=REPORTS,username=session.get("user"))


# ======================
# REPORT PAGE
# ======================
@app.route("/report/<report_id>")
def report(report_id):

    if "user" not in session:
        return redirect("/login")

    return render_template(
        "report.html",
        report_id=report_id,
        report_name=REPORTS.get(report_id)
    )


# ======================
# GENERATE REPORT
# ======================
@app.route("/generate/<report_id>", methods=["POST"])
def generate(report_id):

    if "user" not in session:
        return redirect("/login")

    try:

        password = request.form.get("password", "")

        # ================= AXIS PERFORMANCE =================
        if report_id == "axis_performance":

            file = request.files["file"]

            path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(path)

            output = os.path.join(OUTPUT_FOLDER, "axis_performance.xlsx")

            axis_performance(path, password, output)
            cleanup_files()
            session["success"] = "✅ Report Generated Successfully"

            return send_file(output, as_attachment=True)


        # ================= AXIS DPR =================
        elif report_id == "axis_dpr":

            file = request.files["file"]

            path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(path)

            output = os.path.join(OUTPUT_FOLDER, "axis_dpr.xlsx")

            axis_dpr(path, password, output)
            cleanup_files()
            session["success"] = "✅ Report Generated Successfully"

            return send_file(output, as_attachment=True)


        # ================= AXIS BUKET (MULTI FILE) =================
        elif report_id == "axis_bucket":

            files = request.files.getlist("files")

            month = int(request.form["month"])
            year = int(request.form["year"])

            saved_files = []

            for f in files:
                path = os.path.join(UPLOAD_FOLDER, f.filename)
                f.save(path)
                saved_files.append(path)

            output = os.path.join(OUTPUT_FOLDER, "axis_bucket.xlsx")

            axis_bucket(saved_files, password, month, year, output)
            cleanup_files()
            session["success"] = "✅ Report Generated Successfully"

            return send_file(output, as_attachment=True)


        # ================= BOB DPR (MULTI FILE) =================
        elif report_id == "bob_dpr":

            files = request.files.getlist("files")

            password = request.form["password"]
            month = request.form["month"]
            year = request.form["year"]

            if not password or password.strip() == "":
                return render_template(
                    "report.html",
                    report_id=report_id,
                    report_name=REPORTS.get(report_id),
                    error="❌ Password cannot be empty"
                )

            saved_files = []

            for f in files:
                path = os.path.join(UPLOAD_FOLDER, f.filename)
                f.save(path)
                saved_files.append(path)

            output = os.path.join(OUTPUT_FOLDER, "BOB_DPR.xlsx")

            # IMPORTANT: wrong password handle inside module OR here
            
            try:
                generate_bob(saved_files, password, month, year, output)
                print("File Exists:", os.path.exists(output))
                cleanup_files()

                return send_file(output, as_attachment=True)

            except Exception as e:

                print("BOB ERROR =>", e)

                return render_template(
                    "report.html",
                    report_id=report_id,
                    report_name=REPORTS.get(report_id),
                    error=f"❌ {str(e)}"
                )
        # ================= number clean =================
        elif report_id == "cmd_clean":

            file = request.files["file"]

            path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(path)

            output = os.path.join(OUTPUT_FOLDER, "cmd_clean.xlsx")

            generate_cmd_clean_report(path, output)

            cleanup_files()

            return send_file(output, as_attachment=True)

        # ================= Axis Whatsapp messange =================
        elif report_id == "axis_whatsapp":

            file = request.files["file"]
            password = request.form["password"]

            path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(path)

            message = generate_axis_message(path, password)

            return render_template("report.html",
                           report_id=report_id,
                           report_name=REPORTS.get(report_id),
                           message=message)

        # ================= PAYMENT DUMP CLEAN =================
        elif report_id == "payment_dump_clean":

            file = request.files["file"]

            path = os.path.join(
                UPLOAD_FOLDER,
                file.filename
            )

            file.save(path)

            output = os.path.join(
                OUTPUT_FOLDER,
                "Payment_Dump.xlsx"
            )

            generate_payment_dump(
                path,
                output
            )

            cleanup_files()

            return send_file(
                output,
                as_attachment=True
            )

        # ================= SEPARATE PORTFOLI =================
        elif report_id == "separate_portfolio":

            file = request.files["file"]

            path = os.path.join(
                UPLOAD_FOLDER,
                file.filename
            )

            file.save(path)

            separate_files = request.form.get(
                "separate_files"
            )

            if separate_files == "yes":

                output = os.path.join(
                    OUTPUT_FOLDER,
                    "Separate_Portfolio.zip"
                )

                generate_separate_portfolio(
                    path,
                    output,
                    separate_files=True
                )

            else:

                output = os.path.join(
                    OUTPUT_FOLDER,
                    "Separate_Portfolio.xlsx"
                )

                generate_separate_portfolio(
                    path,
                    output,
                    separate_files=False
                )

            cleanup_files()

            return send_file(
                output,
                as_attachment=True
            )

        # ================= DOCUMENT SCANNER =================
        elif report_id == "document_scanner":

            files = request.files.getlist("files")

            saved_files = []

            for f in files:

                if f.filename == "":
                    continue

                path = os.path.join(UPLOAD_FOLDER, f.filename)
                f.save(path)
                saved_files.append(path)
            output=os.path.join(
                OUTPUT_FOLDER,
                "Document.pdf"
            )

            if len(saved_files) == 0:

                return render_template(
                    "report.html",
                    report_id=report_id,
                    report_name=REPORTS.get(report_id),
                    error="❌ Please select at least one image."
                )

            output = os.path.join(
                OUTPUT_FOLDER,
                "Document_Scanner.pdf"
            )

            generate_document_pdf(
                saved_files,
                output
            )

            cleanup_files()

            return send_file(
                output,
                as_attachment=True
            )

        
        return "Invalid Report"

    except Exception as e:

        import traceback

        traceback.print_exc()

        return render_template(
            "report.html",
            report_id=report_id,
            report_name=REPORTS.get(report_id),
            error=f"❌ {str(e)}"
        )


if __name__ == "__main__":
    app.run(debug=True)