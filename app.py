import io
import os
import uuid
from datetime import datetime

import openpyxl
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session

from converter import add_to_master, convert, ReportParseError

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-secret-change-me")

MAX_CONTENT_LENGTH = 20 * 1024 * 1024  # 20 MB
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH

pending_workbooks = {}
master_workbooks = {}


def workbook_preview(workbook_bytes, max_rows=20, max_columns=8):
    workbook = openpyxl.load_workbook(io.BytesIO(workbook_bytes), read_only=True, data_only=True)
    sheets = []
    for worksheet in workbook.worksheets:
        rows = []
        for row in worksheet.iter_rows(min_row=1, max_row=max_rows, max_col=max_columns, values_only=True):
            rows.append(["" if value is None else str(value) for value in row])
        sheets.append({"name": worksheet.title, "rows": rows})
    workbook.close()
    return sheets


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        converted=None,
        master_available=session.get("master_id") in master_workbooks,
    )


@app.route("/convert", methods=["POST"])
def convert_route():
    file = request.files.get("pdf_file")

    if file is None or file.filename == "":
        flash("Please choose a PDF file to upload.")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith(".pdf"):
        flash("Only .pdf files are supported.")
        return redirect(url_for("index"))

    pdf_bytes = file.read()

    try:
        added_at = datetime.now()
        xlsx_bytes = convert(
            pdf_bytes,
            source_filename=file.filename,
            added_at=added_at,
        )
    except ReportParseError as e:
        flash(str(e))
        return redirect(url_for("index"))
    except Exception:
        flash("Something went wrong while converting this PDF. Please double-check the file and try again.")
        return redirect(url_for("index"))

    token = uuid.uuid4().hex
    pending_workbooks[token] = {
        "bytes": xlsx_bytes,
        "filename": file.filename,
        "added_at": added_at,
    }

    return render_template(
        "index.html",
        converted={
            "token": token,
            "filename": file.filename,
            "preview": workbook_preview(xlsx_bytes),
        },
        master_available=session.get("master_id") in master_workbooks,
    )


@app.route("/download/<token>", methods=["GET"])
def download_converted(token):
    workbook = pending_workbooks.get(token)
    if workbook is None:
        flash("That converted workbook is no longer available. Please convert the PDF again.")
        return redirect(url_for("index"))
    out_name = os.path.splitext(workbook["filename"])[0] + f"_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        io.BytesIO(workbook["bytes"]),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/add-to-master/<token>", methods=["POST"])
def add_to_master_route(token):
    workbook = pending_workbooks.get(token)
    if workbook is None:
        flash("That converted workbook is no longer available. Please convert the PDF again.")
        return redirect(url_for("index"))

    master_id = session.setdefault("master_id", uuid.uuid4().hex)
    master_workbooks[master_id] = add_to_master(
        master_workbooks.get(master_id),
        workbook["bytes"],
        workbook["filename"],
        workbook["added_at"],
    )
    flash(f"{workbook['filename']} was added to your master workbook.")
    return render_template(
        "index.html",
        converted={
            "token": token,
            "filename": workbook["filename"],
            "preview": workbook_preview(workbook["bytes"]),
        },
        master_available=True,
    )


@app.route("/download-master", methods=["GET"])
def download_master():
    master_id = session.get("master_id")
    workbook = master_workbooks.get(master_id)
    if workbook is None:
        flash("Your master workbook is empty. Add a converted report first.")
        return redirect(url_for("index"))
    return send_file(
        io.BytesIO(workbook),
        as_attachment=True,
        download_name=f"RSE_master_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def too_large(e):
    flash("That file is too large (max 20 MB).")
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
