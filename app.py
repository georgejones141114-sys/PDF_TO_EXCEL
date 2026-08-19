import io
import os
from datetime import datetime

from flask import Flask, render_template, request, send_file, flash, redirect, url_for

from converter import convert, ReportParseError

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-secret-change-me")

MAX_CONTENT_LENGTH = 20 * 1024 * 1024  # 20 MB
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


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
        xlsx_bytes = convert(pdf_bytes)
    except ReportParseError as e:
        flash(str(e))
        return redirect(url_for("index"))
    except Exception:
        flash("Something went wrong while converting this PDF. Please double-check the file and try again.")
        return redirect(url_for("index"))

    out_name = os.path.splitext(file.filename)[0] + f"_{datetime.now().strftime('%Y%m%d')}.xlsx"

    return send_file(
        io.BytesIO(xlsx_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def too_large(e):
    flash("That file is too large (max 20 MB).")
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
