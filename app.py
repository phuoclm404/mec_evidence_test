from flask import (
    Flask,
    render_template,
    request,
    send_from_directory,
    redirect,
    url_for,
)
import os
from main import run

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["PROCESSED_FOLDER"] = "processed"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if request.method == "POST" and "file" in request.files:
        file = request.files["file"]
        if file.filename != "":
            # Lưu tên tệp
            filename = file.filename

            # Lưu tệp vào thư mục uploads
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(file_path)

            # Xử lý tệp ở đây (ví dụ: viết mã Python để xử lý tệp)
            run(filename)
            # Lưu tệp đã xử lý vào thư mục processed
            processed_file_path = os.path.join(app.config["PROCESSED_FOLDER"], filename)
            # Để đơn giản, chúng ta sẽ chỉ sao chép tệp ban đầu vào thư mục processed làm ví dụ
            os.rename(file_path, processed_file_path)
            filename = "evidence_" + filename
            # return f"File {filename} uploaded and processed successfully!"
            return redirect(url_for("download", filename=filename))

    return "No file selected."


@app.route("/download/<path:filename>", methods=["GET"])
def download(filename):
    # Đường dẫn tới thư mục chứa các tệp đã xử lý
    directory = "processed"

    # Trả về tệp đã xử lý để tải về cho người dùng
    return send_from_directory(directory, filename, as_attachment=True)


if __name__ == "__main__":
    if not os.path.exists(app.config["UPLOAD_FOLDER"]):
        os.makedirs(app.config["UPLOAD_FOLDER"])
    if not os.path.exists(app.config["PROCESSED_FOLDER"]):
        os.makedirs(app.config["PROCESSED_FOLDER"])
    try:
        app.run(debug=True)
    except Exception as e:
        print("Error: %s" % e)
