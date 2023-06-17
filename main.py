from flask import Flask, send_file
from flask_cors import CORS
from openpyxl import Workbook
import os
from test_report import InspectionExcel

app = Flask(__name__)
cors = CORS(app)


@app.route("/")
def home():
    return "Hey there! Its working!"


@app.route("/demo-product-list")
def demo_product_list():
    return send_file(os.path.join("demo-product-list.xlsx"))


@app.route("/<noteId>")
def test(noteId):
    InspectionExcel(noteId).generate_excel_files()
    return send_file(os.path.join("newR.xlsx"))


if __name__ == "__main__":
    app.run(debug=True)
