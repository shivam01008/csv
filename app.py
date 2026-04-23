import os
import pandas as pd
import requests
from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from io import BytesIO

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def create_excel_from_csv(csv_path, output_path):
    df = pd.read_csv(csv_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    # Detect image columns
    image_columns = [col for col in df.columns if col.lower().startswith("image")]

    # All columns
    all_columns = list(df.columns)

    # Write headers exactly as CSV
    ws.append(all_columns)

    # Set column widths
    for i, col in enumerate(all_columns, start=1):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = 25

    row_num = 2

    for _, row in df.iterrows():
        col_index = 1

        for col_name in all_columns:
            value = row.get(col_name, "")

            # If it's NOT image column → write text
            if col_name not in image_columns:
                ws.cell(row=row_num, column=col_index, value=value)

            else:
                # Handle image
                if value and isinstance(value, str):
                    try:
                        response = requests.get(
                            value,
                            timeout=10,
                            headers={"User-Agent": "Mozilla/5.0"}
                        )

                        if response.status_code == 200:
                            img_data = BytesIO(response.content)
                            img = Image(img_data)

                            img.width = 80
                            img.height = 80

                            cell = f"{get_column_letter(col_index)}{row_num}"
                            ws.add_image(img, cell)

                            ws.row_dimensions[row_num].height = 65
                        else:
                            ws.cell(row=row_num, column=col_index, value="Invalid URL")

                    except Exception as e:
                        print("Error:", e)
                        ws.cell(row=row_num, column=col_index, value="Error")

            col_index += 1

        row_num += 1

    wb.save(output_path)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]

        if file.filename == "":
            return "No file selected"

        upload_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(upload_path)

        output_path = os.path.join(OUTPUT_FOLDER, "output.xlsx")

        create_excel_from_csv(upload_path, output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)