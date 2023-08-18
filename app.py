from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
import face_recognition
import os
import pickle
import datetime

app = Flask(__name__)

# Global variables
directory_path = ""
selected_image_paths = []
subject = ""
class_info = ""
output_folder = ""

# Process images route
@app.route('/', methods=['GET', 'POST'])
def process_images():
    global directory_path, selected_image_paths, subject, class_info

    if request.method == 'POST':
        subject = request.form.get('subject')
        class_info = request.form.get('class_info')
        selected_image_paths = request.files.getlist('selected_images')

        try:
            # Load encodings
            with open("encodings.pickle", "rb") as f:
                data = pickle.load(f)
            known_face_encodings = data["encodings"]
            roll_numbers = data["roll_numbers"]

            # Create dictionaries to store unique recognized faces and unknown faces
            recognized_faces = {}
            unknown_faces = {}

            for selected_image in selected_image_paths:
                image = face_recognition.load_image_file(selected_image)
                face_locations = face_recognition.face_locations(image)
                face_encodings = face_recognition.face_encodings(image, face_locations)

                for face_encoding in face_encodings:
                    roll_number_match = "Unknown"
                    student_name = "---"
                    for i, known_encoding_list in enumerate(known_face_encodings):
                        matches = face_recognition.compare_faces(known_encoding_list, face_encoding)
                        if any(matches):
                            roll_number_match = roll_numbers[i]
                            folder_name_parts = roll_number_match.split('_')
                            if len(folder_name_parts) >= 2:
                                student_name = folder_name_parts[1]
                                roll_number_match = folder_name_parts[0]
                            break

                    if roll_number_match != "Unknown":
                        recognized_faces[roll_number_match] = student_name
                    else:
                        unknown_faces[face_encoding.tobytes()] = unknown_faces.get(face_encoding.tobytes(), 0) + 1

            # Sort and process recognized faces
            sorted_roll_numbers = sorted(recognized_faces.keys())

            # Create Excel sheet and save
            wb = Workbook()
            ws = wb.active
            ws.title = f"{subject}_{class_info}"
            ws.column_dimensions["A"].width = 15
            ws.append(["Student USN", "Student Name"])

            # Add data to Excel sheet
            for roll_number in sorted_roll_numbers:
                student_name = recognized_faces[roll_number]
                ws.append([roll_number, student_name])

            current_date = datetime.datetime.now().strftime("%d_%m_%Y")
            excel_file_path = os.path.join(app.root_path,
                                           f"attendance_sheet_{subject}_{class_info}_{current_date}.xlsx")
            wb.save(excel_file_path)

            return f"Recognition complete. Attendance saved. <a href=\"/download/?file={excel_file_path}\">Download Excel Sheet</a>"
        except Exception as e:
            result = "Recognition failed."
            print(str(e))
            return result

    else:
        return render_template('process_images.html')

# Download Excel sheet route
@app.route('/download/')
def download_excel():
    filename = request.args.get('file')
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
