from flask import Flask, render_template, request, redirect, url_for
import cv2
import face_recognition
import numpy as np
import os
from datetime import date, datetime
import xlrd
from xlutils.copy import copy as xl_copy

app = Flask(__name__)

# Load known faces
def load_known_faces():
    global known_face_encodings, known_face_names, known_roll_number
    CurrentFolder = os.getcwd()
    image_paths = [CurrentFolder + '\\sahil.png', CurrentFolder + '\\ameen.png', CurrentFolder + '\\affan.png']
    names = ["Sahil", "Ameen", "Affan"]
    roll_numbers = ["161023749019", "161023749031", "161023733135"]

    for image_path, name, roll_number in zip(image_paths, names, roll_numbers):
        image = face_recognition.load_image_file(image_path)
        face_encoding = face_recognition.face_encodings(image)[0]
        known_face_encodings.append(face_encoding)
        known_face_names.append(name)
        known_roll_number.append(roll_number)

known_face_encodings = []
known_face_names = []
known_roll_number = []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        class_code = request.form['class_code']
        return redirect(url_for('attendance', class_code=class_code))
    return render_template('login.html')

@app.route('/register')
def register():
    return render_template('register.html')

@app.route('/attendance/<class_code>')
def attendance(class_code):
    return render_template('attendance.html', class_code=class_code)

@app.route('/capture', methods=['POST'])
def capture():
    action = request.form['action']
    if action == 'register':
        capture_image()
    elif action == 'attendance':
        class_code = request.form['class_code']
        take_attendance(class_code)
    return redirect(url_for('index'))

def capture_image():
    video_capture = cv2.VideoCapture(0)
    while True:
        ret, frame = video_capture.read()
        cv2.imshow('Video', frame)
        k = cv2.waitKey(1)
        if k % 256 == 32:  # Space bar to capture image
            img_name = "captured_image.png"
            cv2.imwrite(img_name, frame)
            print(f"Image {img_name} saved!")
        elif k % 256 == 27:  # ESC to exit
            print("Closing camera")
            break
    video_capture.release()
    cv2.destroyAllWindows()

def take_attendance(class_code):
    c_time = datetime.now()
    f_time = c_time.strftime('%I:%M:%S %p')
    CurrentFolder = os.getcwd()
    image_paths = [CurrentFolder + '\\sahil.png', CurrentFolder + '\\ameen.png', CurrentFolder + '\\affan.png']
    names = ["Sahil", "Ameen", "Affan"]
    roll_numbers = ["161023749019", "161023749031", "161023733135"]

    known_face_encodings = []
    known_face_names = []
    known_roll_number = []

    for image_path, name, roll_number in zip(image_paths, names, roll_numbers):
        image = face_recognition.load_image_file(image_path)
        face_encoding = face_recognition.face_encodings(image)[0]
        known_face_encodings.append(face_encoding)
        known_face_names.append(name)
        known_roll_number.append(roll_number)

    video_capture = cv2.VideoCapture(0)
    face_locations = []
    face_encodings = []
    face_names = []
    process_this_frame = True
    rb = xlrd.open_workbook('attendence_excel.xls', formatting_info=True)
    wb = xl_copy(rb)
    sheet = wb.add_sheet(switch_case(class_code))
    sheet.write(0, 0, 'Name')
    sheet.write(0, 1, str(date.today()))
    sheet.write(0, 2, "Time")
    sheet.write(0, 3, "Roll Number")
    row = 1
    already_attendence_taken = set()

    while True:
        ret, frame = video_capture.read()
        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
        rgb_small_frame = small_frame[:, :, ::-1]
        if process_this_frame:
            face_locations = face_recognition.face_locations(rgb_small_frame)
            face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
            face_names = []
            for face_encoding in face_encodings:
                matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
                name = "Unknown"
                face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
                best_match_index = np.argmin(face_distances)
                if matches[best_match_index]:
                    name = known_face_names[best_match_index]
                    roln = known_roll_number[best_match_index]
                face_names.append(name)
                if name not in already_attendence_taken and name != "Unknown":
                    sheet.write(row, 0, name)
                    sheet.write(row, 1, "Present")
                    print(f"Attendance taken for {name}")
                    sheet.write(row, 2, f_time)
                    sheet.write(row, 3, roln)
                    row += 1
                    wb.save('attendence_excel.xls')
                    already_attendence_taken.add(name)
                else:
                    print("Next student")
        process_this_frame = not process_this_frame
        for (top, right, bottom, left), name in zip(face_locations, face_names):
            top *= 4
            right *= 4
            bottom *= 4
            left *= 4
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
            cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
            font = cv2.FONT_HERSHEY_DUPLEX
            cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)
        cv2.imshow('Video', frame)
        k = cv2.waitKey(1)
        if k % 256 == 27:  # ESC to exit
            print("Data saved")
            break
    video_capture.release()
    cv2.destroyAllWindows()

def switch_case(value):
    if value == '1':
        return "English"
    elif value == '2':
        return "Physics"
    elif value == '3':
        return "Chemistry"
    else:
        return "Invalid input"

if __name__ == '__main__':
    load_known_faces()
    app.run(debug=True)
