import face_recognition  # type: ignore
import cv2
import numpy as np
import os
from datetime import date, datetime
import xlrd  # type: ignore
from xlutils.copy import copy as xl_copy  # type: ignore

def load_face_encodings(image_paths):
    encodings = []
    for image_path in image_paths:
        image = face_recognition.load_image_file(image_path)
        encodings.append(face_recognition.face_encodings(image)[0])
    return encodings

def find_next_empty_row(sheet):
    row = 1
    while True:
        try:
            if sheet.cell_value(row, 0) == '':
                break
        except IndexError:
            break
        row += 1
    return row

def switch_case(value):
    if value == 1:
        return "English"
    elif value == 2:
        return "Physics"
    elif value == 3:
        return "Chemistry"
    else:
        return None

def main():
    current_folder = os.getcwd()
    image_paths = [os.path.join(current_folder, f'{name}.png') for name in ['sahil', 'ameen', 'affan']]
    known_face_encodings = load_face_encodings(image_paths)
    known_face_names = ["Sahil", "Ameen", "Affan"]
    known_roll_number = ["161023749019", "161023749031", "161023733135"]

    face_locations = []
    face_encodings = []
    face_names = []
    process_this_frame = True

    rb = xlrd.open_workbook('attendence_excel.xls', formatting_info=True)
    wb = xl_copy(rb)

    print("Choose the class by number:\n 1.English \n 2.Physics \n 3.Chemistry")
    value = int(input("Enter the class number: "))
    sheet_name = switch_case(value)
    
    if sheet_name is None:
        print("Invalid class number. Exiting.")
        return

    try:
        sheet = rb.sheet_by_name(sheet_name)
        writable_sheet = wb.get_sheet(rb.sheet_names().index(sheet_name))
    except xlrd.biffh.XLRDError:
        writable_sheet = wb.add_sheet(sheet_name)
        writable_sheet.write(0, 0, 'Name')
        writable_sheet.write(0, 1, str(date.today()))
        writable_sheet.write(0, 2, "Time")
        writable_sheet.write(0, 3, "Roll Number")
        row = 1
    else:
        row = find_next_empty_row(sheet)

    already_attendence_taken = set()

    video_capture = cv2.VideoCapture(0)

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
                    c_time = datetime.now()
                    f_time = c_time.strftime('%I:%M:%S %p')
                    writable_sheet.write(row, 0, name)
                    writable_sheet.write(row, 1, "Present")
                    writable_sheet.write(row, 2, f_time)
                    writable_sheet.write(row, 3, roln)
                    print(f"Attendance taken for {name}")
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
        if k % 256 == 27:  # ESC key to break
            print("Data saved")
            break

    video_capture.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
