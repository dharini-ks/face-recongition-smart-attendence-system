from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(text):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

# Load data
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('C:/Users/dharini ks/Desktop/python/data/haarcascade_frontalface_default.xml')

# Load the labels and faces data
with open('C:/Users/dharini ks/Desktop/python/data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('C:/Users/dharini ks/Desktop/python/data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Loaded LABELS:', LABELS)
print('Shape of Faces matrix:', FACES.shape)

# Train the KNN model
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground = cv2.imread("background.png")
COL_NAMES = ['NAME', 'TIME']
THRESHOLD = 0.6  # Set a threshold distance for recognizing faces

recognized_faces_set = set()  # Set to keep track of recognized faces
unknown_faces_set = set()     # Set to keep track of faces labeled as "Unknown"

while True:
    ret, frame = video.read()
    if not ret:
        print("Failed to capture image from webcam.")
        continue

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)

        # Get the predicted label and distance to the nearest neighbors
        distances, indices = knn.kneighbors(resized_img)
        min_distance = distances[0][0]
        predicted_label = knn.predict(resized_img)[0]

        print(f"Predicted label: {predicted_label}, Distance: {min_distance}")

        # Check if the face has already been recognized
        if predicted_label not in recognized_faces_set and predicted_label in LABELS:
            recognized_faces_set.add(predicted_label)
            ts = time.time()
            timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
            cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
            cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
            cv2.putText(frame, str(predicted_label), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
            
            attendance = [str(predicted_label), str(timestamp)]
            # Save attendance
            if os.path.isfile("Attendance1.csv"):
                with open("Attendance1.csv", "a", newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(attendance)
            else:
                with open("Attendance1.csv", "w", newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    writer.writerow(attendance)
        else:
            # Only print "Unknown" and speak it if it hasn't been printed before for this face
            if min_distance > THRESHOLD and predicted_label not in unknown_faces_set:
                unknown_faces_set.add(predicted_label)
                print("Unknown face detected.")
                cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 2)
                cv2.putText(frame, "Unknown", (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 255), 2)
                speak("Unknown")

    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)

    k = cv2.waitKey(1)
    if k == ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
    if k == ord('q'):
        breakq

video.release()
cv2.destroyAllWindows()
