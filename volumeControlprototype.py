import cv2
import math
import pyautogui
import ctypes
import mediapipe as mp
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

SNDVOLSSOF = 0x1
SNDVOLSSOS = 0x2

def set_system_volume(volume_level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(volume_level, None)

def get_system_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar() * 100  
    return current_volume

cap = cv2.VideoCapture(0)

mp_hands = mp.solutions.hands
hands = mp_hands.Hands(static_image_mode = False, max_num_hands = 1,
                       min_detection_confidence = 0.5, min_tracking_confidence = 0.5)

mp_drawing = mp.solutions.drawing_utils

CONSTANT_RATIO = 6
UPPER_RATIO = 25
LOWER_RATIO = 5

while True:
    ret, frame = cap.read()
    if not ret:
        break

    image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    results = hands.process(image_rgb)
    volume = get_system_volume()

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

            index_finger_x = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP].x
            index_finger_y = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP].y
            thumb_x = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP].x
            thumb_y = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP].y
            thumb_ip_x = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_IP].x
            thumb_ip_y = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_IP].y

            it_distance = math.sqrt((index_finger_x - thumb_x)**2 + (index_finger_y - thumb_y)**2)
            joint_distance = math.sqrt((thumb_ip_x - thumb_x)**2 + (thumb_ip_y - thumb_y)**2)
            it_distance_adjusted = CONSTANT_RATIO * it_distance/joint_distance

            if int(it_distance_adjusted) > UPPER_RATIO:
                it_distance_adjusted = 100
            elif int(it_distance_adjusted) < LOWER_RATIO:
                it_distance_adjusted = 0
            else:
                it_distance_adjusted = (int(it_distance_adjusted) - 5) * 5 

            set_system_volume(it_distance_adjusted/100)
           
            print("Ratio: " + str(it_distance_adjusted) + "\n")

            '''
            if index_finger_y < thumb_y:
                gesture = 'pointing up'
            elif index_finger_y > thumb_y:
                gesture = 'pointing down'
            else:
                gesture = 'other'

            if gesture == 'pointing up':
                pyautogui.press('volumeup')
            elif gesture == 'pointing down':
                pyautogui.press('volumedown')
            '''
            
            cv2.imshow('Hand Gesture', frame)

            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
        else:
            continue

cap.release()
cv2.destroyAllWindows()
