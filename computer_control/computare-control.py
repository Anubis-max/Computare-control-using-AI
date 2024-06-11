import cv2
import mediapipe as mp
import pyautogui
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import os

capture_hands = mp.solutions.hands.Hands(max_num_hands=1)
screen_width, screen_height = pyautogui.size()
camera = cv2.VideoCapture(0) 

def perform_action(x_list, y_list):
    match True:
        case _ if ((x_list[8]-x_list[4])**2 + (y_list[8]-y_list[4])**2)**(0.5)//4 < 5:
            pyautogui.leftClick()
        case _ if ((x_list[7]-x_list[4])**2 + (y_list[7]-y_list[4])**2)**(0.5)//4 < 5:
            pyautogui.mouseDown(button='left')
        case _ if ((x_list[6]-x_list[4])**2 + (y_list[6]-y_list[4])**2)**(0.5)//4 < 5:
            pyautogui.mouseUp(button='left')
        case _ if ((x_list[12]-x_list[4])**2 + (y_list[12]-y_list[4])**2)**(0.5)//4 < 5:
            pyautogui.scroll(30)
        case _ if ((x_list[10]-x_list[4])**2 + (y_list[10]-y_list[4])**2)**(0.5)//4 < 5:
            pyautogui.scroll(-30)
        case _ if ((x_list[16]-x_list[4])**2 + (y_list[16]-y_list[4])**2)**(0.5)//4 < 5:
            pyautogui.rightClick()
        case _ if ((x_list[20]-x_list[4])**2 + (y_list[20]-y_list[4])**2)**(0.5)//4 < 5:
            increase_volume(0.1)
        case _ if ((x_list[18]-x_list[4])**2 + (y_list[18]-y_list[4])**2)**(0.5)//4 < 5:
            decrease_volume(0.1)
        case _ if ((x_list[20]-x_list[17])**2 + (y_list[20]-y_list[17])**2)**(0.5)//4 < 5 and ((x_list[16]-x_list[13])**2 + (y_list[16]-y_list[13])**2)**(0.5)//4 < 5 and ((x_list[20]-x_list[4])**2 + (y_list[20]-y_list[4])**2)**(0.5)//4 < 10 and ((x_list[11]-x_list[7])**2 + (y_list[11]-y_list[7])**2)**(0.5)//4 > 10:
            os.system('shutdown -s')

def increase_volume(increase_by=0.1):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = min(1.0, current_volume + increase_by)
    volume.SetMasterVolumeLevelScalar(new_volume, None)

def decrease_volume(decrease_by=0.1):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = max(0.0, current_volume - decrease_by)
    volume.SetMasterVolumeLevelScalar(new_volume, None)

while(True):
    _, image = camera.read()
    image_height, image_width, _ = image.shape
    image = cv2.flip(image, 1)
    rgb_image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    output_hands = capture_hands.process(rgb_image)
    all_hands = output_hands.multi_hand_landmarks
    if all_hands:
        hand = all_hands[0]
        one_hand_landmarks = hand.landmark

        x_list = [0] * 21
        y_list = [0] * 21

        for id, lm in enumerate(one_hand_landmarks):
            x = int(lm.x * image_width)
            y = int(lm.y * image_height)

            if id == 0:
                x_list[0] = int(screen_width / image_width * x)
                y_list[0] = int(screen_width / image_width * y)
                pyautogui.moveTo(x_list[0], y_list[0])
            else:
                x_list[id] = x
                y_list[id] = y

        perform_action(x_list, y_list)

    cv2.imshow("Controlling Computer", image)
    key = cv2.waitKey(1)
    if key == 27:
        break

camera.release()
cv2.destroyAllWindows()