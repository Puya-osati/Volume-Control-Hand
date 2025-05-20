import mediapipe as mp
import cv2
import math
import numpy as np
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = interface.QueryInterface(IAudioEndpointVolume)
# volume.GetMute()
# volume.GetMasterVolumeLevel()
# print(volume.GetVolumeRange())
# volume.SetMasterVolumeLevel(-65.0, None)


cap = cv2.VideoCapture(0)
mpHands = mp.solutions.hands
hands = mpHands.Hands()
mpDraw = mp.solutions.drawing_utils  # Assign drawing utils to a variable

while True:
    success, img = cap.read()
    if not success:
        print("Failed to capture image")
        break

    imRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    result = hands.process(imRGB)

    if result.multi_hand_landmarks:

        hand = result.multi_hand_landmarks[0]

        mpDraw.draw_landmarks(img, hand, mpHands.HAND_CONNECTIONS)

        lmList = []

        for id, lm in enumerate(hand.landmark):

            h, w, c = img.shape

            cx, cy = int(lm.x * w), int(lm.y * h)

            lmList.append([id, cx, cy])

        if len(lmList) > 8:
            x1, y1 = lmList[4][1], lmList[4][2]
            x2, y2 = lmList[8][1], lmList[8][2]
            cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

            cv2.circle(img, (x1, y1), 10, (255, 255, 0), cv2.FILLED)
            cv2.circle(img, (x2, y2), 10, (255, 255, 0), cv2.FILLED)
            cv2.circle(img, (cx, cy), 10, (255, 255, 0), cv2.FILLED)
            cv2.line(img, (x1, y1), (x2, y2), (255, 255, 0), 3)

            lenght = int(math.hypot(x2 - x1, y2 - y1))

            handRange =  [100, 300]
            vol = int(np.interp(lenght, handRange, [-25, 0]))
            volume.SetMasterVolumeLevel(vol, None)
            #print(lenght)

    cv2.imshow("Image", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
