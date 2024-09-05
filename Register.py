from cv2 import VideoCapture, imshow, imwrite
import cv2
cam = VideoCapture(0)
# reading the input using the camera

inp = input('Enter person name: ')

# If image will detected without any error,
# show result
while True: 
        result,image = cam.read()
        imshow(inp, image)
        k = cv2.waitKey(1)
        if k%256 == 27:
                print("Closing the window")
                break   
        elif k%256 == 32:
                imwrite(inp+".png", image)
                print("image taken")
