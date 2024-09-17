import cv2
import numpy as np
from mss import mss
from PIL import Image
import pygetwindow as gw
# global variables
farm_img = cv2.imread('assets/farm.png', cv2.IMREAD_UNCHANGED)
wheat_img = cv2.imread('assets/needle.png', cv2.IMREAD_UNCHANGED)
pineapple_img = cv2.imread('assets/pineapple.png', cv2.IMREAD_UNCHANGED)
flower_template_img = cv2.imread('assets/flower_template.jpg')
flower_template_crop_img = cv2.imread('assets/flower_template_crop.jpg')
# component
def show_root_img():
    cv2.imshow('Farm', farm_img)
    cv2.waitKey()
    cv2.destroyAllWindows()
def show_needle_img():
    cv2.imshow('Needle', wheat_img)
    cv2.waitKey()
    cv2.destroyAllWindows()
def show_pineapple_img():
    cv2.imshow('Pineapple', pineapple_img)
    cv2.waitKey()
    cv2.destroyAllWindows()
def show_flower_template_img():
    cv2.imshow('Flower_template', flower_template_img)
    cv2.waitKey()
    cv2.destroyAllWindows()
def capture_screen():
    bounding_box = {'top': 100, 'left': 0, 'width': 1920, 'height': 1080}
    sct = mss()

    while True:
        sct_image = sct.grab(bounding_box)
        cv2.imshow('screen', np.array(sct_image))

        if(cv2.waitKey(1) & 0xFF) == ord('q'):
            cv2.destroyAllWindows()
            break
def capture_specific_screen():
    app_title = 'Visual Studio Code'
    app_window = gw.getWindowsWithTitle(app_title)
    if len(app_window) == 1:
        app_window = app_window[0]

        app_rect = {'left': app_window.left, 'top': app_window.top, 'width': app_window.width, 'height': app_window.height}

        with mss() as sct:
            screenshot = sct.grab(app_rect)
            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
            img.save('screenshot_opencv_specific_app.png')
        
        image = cv2.imread('screenshot_opencv_specific_app.png')

        cv2.imshow(f'Capture Screen - {app_title}', image)
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        print(f"Screenshot of '{app_title}' captured and displayed")
    else:
        print(f"Application window '{app_title}' not found.")

def detect_object():
    # show image
    # show_root_img()
    # show_needle_img()
    # show_pineapple_img()
    img_search = flower_template_crop_img
    img_root = flower_template_img
    # get result match 
    result = cv2.matchTemplate(img_root, img_search, cv2.TM_CCOEFF_NORMED)
    # cv2.imshow('Result', result)
    # cv2.waitKey()
    # cv2.destroyAllWindows()

    # draw rectangle around match
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    w = img_search.shape[1]
    h = img_search.shape[0]

    # draw first match result
    # cv2.rectangle(farm_img, max_loc, (max_loc[0] + w, max_loc[1]+h), (0,255,255), 2)
    # show_root_img()

    # find all match result
    threshold = .60
    yloc, xloc  = np.where(result >= threshold)
    print(len(xloc))

    # draw all result
    # for (x,y) in zip(xloc, yloc):
    #     cv2.rectangle(farm_img, (x,y), (x+w, y+h), (0,255,255), 2)
    # show_root_img()
    
    # distinct duplicate result
    rectangles = []
    for (x,y) in zip(xloc, yloc):
        rectangles.append([int(x), int(y), int(w), int(h)])
        rectangles.append([int(x), int(y), int(w), int(h)])
    print(len(rectangles))

    rectangles, weights = cv2.groupRectangles(rectangles, 1, 0.2)
    print(len(rectangles))

    #draw final result
    for (x,y,w,h) in rectangles:
        cv2.rectangle(img_root, (x,y), (x+w, y+h), (0,255,255), 2)
    # show_root_img()
    show_flower_template_img()
if __name__ == "__main__":
    detect_object()
    # capture_screen()
    # capture_specific_screen()