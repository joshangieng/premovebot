import cv2 as cv
import pyautogui as pg
import numpy as np
from PIL import ImageGrab
import time
import win32ui, win32gui, win32con, win32api
from stockfish import Stockfish
from stockfish import models
import random
import sys
from pathlib import Path

def get_screenshot(x1, y1, w, h, windowname = None):
    '''
    takes a screenshot of the screen
    '''
    # get the window image data
    if not windowname:
        hwnd = None
    else:
        hwnd = win32gui.FindWindow(None, windowname)
    
    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    dataBitMap = win32ui.CreateBitmap()
    dataBitMap.CreateCompatibleBitmap(dcObj, w, h)
    cDC.SelectObject(dataBitMap)
    cDC.BitBlt((0,0), (w, h), dcObj, (x1, y1), win32con.SRCCOPY)

    # convert the raw data into a format opencv can read
    #dataBitMap.SaveBitmapFile(cDC, 'debug.bmp')
    signedIntsArray = dataBitMap.GetBitmapBits(True)
    img = np.frombuffer(signedIntsArray, dtype='uint8')
    img.shape = (h, w, 4)

    # free resources
    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(dataBitMap.GetHandle())

    # drop the alpha channel, or cv.matchTemplate() will throw an error like:
    #   error: (-215:Assertion failed) (depth == CV_8U || depth == CV_32F) && type == _templ.type() 
    #   && _img.dims() <= 2 in function 'cv::matchTemplate'
    img = cv.cvtColor(img, cv.COLOR_RGBA2RGB)

    return img

def find_pieces(piece_path, white: bool):
    '''
    returns 'board coords' like a1, a4 in integer format: ie a1 would be (0,0), a4 would be (0,3)
    '''
    piece = cv.imread(piece_path, cv.COLOR_RGBA2RGB)
    res = cv.matchTemplate(piece, screenshot, cv.TM_CCOEFF_NORMED)
    threshold = 0.9
    locations = np.where(res >= threshold)
    locations = list(zip(*locations[::-1]))
    
    rectangles = []
    for loc in locations:
        rect = [int(loc[0]), int(loc[1]), piece.shape[1], piece.shape[0]]
        # Add every box to the list twice in order to retain single (non-overlapping) boxes
        rectangles.append(rect)
        rectangles.append(rect)
        
    rectangles, _ = cv.groupRectangles(rectangles, groupThreshold=1, eps=0.5)

    points = []
    if len(rectangles):
        if white:
            for (x, y, w, h) in rectangles:
                center_x = x + int(w/2)
                center_y = y + int(h/2)

                points.append((center_x//75, 8 - center_y//75))
        else:
            for (x, y, w, h) in rectangles:
                center_x = x + int(w/2)
                center_y = y + int(h/2)

                points.append((7 - center_x//75, center_y//75 - 1))
    
    return points

def make_fen():
    # '''
    # makes fen code and returns it
    # '''
    for path in piece_paths:
        piece_locations.append(find_pieces(path, is_white))

    # Create an 8x8 board filled with empty strings
    board = np.full((8, 8), '', dtype='<U1')

    # Place pieces on the board
    for piece, locs in zip(piece_order, piece_locations):
        for x, y in locs:
            board[7 - y, x] = piece  # Flip y-axis for FEN rank order

    # Build FEN rows
    fen_rows = []
    for row in board:
        empty_count = 0
        fen_row = []
        for cell in row:
            if cell == '':
                empty_count += 1
            else:
                if empty_count > 0:
                    fen_row.append(str(empty_count))
                    empty_count = 0
                fen_row.append(cell)
        if empty_count > 0:
            fen_row.append(str(empty_count))
        fen_rows.append(''.join(fen_row))

    return '/'.join(fen_rows)

def make_move_on_screen(move):
    #get locations to click to make the move
    char1 = move[0]
    num1 = int(move[1])
    char2 = move[2]
    num2 = int(move[3])

    if is_white:
        init_x = x1 + square_size//2 + square_size*(ord(char1) - ord('a'))
        init_y = y2_board - square_size//2 - square_size*(num1 - 1)
        future_x = x1 + square_size//2 + square_size*(ord(char2) - ord('a'))
        future_y = y2_board - square_size//2 - square_size*(num2 - 1)
    else:
        init_x = x2 - square_size//2 - square_size*(ord(char1) - ord('a'))
        init_y = y1_board + square_size//2 + square_size*(num1 - 1)
        future_x = x2 - square_size//2 - square_size*(ord(char2) - ord('a'))
        future_y = y1_board + square_size//2 + square_size*(num2 - 1)

    #click on init loc
    pg.moveTo(init_x, init_y)
    pg.mouseDown(button="left")

    #drag to future loc
    pg.moveTo(future_x, future_y)
    pg.mouseUp(button="left")

def premove(move_made):
    no_premoves = 20
    rand = 1 + random.randrange(no_premoves)

    stockfish.make_moves_from_current_position([move_made])

    for _ in range(rand):
        enemy_move = stockfish.get_best_move()
        stockfish.make_moves_from_current_position([enemy_move])

        premove = stockfish.get_best_move()
        make_move_on_screen(premove)
        stockfish.make_moves_from_current_position([premove])

########################################################################## MAIN SCRIPT
# foreground = win32gui.GetForegroundWindow()
# win32gui.ShowWindow(foreground, win32con.SW_MAXIMIZE)

pg.PAUSE = 0

print("WARNING: PLEASE ENSURE THAT: SCALE IS SET TO 100 PERCENT AND RESOLUTION IS 1600x900\n")
print("ALSO, MAKE SURE YOU ARE USING THE CORRECT BOARD (looks brown and fuzzy) AND PIECES (looks like default but no shading)")
print("IF BUG OCCURS, TAB OUT OF CHESS.COM AND THEN BACK IN")

if getattr(sys, 'frozen', False):  # Running as exe
    project_dir = Path(sys._MEIPASS)
else:  # Running as script
    project_dir = Path(__file__).parent

speed = None

depth = 13

print("\nWaiting for game start...")

stockfish = Stockfish(path = project_dir / "stockfish" / "stockfish-windows-x86-64-avx2.exe", parameters={"Minimum Thinking Time": 0, "Threads": 1, "Slow Mover": 0, "Hash": 64, "Ponder": "true"})
stockfish.set_depth(depth)

gameend = cv.imread(project_dir / "assets" / "gameend.png", cv.COLOR_RGBA2RGB)
abort = cv.imread(project_dir / "assets" / "abort.png", cv.COLOR_RGBA2RGB)

piece_paths = [
    project_dir / "assets" / "blackbishop.png",
    project_dir / "assets" / "blackking.png",
    project_dir / "assets" / "blackknight.png",
    project_dir / "assets" / "blackpawn.png",
    project_dir / "assets" / "blackqueen.png",
    project_dir / "assets" / "blackrook.png",
    
    project_dir / "assets" / "whitebishop.png",
    project_dir / "assets" /"whiteking.png",
    project_dir / "assets" / "whiteknight.png",
    project_dir / "assets" / "whitepawn.png",
    project_dir / "assets" / "whitequeen.png",
    project_dir / "assets" / "whiterook.png"
    ]

piece_order = ['b','k','n','p','q','r','B','K','N','P','Q','R']

#works on 1920x1080 - to change for different resolutions find the pixel of top left of chessboard
x1 = 231
y1 = 104
x2 = 828
y2 = 832
w = x2-x1
h = y2-y1
t = 3
square_size = 75
y1_board = 153
y2_board = 752

red = (np.uint8(36), np.uint8(31), np.uint8(173))
toggle = False
my_timer_colour = None
is_white = None
side_chosen = False

while True:
    screenshot = get_screenshot(x1, y1, w, h)

    # cv.imshow("screen", screenshot)
    # if cv.waitKey(1) == ord("q"):
    #     cv.destroyAllWindows()
    #     break
    
    gameendcomp = cv.matchTemplate(gameend, screenshot, cv.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv.minMaxLoc(gameendcomp)
    if max_val > 0.9:
        print("Game ended, just look for timer color")
        toggle = False
        continue

    abortcomp = cv.matchTemplate(abort, screenshot, cv.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv.minMaxLoc(abortcomp)
    if max_val > 0.9:
        print("Game ended, just look for timer color")
        toggle = False
        continue

    if tuple(screenshot[h-t, w-t]) == (np.uint8(255), np.uint8(255), np.uint8(255)):
        if side_chosen == False:
            my_timer_colour = (np.uint8(255), np.uint8(255), np.uint8(255))
            my_off_timer_colour = (np.uint8(149), np.uint8(151), np.uint8(152))
            is_white = True
            toggle = False
            print("Playing white")
            side_chosen = True

    elif tuple(screenshot[h-t, w-t]) == (np.uint8(33), np.uint8(36), np.uint8(38)):
        if side_chosen == False:
            my_timer_colour = (np.uint8(33),np.uint8(36),np.uint8(38))
            my_off_timer_colour = (np.uint8(37), np.uint8(40), np.uint8(42))
            is_white = False
            toggle = False
            print("Playing black")
            side_chosen = True

    else:
        side_chosen = False

    if my_timer_colour is None:
        continue

    ################################ if in game 

    piece_locations = []

    if tuple(screenshot[h-t, w-t]) == my_off_timer_colour:
        if toggle:
            print('Their move, waiting...\n')
            print(tuple(screenshot[h-t, w-t]))
            toggle = False
        
    elif tuple(screenshot[h-t, w-t]) == my_timer_colour or tuple(screenshot[h-t, w-t]) == red:
        if not toggle:
            print('My move!')
            print(tuple(screenshot[h-t, w-t]))

        #1. RECORD BOARD POSITION
            fen = make_fen()
            
            #edit fen for turn
            if is_white:
                fen += ' w - - 0 1'
                
            else:
                fen += ' b - - 0 1'

            stockfish.set_fen_position(fen)

        #2. MAKE MY MOVE
            move_to_make = stockfish.get_best_move()
            
            #set_pause(turn, speed)
            try:
                make_move_on_screen(move_to_make)
                #premove everything
                premove(move_to_make)
            except:
                print("bug handled")
                continue

            toggle = True