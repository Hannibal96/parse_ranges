import win32gui
import win32com.client
import win32clipboard
import pyautogui
import pickle
import time
import os.path


CLEAR_RANGE_X = 250
CLEAR_RANGE_Y = 380

EVALUATE_X = 1340
EVALUATE_Y = 385

RANGE_POSITIONS = ['MP2', 'MP3', 'CO', 'BU', 'SB', 'BB']


def parse_data(data):
    splited_data = str(data).split()
    MP2_equity = MP3_equity = CO_equity = DE_equity = SB_equity = BB_equity = 0
    for idx, word in enumerate(splited_data):
        if word == "MP2":
            MP2_equity = float(splited_data[idx+1][:-1])
        if word == "MP3":
            MP3_equity = float(splited_data[idx + 1][:-1])
        if word == "CO":
            CO_equity = float(splited_data[idx+1][:-1])
        if word == "BU":
            DE_equity = float(splited_data[idx + 1][:-1])
        if word == "SB":
            SB_equity = float(splited_data[idx + 1][:-1])
        if word == "BB":
            BB_equity = float(splited_data[idx + 1][:-1])

    return MP2_equity, MP3_equity, CO_equity, DE_equity, SB_equity, BB_equity


def find_equity_lab_window(window_text):
    def tables_collector(hwnd, tables_list, sub_string=window_text):
        if sub_string in win32gui.GetWindowText(hwnd):
            tables_list.append(hwnd)

    aof_tables_list = []
    win32gui.EnumWindows(tables_collector, aof_tables_list)

    return aof_tables_list[0]


def front_ground_window(hwnd):
    win32gui.SetForegroundWindow(hwnd)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')


def click_position(window, position):
    top_x, top_y, low_x, low_y = win32gui.GetWindowRect(window)
    if position == 'MP2':
        pyautogui.click(top_x + 88, top_y + 145)
    elif position == 'MP3':
        pyautogui.click(top_x + 88, top_y + 171)
    elif position == 'CO':
        pyautogui.click(top_x + 88, top_y + 197)
    elif position == 'BU':
        pyautogui.click(top_x + 88, top_y + 223)
    elif position == 'SB':
        pyautogui.click(top_x + 88, top_y + 249)
    elif position == 'BB':
        pyautogui.click(top_x + 88, top_y + 275)
    else:
        print("-E- impossible position for range insert")
        exit()


def insert_range(window, position, range, type):
    click_position(window, position)
    time.sleep(1)

    range_window = find_equity_lab_window(position)
    range_top_x, range_top_y, range_low_x, range_low_y = win32gui.GetWindowRect(range_window)

    if type == 'classic':
        pyautogui.doubleClick(range_top_x + 145, range_top_y + 500)
        for c in str(range):
            pyautogui.press(c)
            time.sleep(0.2)

    elif type == 'adjusted':
        if range == 0:
            pyautogui.doubleClick(range_top_x + 145, range_top_y + 500)
            pyautogui.press('0')
        else:
            pyautogui.click(range_top_x + 500, range_top_y + 110 + 18 * int(range / 5))
    else:
        print("-E- wrong type of ranges")
        exit()
    time.sleep(1)

    pyautogui.click(range_top_x + 515, range_top_y + 630)


def insert_hand(window, position, card_a, card_b, suited):
    click_position(window, position)
    time.sleep(1)
    range_window = find_equity_lab_window(position)
    range_top_x, range_top_y, range_low_x, range_low_y = win32gui.GetWindowRect(range_window)

    cor_x = 20
    cor_y = 65
    block_dim = 30.77
    big = max(card_a, card_b)
    small = min(card_a, card_b)

    if suited:
        cor_x += block_dim * (14 - small)
        cor_y += block_dim * (14 - big)

    else:
        cor_x += block_dim * (14 - big)
        cor_y += block_dim * (14 - small)

    pyautogui.click(range_top_x + cor_x, range_top_y + cor_y)
    time.sleep(0.2)
    pyautogui.click(range_top_x + 515, range_top_y + 630)


def clear_ranges(window):
    top_x, top_y, low_x, low_y = win32gui.GetWindowRect(window)
    pyautogui.click(top_x + CLEAR_RANGE_X, top_y + CLEAR_RANGE_Y)


def evaluate(window, calc_time):
    top_x, top_y, low_x, low_y = win32gui.GetWindowRect(window)
    pyautogui.click(top_x + EVALUATE_X, top_y + EVALUATE_Y)
    time.sleep(calc_time)

    evaluate_window = find_equity_lab_window("Eval")
    eval_top_x, eval_top_y, eval_low_x, eval_low_y = win32gui.GetWindowRect(evaluate_window)
    pyautogui.click(eval_top_x + 135, eval_top_y + 95)


def copy_values(window):
    top_x, top_y, low_x, low_y = win32gui.GetWindowRect(window)
    pyautogui.click(top_x + 60, top_y + 702)
    time.sleep(1)

    win32clipboard.OpenClipboard()
    MP2_equity, MP3_equity, CO_equity, DE_equity, SB_equity, BB_equity = parse_data(win32clipboard.GetClipboardData())
    #print(MP2_equity, MP3_equity, CO_equity, DE_equity, SB_equity, BB_equity)
    win32clipboard.CloseClipboard()
    return MP2_equity, MP3_equity, CO_equity, DE_equity, SB_equity, BB_equity


def calc_equity_tables(window):
    pass


def order_hands(window, vs_range=50, vs_players=1):
    path = "./hands_order_"+str(vs_players)+"p_"+str(vs_range)+"r.pickle"
    if os.path.exists(path):
        hands_order_2p = pickle.load(open(path, "rb"))
    else:
        hands_order_2p = {}
    try:
        for card_a in range(2, 15):
            for card_b in range(2, 15):
                for suit in [False, True]:
                    if card_a == card_b and suit:
                        continue
                    if (card_a, card_b, suit) in hands_order_2p.keys():
                        continue
                    for i in range(vs_players):
                        insert_range(window=window, position=RANGE_POSITIONS[len(RANGE_POSITIONS)-2-i], range=vs_range, type='classic')
                    insert_hand(window=window, position='BB', card_a=card_a, card_b=card_b, suited=suit)
                    evaluate(window=window, calc_time=5)
                    print(card_a, card_b, suit, end=': ')
                    hands_order_2p[(card_a, card_b, suit)] = copy_values(window=window)[5]
                    print(hands_order_2p[(card_a, card_b, suit)])
                    clear_ranges(window=window)
    except Exception as e:
        print("-I- caught exception exiting gracefully...")
        print(e)
        pickle.dump(hands_order_2p, open(path, "wb"))


def print_pickle_dict(dic):
    for key, val in dic.items():
        print(key, ":", val)


def txt_pickle_dict(dic, name):
    file = open(name+".txt", "w+")
    for key, val in dic.items():
        file.write(str(key)+": "+str(val)+"\n")
    file.close()


equilab_window = find_equity_lab_window("Equilab")
front_ground_window(equilab_window)
clear_ranges(equilab_window)


order_hands(window=equilab_window, vs_range=30, vs_players=2)


