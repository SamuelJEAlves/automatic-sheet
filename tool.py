import pyautogui as pag
from mouseinfo import mouseInfo

# mouseInfo()

pag.PAUSE = 2

pag.press('win')
pag.write('AutomaticSheet.xlsx')
pag.press('enter')
pag.click(x=57, y=247)

with open('data.txt', 'r') as file:
    for line in file:
        pag.PAUSE = 0.5
        pag.write(line)

pag.hotkey('ctrl', 'b')
pag.click('save.png')
