# Web Game Playing Bot
# Plays MiniClip Ores
# http://www.miniclip.com/games/ores/en/
#
# Game by Hard Circle
# Bot by Adrian Dale 19/02/2013
#
# Instructions
# Run c:\python32\python OresBot.py
# Start Ores game in browser and wait for Start button to be visible.
# Click on pygame window, then press Space to start.
import pygame, sys
from pygame.locals import *

from numpy import *
from random import randrange

import win32gui
import win32ui
import win32con
import win32api
import os

pygame.init()
fpsClock = pygame.time.Clock()
DISPLAYSURF = pygame.display.set_mode((640, 480))
pygame.display.set_caption("Adrian's Python Test Code")

# Vars we'll need to calculate coords
winx = 0
winy = 0
gameRect = Rect(0,0,0,0)
windowName = "Ores - Puzzle Games at Miniclip.com - Play Free Online Games - Mozilla Firefox"
botState = "WaitingToStart"

# Returns an image of the grabbed windowName
def grabBrowser(windowName):
	global winx
	global winy
	
	try:
		browser_wnd = win32ui.FindWindow(None, windowName)
	except:
		print("Error finding window")
		exit()
		
	browser_wnd.BringWindowToTop()
	
	wDC = browser_wnd.GetWindowDC()
	(x, y, w, h) = browser_wnd.GetClientRect()
	
	(winx,winy,c,d) = browser_wnd.GetWindowRect()
	
	#dcObj=win32ui.CreateDCFromHandle(wDC)
	cDC=wDC.CreateCompatibleDC()
	dataBitMap = win32ui.CreateBitmap()
	dataBitMap.CreateCompatibleBitmap(wDC, w, h)
	cDC.SelectObject(dataBitMap)
	cDC.BitBlt((0,0),(w, h) , wDC, (0,0), win32con.SRCCOPY)
	
	#dataBitMap.SaveBitmapFile(cDC, "test.bmp")
	
	bmpbits = dataBitMap.GetBitmapBits(True)
	# True = Python string
	# False = Tuple of integers
	
	# Use pygame.image.fromstring on this.
	# or maybe .frombuffer
	
	# This took a couple of hours to work out.
	# Hack to swap the byte order around, then convert to
	# bytes for passing to fromstring.
	# Probably horribly inefficient
	rgb = []
	for i in range(0,len(bmpbits),4):
		rgb.append(bmpbits[i+2])
		rgb.append(bmpbits[i+1])
		rgb.append(bmpbits[i])
		rgb.append(bmpbits[i+3])
	
	screenImage = pygame.image.fromstring( (bytes(rgb)), (w,h), "RGBA")
	
	# NB Need to free the DCs
	wDC.DeleteDC()
	cDC.DeleteDC()
	#win32gui.ReleaseDC(browser_wnd, wDC)
	
	return screenImage

# Use the known pattern of the first few pixels in the top left corner
# of the game screen in order to find the screen. This allows for
# when different size ads show above the game or if the browser is
# moved or re-sized.
# Returns a rect of the game screen
def findGameOnScreen(screenImage):
	buff = pygame.PixelArray(screenImage)
	for y in range(0,screenImage.get_height()):
		for x in range(0, screenImage.get_width()):
			if buff[x][y] == screenImage.map_rgb((52,33,14)) and \
			   buff[x+1][y] == screenImage.map_rgb((149,89,32)) and \
			   buff[x+2][y] == screenImage.map_rgb((149,89,32)) and \
			   buff[x+3][y] == screenImage.map_rgb((149,89,32)) and \
			   buff[x+4][y] == screenImage.map_rgb((149,89,32)) and \
			   buff[x+5][y] == screenImage.map_rgb((52,33,14)) and \
			   buff[x+6][y] == screenImage.map_rgb((114,78,69)) and \
			   buff[x+7][y] == screenImage.map_rgb((142,98,86)):
				print("Found game: ", x, y)
				return pygame.Rect(x,y,640,480)
	print("Game not found")
	pygame.quit()
	sys.exit()
	return pygame.Rect(0,0,0,0)
			
def mouseClick(x,y):
	win32api.SetCursorPos((x, y))
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y)
	pygame.time.wait(100)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y)

# If the help screen has appeared (guessing by a pink pixel in the background
# of that screen) then press the Resume button
def dismissHelp():
	if tuple(screenImage.get_at((gameRect.left+480, gameRect.top+360))) == (181,143,133,255):
		print("Dismissing help screen")
		mouseClick(winx+gameRect.left+320, winy+gameRect.top+360)
		pygame.time.wait(1000)
		return True
	return False

def readOres():
	# Leave room for a border around grid
	oreGrid = zeros( (18,12),dtype=int16)
	for yOre in range(0,10):
		y = 4 + 66 + yOre*32
		for xOre in range(0, 16):
			x = 4 + 128 + xOre*32
			ore = 0
			oreColour = tuple(screenImage.get_at((gameRect.left+x, gameRect.top+y)))
			if oreColour == (200,200,200,255):
				ore = 1 # gray
			elif oreColour == (255,208,32,255):
				ore = 2 # yellow
			elif oreColour == (104,230,229,255):
				ore = 3 # blue
			elif oreColour == (125,220,113,255):
				ore = 4 # green
			elif oreColour == (204,102,102,255):
				ore = 5 # red
			#else:
			#	print(oreColour)
			oreGrid[xOre+1][yOre+1] = ore
	return oreGrid
	
# returns a count of how many squares join this
# one. Recursive
def flood(oreGrid, sqCol, xOre, yOre):
	count = 1
	# Mark square as visited
	oreGrid[xOre][yOre] = 0
	if oreGrid[xOre+1][yOre] == sqCol:
		count = count + flood(oreGrid, sqCol, xOre+1, yOre)
	if oreGrid[xOre-1][yOre] == sqCol:
		count = count + flood(oreGrid, sqCol, xOre-1, yOre)
	if oreGrid[xOre][yOre+1] == sqCol:
		count = count + flood(oreGrid, sqCol, xOre, yOre+1)
	if oreGrid[xOre][yOre-1] == sqCol:
		count = count + flood(oreGrid, sqCol, xOre, yOre-1)
	return count
	
def findBestMatch(oreGrid):
	bestMatch=(-1,-1)
	bestMatchCount = 0
	for yOre in range(1,11):
		for xOre in range(1,17):
			# Have we visited this square yet
			sqCol = oreGrid[xOre][yOre]
			if sqCol != 0:
				floodCount = 1 + flood(oreGrid, sqCol, xOre, yOre)
				if floodCount > bestMatchCount:
					bestMatchCount = floodCount
					bestMatch = (xOre-1, yOre-1)
	print("bmc=",bestMatchCount,"bm=", bestMatch)
	return bestMatch
				

def findBestMatch_RANDOM(oreGrid):
	return (randrange(16),randrange(10))
	
while True: # main game loop
	for event in pygame.event.get():
		if event.type == QUIT:
			pygame.quit()
			sys.exit()
		elif event.type == KEYDOWN:
			if event.key == K_SPACE:
				print("Pressed Space")
				botState = "StartPressed"
				
	
	if botState == "StartPressed":
		screenImage = grabBrowser(windowName)
		# screenImage = grabBrowser('Wikipedia, the free encyclopedia - Mozilla Firefox')
		#screenImage = grabBrowser('browsergrab2.bmp - Windows Photo Viewer')
		gameRect = findGameOnScreen(screenImage)
				
		#screenImage.set_at( (gameRect.left+480, gameRect.top+360), (255,255,255) )
		# Crops the game screen onto our display surface
		DISPLAYSURF.blit(screenImage, (0,0), gameRect)
				
		mouseClick(winx+gameRect.left+320, winy+gameRect.top+400) # Start button
		pygame.time.wait(2000)
		botState = "GameRunning"
	elif botState == "GameRunning":
		screenImage = grabBrowser(windowName)
		DISPLAYSURF.blit(screenImage, (0,0), gameRect)
		
		if tuple(screenImage.get_at((gameRect.left+0, gameRect.top+0))) == (255,255,255,255):
			print("Game Over")
			pygame.quit()
			sys.exit()
		
		if dismissHelp() == True:
			screenImage = grabBrowser(windowName)
		
		ores = readOres()
		print(ores)
		bm = findBestMatch(ores)
		if bm[0] != -1:
			# click the selected square
			mouseClick(winx+gameRect.left + 4 + 128 + bm[0]*32, winy+gameRect.top + 4 + 66 + bm[1]*32)
			# Move cursor away so highlight box doesn't obstruct view. Not sure if this happens anyway.
			# A better fix would be to sample colour of ore from nearer its centre
			win32api.SetCursorPos((winx+gameRect.left, winy+gameRect.top))
			# wait for debris to fall
			pygame.time.wait(250)
				
	pygame.display.update()
	#fpsClock.tick(30)