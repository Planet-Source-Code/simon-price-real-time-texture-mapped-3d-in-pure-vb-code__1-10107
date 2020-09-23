TEXTURED 3D IN PURE VISUAL BASIC CODE !!! - BY SIMON PRICE

IMPORTANT

PLEASE TEST THE PROGRAM BY CLOSING ALL OTHER APPLICATIONS AND RUNNING THE SUPPLIED .EXE! DO NOT RUN FROM VB, IT'S SLOWER! FOLLOWING THESE INSTRUCTIONS SHOULD GET THE PROGRAM RUNNING AS FAST AS POSSIBLE!

What's this program do?

This program allows you spin a 3D cube in real time at high frame rates with a choice of several different rendering modes. Here is a list of rendering modes and the speeds in FPS (frames per second) for each on my 400Mhz PC.

Corners Only - 135 FPS
WireFrame - 120 FPS
Outline - 110 FPS
Flat Colour Polygons - 40 FPS
Single Texture Mapped Polygons - 20 FPS
Multiple Texture Mapped Polygons - 11 FPS

Please note that this program does not use DirectX, OpenGL, DLL's or controls (other than the built in VB controls). It only uses API's during loading. It uses no drawing API's. It even shows you how to draw in an image box (which is faster).

How's it soooo fast?

It draws in an image box which is simpler and faster. It's kinda like DirectDraw, for example, to draw a line, it doesn't waste time asking Windows to draw a line by calling the API 'LineTo', it just goes and draws the line itself. Cool huh? Also, the maths is (in my opinion!) well optimised. 

You must have cheated somewhere!

OK, the texture mapping is not dead accurate, I have used a method that doesn't take into account the z co-ord of each pixel, like, it's 2D texture mapping instead of 3D. But it's hard to spot the difference and speeds the program up loads so I'll allow a bit of inaccuarcy.


by Simon Price 26/7/00
email : Simon@HisPalace.fsbusiness.co.uk




