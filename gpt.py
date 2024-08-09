import turtle
from PIL import Image

# Set the size of the pixel art
SIZE = 32

# Create a turtle object and set the speed to the maximum
t = turtle.Turtle()
t.speed(0)

# Set the pen color to red
t.pencolor("red")

# Move the turtle to the starting position
t.penup()
t.setpos(-SIZE/2, 0)
t.pendown()

# Draw the shape of the heart using turtle commands
t.left(140)
t.forward(SIZE)
t.circle(-SIZE/2, 200)
t.seth(60)
t.circle(-SIZE/2, 200)
t.forward(SIZE)

# Hide the turtle and set the background color to white
t.hideturtle()
#t.bgcolor("white")

# Get the screen object and set the size of the window
screen = turtle.getscreen()
screen.setup(SIZE, SIZE)

# Get the image from the screen and save it as a png file
ts = screen.getcanvas()
image = Image.frombytes("RGB", (SIZE, SIZE), ts.tostring())
image.save("heart.png")
