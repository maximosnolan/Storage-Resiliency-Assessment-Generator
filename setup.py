
x = int(input("How far would you like to travel in miles?"))
if x < 0:
    print("This is a negative distance :(")
elif x < 3:
    print("You should walk!")
elif x < 300:
    print("You should drive!")
else:
    print("You should fly!")