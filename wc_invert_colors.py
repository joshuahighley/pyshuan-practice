# CODING SLOTH WEEKLY CHALLENGE:
# Invert Colors
# Create a function that inverts the rgb values of a given list

# Examples (I've formatted these to use below)
exampleOne = [255, 255, 255] # output should = [0, 0, 0]
exampleTwo = [0, 0, 0] # output should = [255, 255, 255]
exampleThree = [165, 170, 221] # output should = [90, 85, 34]



# SOLUTION:

# list comprising the 256 rgb values
cD = list(range(0, 255, 1))

# the function that inverts given RGB values in a given list
def color_invert(rgb):
    output = []
    for x in rgb:
        output.append(cD[0-(x)])
    print(output)


color_invert(exampleThree)
