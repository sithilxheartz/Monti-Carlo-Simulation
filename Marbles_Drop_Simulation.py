import random
import math
from statistics import mean
import matplotlib
matplotlib.use('Agg')  # Non-GUI backend for matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment


def mcs_MarblesDropSimulation():

    try:
        print("!! This section will not effect monte carlo simulation calculations !!")
        count = int(input("Enter the number of marbles to drop: "))
        if count < 1:
            print("Please enter a positive integer greater than 0.")
            return
        DrawTable(count)
    except ValueError:
        print("Invalid input. Please enter a valid positive integer.")
        return

def simulation(RunCount):
    # Circle center = (0,0) and Radius = 1
    # Rectangle:  x = 2 to 3, y = -0.5 to 0.5

    RectangleDropCount = 0
    CircleDropCount = 0
    RectanglePoints = []
    CirclePoints = []
    OutOfBoundsPoints = []

    for _ in range(RunCount):
        x = random.uniform(-2,4)
        y = random.uniform(-2,2)

        if(x > 2 and x < 3 and y > -0.5 and y < 0.5):
            RectangleDropCount += 1
            RectanglePoints.append((x, y))
        elif(x**2 + y**2 <= 1):
            CircleDropCount += 1
            CirclePoints.append((x, y))
        else:
            OutOfBoundsPoints.append((x, y))

    return RectangleDropCount, CircleDropCount, RectanglePoints, CirclePoints, OutOfBoundsPoints

def DrawTable(RunCount = 100000):
    RectangleDropCount, CircleDropCount, RectanglePoints, CirclePoints, OutOfBoundsPoints = simulation(RunCount)

    # Separate x and y coordinates for plotting
    rect_x, rect_y = zip(*RectanglePoints)
    circ_x, circ_y = zip(*CirclePoints)
    outOB_x, outOB_y = zip(*OutOfBoundsPoints)

    # Plot the points
    plt.figure(figsize=(12, 9))
    plt.scatter(outOB_x, outOB_y, color='black', s=0.5, alpha=0.3, label="Out of Bounds Points")
    plt.scatter(rect_x, rect_y, color='green', s=.5, alpha=1, label="Rectangle Points")
    plt.scatter(circ_x, circ_y, color='red', s=.5, alpha=1, label="Circle Points")

    # Configure plot
    plt.xlim(-2, 4)
    plt.ylim(-2, 2)
    plt.axhline(0, color='black', linewidth=1, linestyle='-')  # x-axis
    plt.axvline(0, color='black', linewidth=1, linestyle='-')  # y-axis
    plt.title(f"Monte Carlo Simulation with {RunCount} Points")
    plt.xlabel("x-coordinate")
    plt.ylabel("y-coordinate")

    # Add the legend outside the plot
    plt.legend(
        loc='upper center', 
        bbox_to_anchor=(0.5, -0.1), 
        ncol=3, 
        fontsize=10, 
        frameon=True,
        markerscale=10,  # Increases the size of legend markers
        scatterpoints=1  # Number of points displayed for scatter legend
    )
    # plt.show()
    
    # Save the plot as a PNG file 
    plt.savefig("monte_carlo_simulation.png")
    print("Plot saved as 'monte_carlo_simulation.png'")


if __name__ == '__main__':
    mcs_MarblesDropSimulation()
