from statistics import mean
import matplotlib
matplotlib.use('Agg')  # Non-GUI backend for matplotlib

import Dice_Simulation as DicePyFile
import Family_Simulation as FamilyPyFile
import Marbles_Drop_Simulation as MarblesPyFile
import Monte_Carlo_Simulation as MontePyFile

#Main function to run the script
def Main():
    while True:
        print("\n=============================== Group C Coursework ===============================\n")
        print("1. Run Monte Carlo Simulation")
        print("2. Run Dice Simulation")
        print("3. Run Family Simulation")
        print("4. Close the Program")
        print("--------------------------------------------------------")
        
        try:
            choice = int(input("Please choose an option (1-4): "))
            print("--------------------------------------------------------\n")
            
            if choice == 1:
                MontePyFile.Monte_Carlo_Simulation()
            elif choice == 2:
                print("Please wait...")
                DicePyFile.dice_simulation_Main()
            elif choice == 3:
                FamilyPyFile.familySimulation_Main()
            elif choice == 4:
                print("Closing the program. Goodbye!")
                break
            else:
                print("Invalid choice. Please enter a number between 1 and 4.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")


# Call Main() to run the script
if __name__ == "__main__":
    Main()



