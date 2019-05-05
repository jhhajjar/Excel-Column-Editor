import pandas
import os
import time

# Funciton to write instructions to the 'instructions.txt'
def writeinstr():
    filename = "instructions.txt"
    fp = open(filename, 'a')
    fp.write()
    fp.close()

# Function to make a new instruction
# Params:
#       arr, the parsed instructions to write
#       name, the name of the instruction to write
def newinstr(arr, name):
    file = open('instructions.txt', 'a')

    # Create the array with the correct formatting
    arrtostr = arr[0].strip()
    for col in arr[1:]:
        arrtostr = arrtostr + "," + col.strip()

    # Write the line [name]:[instruction]
    print("Writing instruction...")
    stringtowrite = name + ":" + arrtostr +'\n'
    file.write(stringtowrite)
    print("Instruction written successfully, saved as " + name + ".")

    file.close()

# Function to run a spreadsheet through the instruction
# Params:
#       sheet, the spreadsheet to run through the instruction
#       name, the name of the instruction to run
def runinstr(sheet, name):
    file = open('instructions.txt', 'r')
    instr = []

    # Read in the instruction from the file
    for line in file:
        if(line.split(":")[0] == name):
            instr = line.split(":")[1].strip('\n').split(",")
            break

    file.close()
    if(instr == []):
        print("This instruction doesn't exist.")
        return

    # Run the instruction on the sheet
    sheet = sheet.reindex(columns=instr)
    return sheet

# Function to save the new file
# Params:
#       sheet, the new sheet to save
#       writeto, the file to save to
def savefile(sheet, writeto):
    write = pandas.ExcelWriter(writeto)
    sheet.to_excel(write, index=False)
    write.save()

# Function to get the filename to read from
def getFile():
    filename = ""
    sheet = None
    while 1:
        try:
            print("Enter the name of the file to read from:")
            filename = input(os.getcwd()+"\\") + ".xlsx"

            print("Reading file...")
            sheet = pandas.read_excel(filename)
            sheet.columns = sheet.columns.str.upper()
            print("File read successfully.")
            break
        except Exception as e:
            print("Error reading " + filename + ":")
            print(e)
            print()
    return (filename, sheet)

def main():
    # Background on the program
    print("This script will allow you to:")
    print("- create instructions on which columns to send to a separate spreadsheet.")
    print("- run a spreadsheet through an existing set of instructions.")

    # Main loop
    while 1:
        # Options
        print("A> Set new instructions")
        print("B> Run a spreadsheet through existing instructions")
        print("X> Done")
        option = input().upper()

        # Set new instructions
        if(option == 'A'):
            # Get name of instruction
            name = input("Enter the name of this instruction (case insensitive): ").lower()
            print("Enter the names of the columns in the new spreadsheet separated by commas: ")

            # Get columns
            arr = input().upper().split(',')
            newinstr(arr, name)

        # Run existing instructions
        elif(option == 'B'):
            # Get file and name of instruction
            (filename, sheet) = getFile()
            name = input("Enter the name of the instruction (case insensitive): ").lower()

            # Run instruction
            sheet = runinstr(sheet, name)

            # Save file
            print("Saving file...")
            fileto = filename[:filename.find(".")] + "_" + name + ".xlsx"
            savefile(sheet, fileto)
            print("File saved as " + fileto + ".")

        # Exit script
        elif(option == 'X'):
            print("Exiting...")
            time.sleep(1)
            break

if __name__ == '__main__':
    main()
