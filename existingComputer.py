import subprocess
import pandas as pd
import numpy as np
 

# from an excel sheet it reads the name of all the computers in a specified excel document
userSheet = input("enter excel document name: ")
userSheet = userSheet + ".xlsx"
sheet = pd.read_excel(userSheet, sheet_name = "Scan List")
devices = sheet['Name'].tolist()
output = []

# Creates Column Name
Columns = ["Devices", "Output"]

# Name of the output text file and excel file
output_file = "ping_results.txt"
excel_output = "Reachable.xlsx"

 

# Open the output file in write mode
with open(output_file, "w") as file:
    for device in devices:
        # Use the 'subprocess' module to run the 'ping' command
        result = subprocess.run(["ping", device], stdout=subprocess.PIPE, text=True)
        if "Ping request could not find host " in str(result):
            output.append("does not exist")
        elif "Approximate round trip times in milli-seconds:" in str(result):
            output.append("exists")
        else:
            output.append("does not exist")
 

        # Write the results to the output file
        file.write(f"Results for {device}:\n")
        file.write(result.stdout)
        file.write("\n")

 
df = pd.DataFrame(list(zip(devices, output)), columns = Columns) 
df.to_excel(excel_output)
print(f"Ping results saved to {output_file}")