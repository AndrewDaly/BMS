import re
def hasNumber(string):
    return bool(re.search(r'\d', string))

with open("12152021.txt", "r", encoding="utf-8") as file:
    lines = file.readlines()
count = 0
# Strips the newline character
for line in lines:
    count += 1
    if("US" in line and hasNumber(line)):
        print(line.rstrip("\n"))