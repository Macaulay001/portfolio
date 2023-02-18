fileName = 'input.txt'
text = 'text to find'
fileHandle = open(fileName, "r")
lines = fileHandle.readlines()
for line in lines:
  if text not in line:
    with open('output.txt', 'a') as f:
        f.write(line)
        print(line)
fileHandle.close()
