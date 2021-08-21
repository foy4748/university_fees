import csv
import json

with open('./compiled.csv', newline='') as f:
    reader = csv.reader(f)
    storage = dict()
    for row in reader:
        if row[2] == '':
            continue
        university = row[0].split()
        uni = ' '.join(university)
        if uni not in storage:
            storage[uni] = {'dept':[row[1]], 'fee':[row[2]] }
        else:
            storage[uni]['dept'].append(row[1])
            storage[uni]['fee'].append(row[2])

with open('compiled.json', 'w') as f:
    json.dump(storage, f)


