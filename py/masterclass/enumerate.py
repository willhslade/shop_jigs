data = [4,5,104,110,300,120,4,150]

min_valid = 100
max_valid = 200

top_index = len(data) - 1
for index, value in enumerate(reversed(data)):
 	actual_index = top_index - index
 	# print(index, value)
 	if (value < min_valid) or (value > max_valid):
 		del data[actual_index]

print(data)
