from collections import Counter

capitals = {'UK': 'London', 'France': 'Paris', 'Germany': 'Berlin'}


print(capitals)
print(len(capitals))
print(capitals['UK'])
print(capitals['France'])
print(capitals.keys())
print(capitals.values())
capitals['NZ'] = 'Auckland'
print(capitals)

dict_values = list([[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 6, 6, 6, 7, 7, 7, 7, 9, 10, 4, 5], [1, 1, 1, 2, 2, 2, 2, 2, 6, 8], [1], [2], [2, 5, 6, 6, 7], [2, 4], [2], [2, 3, 3, 3, 3, 4, 4, 4, 5, 6, 6, 9, 4, 2], [4], [4], [5], [5], [5], [6], [6], [6], [6]])
print(type(dict_values))

for counter in dict_values:
    frequencies = Counter(counter)
    print(counter)
    print(frequencies)



