import json
help('modules')
material = {
    'Тип' : 'Резина',
    'Шифр(дата)': '29',
    'Шифр(полный)': '29-3374',
    'Назначение': 'технол прослойка',
    'Массовая плотность': 1.132,
    'Твёрдость по шору': 61,
    'C10': 0.132297,
    'C01': 0.619986,
    'K': 7522.84
}
with open("data_file.json", "r") as read_file:
    materials = json.load(read_file)
print(len(materials))
f = open('D:\\buffer\\data\\Rezina_29-3374_karkas.txt')
odnoos = []
for line in f:
    l = line.split()
    odnoos.append(l)
material['Одноосное растяжение'] = odnoos
if material not in materials:
    materials.append(material)
else:
    print('Материал уже есть в базе')
with open("data_file.json", "w") as write_file:
    json.dump(materials, write_file)

