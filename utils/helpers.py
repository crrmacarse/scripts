def write_to_file(filePath, data):
    with open(filePath, 'w') as file:
        file.write(data)
        print(f"Data written to {filePath}")