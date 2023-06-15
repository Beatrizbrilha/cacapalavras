import openpyxl
import random
import string

#Função para gerar as palavras e a grade
def generate_word_search(words):
    size = 15  # Tamanho da grade
    grid = [[' ' for _ in range(size)] for _ in range(size)]  # Grade vazia

# Para adicionar as palavras à grade
    for word in words:
        word_len = len(word)
        placed = False

        while not placed:
            step_x = random.choice([-1, 0, 1])  # Direção aleatória para x
            step_y = random.choice([-1, 0, 1])  # Direção aleatória para y

            if step_x == 0 and step_y == 0:
                continue

            start_x = random.randint(0, size - 1)
            start_y = random.randint(0, size - 1)

            end_x = start_x + (word_len - 1) * step_x
            end_y = start_y + (word_len - 1) * step_y

            if end_x < 0 or end_x >= size or end_y < 0 or end_y >= size:
                continue

            cells_available = True
            for i in range(word_len):
                x = start_x + i * step_x
                y = start_y + i * step_y
                if grid[x][y] != ' ' and grid[x][y] != word[i]:
                    cells_available = False
                    break

            if cells_available:
                for i in range(word_len):
                    x = start_x + i * step_x
                    y = start_y + i * step_y
                    grid[x][y] = word[i]
                placed = True

# Para preenche as células vazias com letras aleatórias
    for i in range(size):
        for j in range(size):
            if grid[i][j] == ' ':
                grid[i][j] = random.choice(string.ascii_uppercase)

# Para criar um novo arquivo do Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active

# Para add a grade ao arquivo do Excel
    for i, row in enumerate(grid):
        for j, cell in enumerate(row):
            sheet.cell(row=i+1, column=j+1).value = cell

    # Salva o arquivo do Excel
    workbook.save('cacapalavras.xlsx')

# A Lista de palavras para o caça-palavras
palavras = ['SUPERMARIO', 'CRASH', 'BATMAN', 'BOMB']

# Para gerar o caça-palavras e salva no arquivo do Excel
generate_word_search(palavras)
