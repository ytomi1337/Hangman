from openpyxl import  Workbook, load_workbook
import random
import string
excel_file = load_workbook('words.xlsx')
excel_sheet = excel_file.active

def word_roll():
    words_list = []
    for cell in excel_sheet['A']:
        words_list.append(cell.value)

    random_word = random.choice(words_list)

    return str(random_word).upper()

def hangman_quess():
    word = word_roll()
    word_letters = list(word)
    alphabet = set(string.ascii_uppercase)
    used_letter = set()
    lifes =['<3', '<3']

    for cell in excel_sheet['A']:
        if cell.value == str(word):
            nr_cell = str(cell.coordinate)
            nr_cell = nr_cell.lower().replace("a", "B")

            category = excel_sheet[nr_cell].value


    while len(word_letters) > 0:

        if lifes == []:
            print(f'Koniec żyć, przegrana ! słowo: {word}')
            break
        else:
            print(word)
            print('Kategoria:',category,'Zycia:',' '.join(lifes))
            print('Zużyte litery: ', ' '.join(used_letter))
            word_blank = [letter if letter in used_letter else '-' for letter in word]
            print('Słowo: ', ' '.join(word_blank))
            answer = input("Zgadnij wyraz: ").upper()

            if answer in alphabet - used_letter:
                used_letter.add(answer)
                if answer in word_letters:
                    for i,x in enumerate(word_letters):
                        if x == answer:
                            word_letters.remove(answer)
                else:
                    lifes.remove('<3')

            elif answer in used_letter:
                print(f'Literka: {answer} Została już wprowadzona!')
            else:
                print('Nie prawidłowa wartość!')










def main():
    choice_list = ['1', '2', '3', '4']

    while True:
        print("1: Zagraj")
        print("2: Dodaj Słowo")
        print("3: Historia Gier")
        print("4: Wyjdź")
        print("")
        choice = input("Wybór: ")

        if choice == '1':
            hangman_quess()
        elif choice == '2':
            print("Dodaj")
        elif choice == '3':
            print("Historia gier")
        elif choice == '4':
            print("opuszczanie")
        elif choice is not choice_list:
            print("Nie prawidłowy wybór ! Wpisz ponownie")
            print("")

main()
