#a = "Silver spoon"
#print(a)
#print(a.upper())
#print(a.lower())
#print(a.strip("!"))
#print(a.rstrip("!"))
#print(a.replace("sp", "m"))
#print(a.split(" "))
#blogHeading = "inTroDctIon tO j"
#print(blogHeading.capitalize())

#str = "Welcome to the Console!!!"
#print(len(str))
#print(len(str.center(50)))
#print(a.count("Harry"))

#ALPHABETS = "A B C D E"
#print(ALPHABETS)
#b ="welcome to the console!!!"
#countstr = b. count("o")
#print(countstr)
#print(b.endswith("to",4,10))
#str1 ="He's name is Dan. He is an honest man."
#rint(str1.find("is"))

#str1 ="HenameisDan9"
#print(str1.isalnum())

#str1 ="HenameisDan"
#print(str1.isalpha())

#str1 = "We wish you a Merry christmas\n"

#print(str1.isprintable())


#str1 = "    "
#print(str1.isspace())
#str2 = "    "
#print(str2.isspace())

#str1 = "World Health Organization"
#print(str1.istitle())

#str1 = "WORLD HEALTH ORGANIZATION"
#print(str1.isupper())

#str1 = "Python is a Compiled Language."
#print(str1.replace("Compiled", "Interpreted"))

#str1 = "Python is Interpreted Language"
#print(str1.startswith("Python"))


# use ord() function to find the Unicode of

#Unicode
# UpperCase_Char =ord( 'A')
# lowercase_char = ord('a')
# Special_Char = ord('$')
# num = ord('9')

# print('The Unicodes are:', UpperCase_Char, lowercase_char, Special_Char, num)

# use chr() function to find the characters corresponding to the given Unicode

# num = [65, 97, 36, 57]
# for i in num:
#   print(chr(i))



#For Loop
#languages = ['Swift', 'Python', 'Go', 'Javascript']

#rung a loop for each item of the list
#for language in languages:
#    print(language)



import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15,1)),
                            sg.Checkbox('German', key='German'),
                            sg.Checkbox('Spanish', key='Spanish'),
                            sg.Checkbox('English', key='English')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()



