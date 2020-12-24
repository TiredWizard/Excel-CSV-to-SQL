import sys
import os
import signal
import colorama
import ExcelCsv2SQL as settings

colorama.init()  # Color init per Windows


# =======================
#    Definizioni
# =======================

if not os.path.exists("SQL"):
    os.makedirs("SQL")
sql_file_path = "SQL/default.sql"

colors = {
    "info": "35m",  # Orange - info messages
    "error": "31m",  # Red - error messages
    "ok": "32m",  # Green - success messages
    "menu2c": "\033[46m",  # Light blue menu
    "menu1c": "\033[44m",  # Blue menu
    "close": "\033[0m"  # Color coding close
}
cc = "\033[0m"
ct = "\033[101m"
cs = "\033[41m"
c1 = colors["menu1c"]
c2 = colors["menu2c"]

# =======================
#    USER CONFIG
# =======================
programtitle = "Excel/CSV to SQL"

# Fill the options as needed.
menu2_colors = {
    "ct": ct,
    "cs": cs,
    "opt": c2
}

menu2_options = {
    "title": "Menu SQL",
    "6": "Imposta file predefinito",
    "7": "Pulisci file",
    "8": "Torna al menu principale",
    "0": "Esci (or use CNTRL+C)",
}

menu1_colors = {
    "ct": ct,
    "cs": cs,
    "opt": c1
}
menu1_options = {
    "title": "Menu Principale",
    "1": "Crea tabella",
    "2": "Importa da Excel (xls)",
    "3": "Importa da Excel (xlsx)",
    "4": "Importa da CSV",
    "5": "Impostazioni file SQL",
    "0": "Esci (or use CNTRL+C)",
}


# =======================
#      HELPERS
# =======================
def printWithColor(color, string):
    print("\033[" + colors[color] + " " + string + cc)


def printError():
    printWithColor("error", "Errore!!")
    return 1


def printSuccess():
    printWithColor("ok", "Success!!")
    return 0


# Uscita dal programma
def exit():
    sys.exit()


# Uscita del programma se si preme control+c
def sigint_handler(signum, frame):
    print("CNTRL+C exit")
    sys.exit(0)


# =======================
#      Azioni
# =======================

# Menu template
def action(ch):
    global sql_file_path
    if ch == '1':
        settings.create_table(sql_file_path)
    elif ch == '2':
        file_path = input("Inserire la path del file excel (xls): ")

        fogli = list()
        while True:
            foglio = input("inserire il nome dei fogli - end per completare l'inserimento: ")
            if foglio == "end":
                break
            else:
                fogli.append(foglio)

        tabelle = list()
        while True:
            tabella = input("inserire il nome delle tabelle - end per completare l'inserimento: ")
            if tabella == "end":
                break
            else:
                tabelle.append(tabella)

        settings.convert_excel_xlsx(file_path, sql_file_path, fogli, tabelle)
    elif ch == '3':
        file_path = input("Inserire la path del file excel (xlsx): ")

        fogli = list()
        while True:
            foglio = input("inserire il nome dei fogli - end per completare l'inserimento: ")
            if foglio == "end":
                break
            else:
                fogli.append(foglio)

        tabelle = list()
        while True:
            tabella = input("inserire il nome delle tabelle - end per completare l'inserimento: ")
            if tabella == "end":
                break
            else:
                tabelle.append(tabella)

        settings.convert_excel_xlsx(file_path, sql_file_path, fogli, tabelle)
    elif ch == '4':
        file_path = input("Inserire la path del file excel (xlsx): ")
        table_name = input("Inserire il nome della tabella: ")
        attributi = list()
        while True:
            attributo = input("inserire il nuovo attributo - end per completare l'inserimento: ")
            if attributo == "end":
                break
            else:
                attributi.append(attributo)

        settings.convert_csv(file_path, sql_file_path, table_name, attributi)
    elif ch == '6':
        sql_file_path = input("inserire il nome del file (senza il formato): ")
        sql_file_path = "SQL/" + sql_file_path + ".sql"
    elif ch == '7':
        settings.clean_file(sql_file_path)
    elif ch == '':
        pass  # mostra di nuovo il menu
    elif ch == '0':
        sys.exit()
    else:
        printError()


class menu_template:
    def __init__(self, options, colors):
        self.menu_width = 50  # width dei caratteri nel menu
        self.options = options
        self.colors = colors

    # =======================
    #      Print dei menu
    # =======================

    def createMenuLine(self, letter, color, length, text):
        menu = color + " [" + letter + "] " + text
        line = " " * (length - len(menu))
        return menu + line + cc

    def createMenu(self, size):
        line = self.colors["ct"] + " " + programtitle
        line += " " * (size - len(programtitle) - 6)
        line += cc
        print(line)  # Title
        line = self.colors["cs"] + " " + self.options["title"]
        line += " " * (size - len(self.options["title"]) - 6)
        line += cc
        print(line)  # Subtitle
        for key in self.options:
            if key != "title":
                print(self.createMenuLine(key, self.colors["opt"], size, self.options[key]))

    def printMenu(self):
        self.createMenu(self.menu_width)


class menu1(menu_template):
    pass


class menu2(menu_template):
    pass


# =======================
#      Programma principale
# =======================

class menu_handler:

    def __init__(self):
        self.current_menu = "main"
        self.m1 = menu1(menu1_options, menu1_colors)
        self.m2 = menu2(menu2_options, menu2_colors)

    def menuExecution(self):
        if self.current_menu == "main":
            self.m1.printMenu()
        else:
            self.m2.printMenu()
        choice = input(" >> ")
        if self.current_menu == "main":
            if choice == "5":
                self.current_menu = "second"
            else:
                self.actuator(0, choice)
        else:
            if choice == '8':
                self.current_menu = "main"
            else:
                self.actuator(1, choice)
        print("\n")

    def actuator(self, type, ch):
        if type == 0:
            action(ch)
        else:
            action(ch)


# Programma principale
if __name__ == "__main__":
    menu = menu_handler()
    signal.signal(signal.SIGINT, sigint_handler)
    while True:
        menu.menuExecution()
