import os
#import sys
import time
import json
#from PyQt5 import QtWidgets, uic, QtGui, QtCore
import configparser
from colorama import Fore, init, Cursor
import easyTui as tui
#from design import Ui_MainWindow

def check_links(path, refresh):                            # Проверка наличия ярлыков
    '''
    Parsing all links in path folder
    and save into list
    '''
    if os.path.exists('settings.json') and refresh == False:
        links_dict = json_read('settings.json')[1]
        print(links_dict)
    elif not os.path.exists('settings.json') or refresh == True:
        links = os.listdir(path)
        links_ = []
        for i in links:
            if '.url' in i or '.lnk' in i:
                links_.append(i)
        if len(links_) == 0:
            quit()
        links = []
        for i in links_:
            while i[-1] != '.':
                i = i[:-1]
            i = i[:-1]
            links.append(i)

        links_dict={}
        for i in range(len(links)):
            links_dict.setdefault(links[i], links_[i])

        links_list = list(links_dict.keys())
        json_config(links_list, links_dict)

        return links_dict

def json_config(links_list, links_dict):
    config_list = [links_list, links_dict]
    json_write('settings.json', config_list)

def json_read(file):                                    # Чтение JSON
    with open(file, "r", encoding='utf-8') as read_file:
        data = json.load(read_file)
    return data

def json_write(file, data):                             # Запись JSON
    with open(file, 'w', encoding='utf-8') as write_file:
        write_file.write(json.dumps(data))

def run(app):
    os.startfile(os.path.join('links', app))
    
def main():
    version = '0.5 beta'
    json_ = json_read('settings.json')
    links_list = json_[0]
    links = json_[1]

    os.system('cls||clear')
    print(tui.label('Game Launcher v{}'.format(version)))
    print(Fore.CYAN + tui.ol(links_list))
    cmd = input('Enter number: ')
    try:
        cmd = int(cmd)
        print('Starting {}...'.format(links_list[cmd]))
        run(links[links_list[cmd]])
        time.sleep(2)
        main()
    except ValueError:
        if cmd.lower() in ['ex', 'exit', 'quit', 'x']:
            print('quiting...')
            time.sleep(.5)
            quit()
        elif cmd.lower() in ['update', 'refresh', 'upd']:
            links = check_links('links', True)
            new_links_list = list(links.keys())
            result = list(set(new_links_list) - set(links_list))
            for i in result:
                links_list.append(i)
            result = list(set(links_list) - set(new_links_list))
            for i in result:
                links_list.pop(links_list.index(i))
            result = None
            json_config(links_list, links)
            main()
        elif cmd.lower() in ['steam']:
            print('Starting {}...'.format(cmd))
            run('sys/steam.lnk')
            time.sleep(2)
            main()
        elif cmd.lower() in ['epic', 'epic games']:
            print('Starting {}...'.format(cmd))
            run('sys/epic.lnk')
            time.sleep(2)
            main()
        elif cmd.lower() in ['remove', 'delete', 'rm']:
            def check(index):
                if index in ['ex', 'cancel', 'stop']:
                    main()
                else:
                    try:
                        index = int(index)
                    except ValueError:
                        print('You must enter a number or "cancel" to cancel deleting..')
                        time.sleep(2)
                        check(index)
                return index

            index = input('Enter index: ')
            index = check(index)
            print(Fore.RED + 'Do you really want to delete "{}"?'.format(
                links_list[index]))
            com = input()

            if com in ['yes', 'y', 'yeah']:
                print('Removing {}...'.format(links_list[index]))
                os.remove('links\\' + links[links_list[index]])
                links_list.pop(index)
                json_config(links_list, links)
                time.sleep(1)
                main()
            else:
                print('Removing canceled..')
                time.sleep(1)
                main()

        elif cmd.lower() in ['links', 'lnk', 'folder']:
            print('Opening {}...'.format(cmd))
            print(os.getcwd())
            os.system("explorer.exe {}".format(os.getcwd() + '\\links'))
            print("explorer.exe {}".format(os.getcwd() + '\\links'))
            time.sleep(10)
            main()
        elif cmd.lower() in ['sort']:
            from_ = input('from index: ')
            to_ =  input('to index: ')
            try:
                from_, to_ = int(from_), int(to_)
                links_list[from_], links_list[to_] = links_list[to_], links_list[from_]
                
                json_config(links_list, links)
                main()
            except ValueError:
                print('You must enter numbers..')
                time.sleep(2)
                main()
        else:
            print('You must enter a Number..')
            time.sleep(1)
            main()
'''
class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
'''
if __name__ == '__main__':
    init(autoreset=True)
    links = check_links('links', False)
    #app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    #window = Ui()  # Создаём объект класса ExampleApp
    #window.show()  # Показываем окно
    #app.exec_()  # и запускаем приложение

    main()
