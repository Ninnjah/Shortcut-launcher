import os
import time
import json
from win32com.client import Dispatch
from colorama import Fore, init
import easyTui as tui

class Json():                                           # Функции JSON

    def read(file):                                     ## Чтение JSON
        with open(file, "r", encoding='utf-8') as read_file:
            data = json.load(read_file)
        return data

    def write(file, data):                              ## Запись JSON
        with open(file, 'w', encoding='utf-8') as write_file:
            write_file.write(json.dumps(data))

    def config(links_list, links_dict):                 ## Сохранение настроек
        config_list = [links_list, links_dict]
        Json.write('settings.json', config_list)

def load_links(path, refresh):                          # Загрузка ярлыков
    '''
    Parsing all links in path folder
    and save into list
    '''
    if os.path.exists('settings.json') and refresh == False:
        links_dict = Json.read('settings.json')[1]
        print(links_dict)
    elif not os.path.exists('settings.json') or refresh == True:
        if not os.path.exists(path):
            os.mkdir(path)
        else:
            pass
        links = os.listdir(path)
        links_ = []
        for i in links:
            if '.url' in i or '.lnk' in i:
                links_.append(i)
        if len(links_) == 0:
            print(Fore.RED + "You don't have any shortcuts in the folder!\n")
            input('Press enter')
            Json.write('settings.json', [[],{}])
            time.sleep(1)
            main()
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
        Json.config(links_list, links_dict)

        return links_dict

def add_shortcut():                                     # Создание ярлыка
    os.system('cls||clear')
    print(tui.label('Adding new shortcut'))
    print(Fore.CYAN + 'Enter path to .exe file: ', end='')
    wDir = input()
    if os.path.exists(wDir):
        if '.exe' in wDir:
            target = wDir
            wDir = ''
            i = target
            while i[-1] != '\\':
                i = i[:-1]
            i = i[:-1]
            wDir = i
            link = ''
            i = target
            while i[-1] != '\\':
                link += i[-1]
                i = i[:-1]
            i = i[:-1]
            link = link[::-1]
            while link[-1] != '.':
                link = link[:-1]
            link = link[:-1]

            path = os.path.join("links\\{}.lnk".format(link))
            icon = target

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = wDir
            shortcut.IconLocation = icon
            shortcut.save()
        else:
            print(Fore.RED + "Wrong path!\npath must include .exe\n or Enter 'cancel' to cancel add..")
            time.sleep(1)
            add_shortcut()
    elif wDir in ['exit', 'cancel', 'stop', 'ex', 'x']:
        main()
    else:
        print(Fore.RED + "This path doesn't exists!\n or Enter 'cancel' to cancel add..")
        time.sleep(1)
        add_shortcut()

def remove_shortcut(links_list):                        # Удаление ярлыка
    index = input('Enter index: ')
    if index in ['ex', 'cancel', 'stop']:
        main()
    else:
        try:
            index = int(index)
        except ValueError:
            print('You must enter a number or "cancel" to cancel deleting..')
            time.sleep(2)
            remove_shortcut(links_list)
        except IndexError:
            print('This index does not exists')
            time.sleep(1)
            remove_shortcut(links_list)

    print(Fore.RED + 'Do you really want to delete "{}"?'.format(
        links_list[index]))
    com = input()

    if com in ['yes', 'y', 'yeah']:
        print('Removing {}...'.format(links_list[index]))
        os.remove('links\\' + links[links_list[index]])
        links_list.pop(index)
        Json.config(links_list, links)
        time.sleep(1)
        main()
    else:
        print('Removing canceled..')
        time.sleep(1)

def update_links(links_list):                           # Обновление списка ярлыков
    links = load_links('links', True)
    new_links_list = list(links.keys())
    result = list(set(new_links_list) - set(links_list))
    for i in result:
        links_list.append(i)
    result = list(set(links_list) - set(new_links_list))
    for i in result:
        links_list.pop(links_list.index(i))
    result = None
    Json.config(links_list, links)

def sort_links(links_list):                             # Перемещение ярлыков в списке
    from_ = input('from index: ')
    if from_ in ['exit', 'cancel', 'stop', 'ex', 'x']:
        main()
    to_ =  input('to index: ')
    if to_ in ['exit', 'cancel', 'stop', 'ex', 'x']:
        main()
    try:
        from_, to_ = int(from_), int(to_)
        links_list[from_], links_list[to_] = links_list[to_], links_list[from_]
        
        Json.config(links_list, links)
    except ValueError:
        print('You must enter numbers..\n or Enter "cancel" to cancel sort..')
        time.sleep(2)
        sort_links(links_list)
    except IndexError:
        print('This index does not exists')
        time.sleep(1)
        sort_links(links_list)

def rename_links(links, links_list):                    # Переименование ярлыка
    index = input('Enter index: ')
    if index in ['ex', 'cancel', 'stop']:
        main()
    else:
        try:
            index = int(index)
            old_name = links[links_list[index]]
            i = old_name
            file_ = ''
            while i[-1] != '.':
                file_ += i[-1]
                i = i[:-1]
            file_ += i[-1]
            file_ = file_[::-1]
        except ValueError:
            print('You must enter a number or "cancel" to cancel deleting..')
            time.sleep(2)
            rename_links(links, links_list)
        except IndexError:
            print('This index does not exists')
            time.sleep(1)
            rename_links(links, links_list)
    new_name = input('Enter new name: ')
    new_name += str(file_)
    os.rename('links\\' + old_name, 'links\\' + new_name)
    while new_name[-1] != '.':
        new_name = new_name[:-1]
    new_name = new_name[:-1]
    links_list.insert(index, new_name)
    update_links(links_list)

def run(app):                                           # Запуск приложения
    os.startfile(os.path.join('links', app))
    
def main():                                             # Интерфейс
    json_ = Json.read('settings.json')
    links_list = json_[0]
    links = json_[1]

    os.system('cls||clear')
    print(tui.label('Game Launcher v{}'.format(version)))
    if len(links_list) > 0:
        print(Fore.CYAN + tui.ol(links_list))
    else:
        print(Fore.CYAN + tui.ul(["You can add them with 'add' command", 
            "or manualy add shortcuts in 'links' folder\n  You can open folder with 'links' command"]))
    cmd = input('Enter number: ').lower()
    try:
        cmd = int(cmd)
        print('Starting {}...'.format(links_list[cmd]))
        run(links[links_list[cmd]])
        time.sleep(2)
        main()
    except ValueError:
        if cmd in ['ex', 'exit', 'quit', 'x']:
            print('quiting...')
            time.sleep(.5)
            quit()
        elif cmd in ['update', 'refresh', 'upd']:
            update_links(links_list)
            main()
        elif cmd in ['steam']:
            print('Starting {}...'.format(cmd))
            run('sys/steam.lnk')
            time.sleep(2)
            main()
        elif cmd in ['epic', 'epic games']:
            print('Starting {}...'.format(cmd))
            run('sys/epic.lnk')
            time.sleep(2)
            main()
        elif cmd in ['remove', 'delete', 'rm']:
            remove_shortcut(links_list)
            main()
        elif cmd in ['links', 'lnk', 'folder']:
            print('Opening {}...'.format(cmd))
            os.system("explorer.exe {}".format(os.getcwd() + '\\links'))
            time.sleep(1)
            main()
        elif cmd in ['sort', 'move']:
            sort_links(links_list)
            main()
        elif cmd in ['add', 'create', 'make', 'mk']:
            add_shortcut()
            update_links(links_list)
            main()
        elif cmd in ['help']:
            os.system('cls||clear')
            command_list = [
                '"exit" ("quit", "ex", "x")\n\tUsing to quit the program',
                '"update" ("refresh", "upd")\n\tUsing to update shortcut list',
                '"remove" ("delete", "rm")\n\tUsing to delete app shortcut',
                '"add" ("create", "make", "mk")\n\tUsing to add new shortcut',
                '"sort" ("move")\n\tUsing to move shortcut in list',
                '"links" ("folder", "lnk")\n\tUsing to open shortcuts folder',
                '"steam"\n\tUsing to start steam from "links/sys/steam.lnk"',
                '"epic" ("epic games")\n\tUsing to start Epic games store from "links/sys/epic.lnk"',]
            print(tui.label('Help'))
            print(Fore.CYAN + tui.ul(command_list))
            input('Press enter to go to main menu')
            main()
        elif cmd in ['rename']:
            rename_links(json_[1], json_[0])
            main()

        else:
            print('You must enter a Number..')
            time.sleep(1)
            main()
    except IndexError:
        print('This index does not exists')
        time.sleep(1)
        main()

if __name__ == '__main__':                              # Запуск
    init(autoreset=True)
    version = '0.7 beta'
    links = load_links('links', False)

    main() 
