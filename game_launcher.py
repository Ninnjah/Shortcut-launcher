import os
import time
import json
from colorama import Fore, init
import easyTui as tui

#todo:
#todo   1: Добавить категории
#todo
#todo   ~2: Изменить запуск програм на 'start {index}'
#todo

###### Функции JSON ######
## Чтение JSON
def Jsonread(file):                               
    with open(file, "r", encoding='utf-8') as read_file:
        data = json.load(read_file)
    return data

## Запись JSON
def Jsonwrite(file, data):                        
    with open(file, 'w', encoding='utf-8') as write_file:
        write_file.write(json.dumps(data))

## Сохранение настроек
def Jsonconfig(links_list, links_dict, sys_links):
    config_list = [links_list, links_dict, sys_links]
    Jsonwrite('settings.json', config_list)
###### Функции JSON ######

def load_links(path, refresh):                          # Загрузка ярлыков
    '''
    Parsing all links in path folder
    and save into list
    '''
    links_ = []
    links = []
    sys_links = []
    links_dict={}
    if os.path.exists('settings.json') and refresh == False:
        links_dict = Jsonread('settings.json')[1]
        return links_dict
    elif not os.path.exists('settings.json') or refresh == True:
        for x in path:
            if not os.path.exists(x):
                os.mkdir(x)
            else:
                pass
            dirlinks = os.listdir(x)
            for i in dirlinks:
                if '.url' in i or '.lnk' in i:
                    links_.append(i)
                    if x == 'links\\sys': 
                        while i[-1] != '.':
                            i = i[:-1]
                        i = i[:-1]
                        sys_links.append(i.lower())
            if len(links_) == 0:
                if x != 'sys':
                    print(Fore.RED + "You don't have any shortcuts in the folder!\n")
                    input('Press enter')
                    Jsonwrite('settings.json', [[],{}])
                    time.sleep(1)
                    main()
            for i in links_:
                value = i
                while i[-1] != '.':
                    i = i[:-1]
                i = i[:-1]
                links.append(i)
                links_dict.setdefault(i, value)

            links = list(set(links))

        links_list = []
        for i in list(links_dict.keys()):
            if i.lower() not in sys_links:
                links_list.append(i)
        Jsonconfig(links_list, links_dict, sys_links)

        return [links_dict, sys_links]

def add_shortcut(path):                                 # Создание ярлыка
    os.system('cls||clear')
    print(tui.label('Adding new shortcut'))
    print(Fore.CYAN + 'Enter path to .exe file: ', end='')
    wDir = input()
    if os.path.exists(wDir):
        if os.path.isfile(wDir):
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

            path = os.path.join("{}\\{}.lnk".format(path, link))
            icon = target

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = wDir
            shortcut.IconLocation = icon
            shortcut.save()
            target = ''
            if 'sys' in path:
                for i in wDir:
                    while i[-1] != '.':
                        i = i[:-1]
                    i = i[:-1]
                    while i[-1] != '\\':
                        target += i[-1]
                        i = i[:-1]
            return target
        else:
            print(Fore.RED + "Wrong path!\npath must include .exe\n or Enter 'cancel' to cancel add..")
            time.sleep(1)
            add_shortcut(path)
    elif wDir in ['exit', 'cancel', 'stop', 'ex', 'x']:
        main()
    else:
        print(Fore.RED + "This path doesn't exists!\n or Enter 'cancel' to cancel add..")
        time.sleep(1)
        add_shortcut(path)

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
        Jsonconfig(links_list, links)
        time.sleep(1)
        main()
    else:
        print('Removing canceled..')
        time.sleep(1)

def update_links(links_list):                           # Обновление списка ярлыков
    links = load_links(['links', 'links\\sys'], True)
    new_links_list = list(links[0].keys())
    result = list(set(new_links_list) - set(links_list))
    for i in result:
        links_list.append(i)
    result = list(set(links_list) - set(new_links_list))
    for i in result:
        links_list.pop(links_list.index(i))
    result = None
    try:
        result = ''
        sys_links = []
        for i in os.listdir('links\\sys'):
            while i[-1] != '.':
                i = i[:-1]
            i = i[:-1]
            while i [-1] != '\\':
                result += i
                i = i[:-1]
            sys_links.append(result.lower())
    except:
        sys_links = links[1]
    Jsonconfig(links_list, links[0], sys_links)

def sort_links(links_list):                             # Перемещение ярлыков в списке
    from_ = input('from index: ')
    if from_ in ['exit', 'cancel', 'stop', 'ex', 'x']:
        main()
    to_ =  input('to index: ')
    if to_ in ['exit', 'cancel', 'stop', 'ex', 'x']:
        main()
    try:
        from_, to_ = int(from_), int(to_)
        link = links_list[from_]
        links_list.remove(link)
        links_list.insert(to_, link)
        sys_links = Jsonread('settings.json')[2]
        Jsonconfig(links_list, links, sys_links)
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

def run(app, path):                                     # Запуск приложения
    os.startfile(os.path.join(path, app))

def main():                                             # Интерфейс
    json_ = Jsonread('settings.json')
    links_list = []
    sys_links = json_[2]
    links = json_[1]
    for i in json_[0]:
        if i.lower() not in sys_links:
            links_list.append(i)

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
        run(links[links_list[cmd]], 'links')
        time.sleep(2)
        main()
    except ValueError:
        cmd = cmd.split()
        if cmd[0] in ['ex', 'exit', 'quit', 'x']:
            print('quiting...')
            time.sleep(.5)
            quit()
        elif cmd[0] in ['update', 'refresh', 'upd']:
            update_links(links_list)
            main()
        elif cmd[0] in ['remove', 'delete', 'rm']:
            remove_shortcut(links_list)
            main()
        elif cmd[0] in ['links', 'lnk', 'folder']:
            print('Opening {}...'.format(cmd[0]))
            os.system("explorer.exe {}".format(os.getcwd() + '\\links'))
            time.sleep(1)
            main()
        elif cmd[0] in ['sort', 'move']:
            sort_links(links_list)
            main()
        elif cmd[0] in ['add', 'create', 'make', 'mk']:
            try: 
                if cmd[1] == 'sys':
                    sys_links.append(add_shortcut('links\\sys'))
            except IndexError:
                add_shortcut('links')
            update_links(links_list)
            Jsonconfig(links_list, links[0], sys_links)
            main()
        elif cmd[0] in ['help']:
            os.system('cls||clear')
            command_list = [
                '"exit" ("quit", "ex", "x")\n\tUsing to quit the program',
                '"update" ("refresh", "upd")\n\tUsing to update shortcut list',
                '"remove" ("delete", "rm")\n\tUsing to delete app shortcut',
                '"add" ("create", "make", "mk")\n\tUsing to add new shortcut',
                '"add sys"\n\tUsing to add new sys shortcut',
                '"sort" ("move")\n\tUsing to move shortcut in list',
                '"rename"\n\tUsing to rename shortcuts in list',
                '"links" ("folder", "lnk")\n\tUsing to open shortcuts folder',
                '"sys"\n\tUsing to open sys shortcuts folder',
                '"shortcut name"\n\tUsing to start sys shortcut with "shortcutname"',
                ]
            print(tui.label('Help'))
            print(Fore.CYAN + tui.ul(command_list))
            input('Press enter to go to main menu')
            main()
        elif cmd[0] in ['rename']:
            rename_links(links, links_list)
            main()
        elif cmd[0] in ['sys']:
            os.system('cls||clear')
            print(tui.label('Sys apps'))
            print(tui.ul(sys_links))
            input()
            main()
        elif cmd[0] in sys_links:
            cmd = str(cmd[0])
            print('Starting {}...'.format(cmd))
            run(cmd, 'links\\sys')
            time.sleep(2)
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
    links = load_links(['links', 'links\\sys'], False)

    main() 
