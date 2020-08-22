import os
from win32com.client import Dispatch
import time
import json
from colorama import Fore, init
import easyTui as tui

def check_links(path, refresh):                         # Проверка наличия ярлыков
    '''
    Parsing all links in path folder
    and save into list
    '''
    if os.path.exists('settings.json') and refresh == False:
        links_dict = json_read('settings.json')[1]
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

def update_links(links_list):
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

def json_read(file):                                    # Чтение JSON
    with open(file, "r", encoding='utf-8') as read_file:
        data = json.load(read_file)
    return data

def json_write(file, data):                             # Запись JSON
    with open(file, 'w', encoding='utf-8') as write_file:
        write_file.write(json.dumps(data))

def json_config(links_list, links_dict):                # Сохранение настроек
    config_list = [links_list, links_dict]
    json_write('settings.json', config_list)

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
            update_links(links_list)
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
            os.system("explorer.exe {}".format(os.getcwd() + '\\links'))
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
        elif cmd.lower() in ['add', 'create', 'make']:
            add_shortcut()
            update_links(links_list)
            main()
        else:
            print('You must enter a Number..')
            time.sleep(1)
            main()

if __name__ == '__main__':
    init(autoreset=True)
    links = check_links('links', False)

    main()
