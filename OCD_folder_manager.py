import datetime
import os
import pythoncom
import re
import win32com.client

#local
from ocd_kw import destination_folders, phases

'''
To Do List:
    - read destination folder from excel
    - Edit destination folder folder_shape
    - read multi lever destination folder structures
'''
class Manager:
    def __init__(self):
        self.BASE_DIR = os.path.dirname(os.path.realpath(__file__))
        self.folder_to_fix = None
        self.destination_folders = destination_folders
        self.folder_shape = {
            'destination_folder':self.destination_folders,
            'date':'date',
            'type':'type',
            }

    def create_folder_shape(self):
        shape = ['date','file_type']
        for folder in self.destination_folders:
            shape.insert(0, folder)

        example_shape = ['elemets','drawing_type','date','file_type']
        pass


    def display_menu(self):
        print(' 1 - Enter name of folder to be fixed.')
        if self.folder_to_fix is not None:
            print(' 2 - Start reorganizing the "{}" folder.'.format(self.folder_to_fix))
        print(' 3 - To Open Destination Folders Menu.')
        print(' 4 - Display final folder shape')
        print(' 9 - Exit.')
        choice = input(': ')
        clear()
        self.manager_loop(choice)

    def destination_folders_menu(self):
        print('Edit Destination Folders.')
        print(' 1 - To remove destination folder.')
        print(' 2 - To add destination folder.')
        print(' 3 - To edit destination folder')
        print(' 9 - To exit to main menu.')
        choice = input(': ')
        choice = self.check_input(choice)
        clear()
        self.manager_loop((choice*10)) if choice != 0 else self.destination_folders()


    def remove_form_destination_folders(self):
        clear()
        print('List of folders to be removed.')
        for key in self.destination_folders:
            print(key, end=", ")
        to_remove = input('Type name of folder to be removed: ')
        try:
            self.destination_folders.pop(to_remove)

        except :
            print('Error, folder was not removed. Check the spelling and try again.')
            self.destination_folders_menu()
        clear()
        print('"{}" was removed from the list.'.format(to_remove))
        self.destination_folders_menu()

    def add_to_destination_folder(self):
        clear()
        to_add = input('Type name of folder you want to add: ')
        to_add = str(to_add)
        self.destination_folders[to_add] = {'keywords':[],'exclude':[]}
        print('"{}" was added to the list.'.format(to_add))
        #include
        print('Enter words that will include files in "{}" folder.'.format(to_add))
        print(' * leave a space between the words. Press Enter to leave this list empty.')
        kw = input('keywords:')
        kw = re.sub(r'\W+', '', kw)
        self.destination_folders[to_add]['keywords'] = kw
        print('keywords', kw)
        #exclude
        print('Enter words that will exclude files in "{}" folder.'.format(to_add))
        print(' * leave a space between the words. Press Enter to leave this list empty.')
        ex = input('exludes:')
        ex = re.sub(r'\W+', '', ex)
        self.destination_folders[to_add]['exclude'] = re.sub(r'\W+', '', ex)
        print('excludes: ', ex)
        clear()
        print( 'new folder added to the list: {}', self.destination_folders[to_add] )
        self.destination_folders_menu()

    def edit_destination_folder(self):
        pass

    def check_input(self, input):
        try:
            return int(input)
        except:
            print('Wrong command.')
        return 0

    def manager_loop(self, choice):
        while True:
            # MAIN MENU CHOICE
            if choice == 9 or choice == '9':
                return
            if choice == 3 or choice == '3':
                self.destination_folders_menu()
            if choice == 1 or choice == '1':
                self.define_folder_to_fix()
            if choice == 2 or choice == '2':
                self.fix_folder(self.folder_to_fix)
            if choice == 4 or choice == '4':
                print('folder structure: ', self.folder_shape)
                break
            # DESTINATION FOLDER MENU
            if choice == 10 or choice == '10':
                self.remove_form_destination_folders()
            if choice == 20:
                self.add_to_destination_folder()
            if choice == 30:
                edit_destination_folder(self)
            if choice == 90:
                break
        self.display_menu()

    def define_folder_to_fix(self):
                folder_to_fix = input('Enter folder name:')
                path_to_fix = os.path.join(self.BASE_DIR, folder_to_fix)
                if os.path.isdir(path_to_fix):
                    self.folder_to_fix = folder_to_fix
                    self.display_menu()
                else:
                    print('Can not find folder {} in here.'.format(folder_to_fix))
                    self.display_menu()



    #MAIN FUNCTION - FIX FOLDER
    def fix_folder(self, folder_to_fix=None):
        if folder_to_fix == None:
            print('Ups! something went wrong, Choose folder again or exit.')
            self.display_menu()
        fixed_folder = 'ocd_{}'.format(folder_to_fix)
        for (root,dirs,files) in os.walk(folder_to_fix, topdown=True):
            #clear()
            # LOOP THOUGH ALL FILES
            for file in files:
                file_path = os.path.join(self.BASE_DIR, root, file)
                filename, file_extension = os.path.splitext(file) # save extension an path to separated variables
                destination_folders_len = len(destination_folders) # nr of all folders that are specified in ocd_kw.py
                was_included = False # flag to determine if file goes to folder 'other'

                # check if file belongs to a folder form ocd_kw.py
                for iterator, folder in enumerate(destination_folders):
                    exclude_list = destination_folders[folder].get('exclude')
                    # CHECK FOR EXCLUDE WORDS
                    if len(exclude_list) > 0:
                        exclude_flag = self.check_if_list_contain(exclude_list, file_path)
                        if exclude_flag == 1:
                            continue

                    # CHECK FOR INCLUDE 'KEYWORDS' WORDS
                    include_list = destination_folders[folder].get('keywords')
                    include_flag = self.check_if_list_contain(include_list, file_path)
                    if include_flag == 1:
                        #save file with new path
                        was_included = True
                        self.create_shortcut(file_path, fixed_folder, folder, file=file)
                        continue


                    # SAVE FILE 'IN OTHER' FOLDER
                    if (iterator + 1) == destination_folders_len and not was_included:
                        #save file in Other folder
                        self.create_shortcut(file_path, fixed_folder, 'other', file=file)
                        include_flag = 0
                        continue
        print('done.')
        self.display_menu()

    # MAIN FUNCTION SUPPORT
    def check_if_list_contain(self,check_list=None, check_path=None):
        for item in check_list:
            if check_path.find(item) != -1:
                return 1 # path contain one of the keywords (or excludes) in the name
        return 0

    # MAIN FUNCTION SUPPORT
    def create_shortcut(self,file_path, fixed_folder, folder, file):
        filename, file_extension = os.path.splitext(file)
        iso_date = datetime.date.fromtimestamp(os.path.getmtime(file_path)) # get time of file last modification
        date_folder_name = '{}-{}'.format(iso_date.strftime("%y%d%m"), iso_date.strftime("%A")) # create name of date folder

        # create full path of new foler
        new_destination_folder = str(os.path.join(
            fixed_folder, folder,
            date_folder_name, file_extension[1:],).replace(
                os.sep, '/').replace(' ',''))
        # create new directory  for file to move
        os.makedirs(new_destination_folder, 0o666, exist_ok=True)
        # create shortcut
        shortcut_name = '{}.lnk'.format(filename)
        path = os.path.join(new_destination_folder, shortcut_name)
        target = os.path.join(file_path)

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WindowStyle = 1 # 7 - Minimized, 3 - Maximized, 1 - Normal
        shortcut.save()

# BASE_DIR = os.path.dirname(os.path.realpath(__file__))

clear = lambda: os.system('cls')


if __name__=='__main__':
    manager = Manager()
    manager.display_menu()
    exit()
