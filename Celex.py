# This is a sample Python script.

# Press Maiusc+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    import os.path
    import subprocess
    import sys
    import mmap, re
    import warnings
    import base64
    import sqlite3, pandas as pd

    import PySimpleGUI as sg
    import webbrowser
    import ctypes
    from threading import Thread

    class Closed:
        # Used to close other windows when main is closed
        isMainClosed = False
        isSettingsClosed = True

    def running_linux():
        return sys.platform.startswith('linux')


    def running_windows():
        return sys.platform.startswith('win')

    def get_file_list_dict():
        """
        Returns dictionary of files
        Key is short filename
        Value is the full filename and path
        :return: Dictionary of demo files
        :rtype: Dict[str:str]
        """

        demo_path = get_demo_path()
        demo_files_dict = {}
        for filenames in os.listdir(demo_path):
            if filenames.endswith('.xls') or filenames.endswith('.xlsx'):
                if filenames.startswith('.') or filenames.startswith('~'):
                    continue
                fname_full = demo_path + '/' + filenames
                if filenames not in demo_files_dict.keys():
                    demo_files_dict[filenames] = fname_full
                else:
                    # Allow up to 100 duplicated names. After that, give up
                    for i in range(1, 100):
                        new_filenames = f'{filenames}_{i}'
                        if new_filenames not in demo_files_dict:
                            demo_files_dict[new_filenames] = fname_full
                            break

        return demo_files_dict


    def get_file_list():
        """
        Returns list of filenames of files to display
        No path is shown, only the short filename
        :return: List of filenames
        :rtype: List[str]
        """
        return_value = []
        for _ in get_file_list_dict().keys():
            return_value.append(_ + ' (' + get_file_list_dict()[_] + ')')
        return return_value


    def get_demo_path():
        """
        Get the top-level folder path
        :return: Path to list of files using the user settings for this file.  Returns folder of this file if not found
        :rtype: str
        """
        demo_path = sg.user_settings_get_entry('-demos folder-', os.path.dirname(__file__))

        return demo_path


    def get_global_editor():
        """
        Get the path to the editor based on user settings or on PySimpleGUI's global settings
        :return: Path to the editor
        :rtype: str
        """
        try:  # in case running with old version of PySimpleGUI that doesn't have a global PSG settings path
            global_editor = sg.pysimplegui_user_settings.get('-editor program-')
        except:
            global_editor = ''
        return global_editor


    def get_editor():
        """
        Get the path to the editor based on user settings or on PySimpleGUI's global settings
        :return: Path to the editor
        :rtype: str
        """
        try:  # in case running with old version of PySimpleGUI that doesn't have a global PSG settings path
            global_editor = sg.pysimplegui_user_settings.get('-editor program-')
        except:
            global_editor = 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'
        user_editor = sg.user_settings_get_entry('-editor program-', 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')
        if user_editor == '':
            user_editor = global_editor

        return user_editor


    def using_local_editor():
        user_editor = sg.user_settings_get_entry('-editor program-', None)
        return get_editor() == user_editor


    def get_explorer():
        """
        Get the path to the file explorer program
        :return: Path to the file explorer EXE
        :rtype: str
        """
        try:  # in case running with old version of PySimpleGUI that doesn't have a global PSG settings path
            global_explorer = sg.pysimplegui_user_settings.get('-explorer program-', '')
        except:
            global_explorer = ''
        explorer = sg.user_settings_get_entry('-explorer program-', '')
        if explorer == '':
            explorer = global_explorer
        return explorer


    def advanced_mode():
        """
        Returns True is advanced GUI should be shown
        :return: True if user indicated wants the advanced GUI to be shown (set in the settings window)
        :rtype: bool
        """
        return sg.user_settings_get_entry('-advanced mode-', False)


    def get_theme():
        """
        Get the theme to use for the program
        Value is in this program's user settings. If none set, then use PySimpleGUI's global default theme
        :return: The theme
        :rtype: str
        """
        # First get the current global theme for PySimpleGUI to use if none has been set for this program
        try:
            global_theme = sg.theme_global()
        except:
            global_theme = sg.theme()
        # Get theme from user settings for this program.  Use global theme if no entry found
        user_theme = sg.user_settings_get_entry('-theme-', 'DarkGrey14')
        if user_theme == '':
            user_theme = global_theme
        return user_theme


    # We handle our code properly. But in case the user types in a flag, the flags are now in the middle of a regex. Ignore this warning.

    warnings.filterwarnings("ignore", category=DeprecationWarning)


    # New function
    def get_line_number(file_path, string, dupe_lines):
        lmn = 0
        with open(file_path, encoding="utf-8") as f:
            for num, line in enumerate(f, 1):
                if string.strip() == line.strip() and num not in dupe_lines:
                    lmn = num
        return lmn


    def kill_ascii(s):
        return "".join([x if ord(x) < 128 else '?' for x in s])

    """
    User defined functions for excel usage
    """
    class Celex:
        class Excel:
            def __init__(self, excel):
                self.excel = excel
            def filterByColumn(self, columns: list):
                """
                Takes an excel data frame string literal representation created by pandas and generates a new representation
                which uses the filter specified in the parameter columns, which needs to be a list
                :param excel:
                :param columns:
                :return string:
                """
                return self.excel[columns]

        class SqLite:
            def __init__(self, db):
                self.db = db

        def __init__(self, values):
                self.values = values
                self.inputBufferList = self.createInputBufferList()
            
        def checkKeyWord(self, index, entry, splitList: list, word, keyWord, typeToCheck, isLastOneUsed):
            """
            Used to check if the current word being checked contains the keyword, and if it is so
            it checks the last and next word in the current entry of the splitList (List of strings
            that need to be split. It also sets the isLastOneUsed bool, which is used to determine if the next word
            has already been checked in the last iteration in the loop
            :param index:
            :param entry:
            :param splitList:
            :param word:
            :param keyWord:
            :param typeToCheck:
            :param isLastOneUsed:
            :return [splitList, isLastOneUsed]:
            """
            # Check if the keyword is in the string to check
            return_value = splitList
            if not (keyWord in word):
                return [[], False]
            # If it is the first element check only the next one
            if index:
                # Check last element first, then next element
                # If successful assign the value to split list and pop both values
                if isinstance(entry.split(' ')[index - 1], typeToCheck):
                    splitList.append(entry.split(' ')[index - 1] + ' ' + entry.split(' ')[index])
                    isLastOneUsed = False
                elif isinstance(entry.split(' ')[index + 1], typeToCheck):
                    splitList.append(entry.split(' ')[index + 1] + ' ' + entry.split(' ')[index])
                    isLastOneUsed = True
            else:
                if isinstance(entry.split(' ')[index + 1], typeToCheck):
                    splitList.append(entry.split(' ')[index + 1] + ' ' + entry.split(' ')[index])
                    isLastOneUsed = True
            return [splitList, isLastOneUsed]

        def createInputBufferList(self):
            """
            Creates a buffer for the user input in the main ML element, then converts it into a list
            containing a line in each entry
            :return list:
            """
            return_value = []
            for line in self.values[ML_KEY].split('\n'):
                return_value.append(line)
            return return_value

        def ignoreComments(self):
            """
            Deletes all the lines marked as comments or empty in the input buffer
            :return:
            """
            index = 0
            return_value = self.inputBufferList
            for line in self.inputBufferList:
                index = self.inputBufferList.index(line)
                isEmpty = False
                if line.startswith('//') or line.startswith('\n') or line == '':
                    return_value.pop(index)

                # Loops through the characters to check is the line is empty
                # If it is, it will be ignored
                for char in line:
                    if char == ' ' or char == '\n':
                        isEmpty = True
                        continue
                    else:
                        isEmpty = False
                        break
                if isEmpty:
                    return_value.pop(index)
            return return_value

        def getKeyWords(self):
            """
            Gets the keywords to look for in multi-word cells
            :return keyword list:
            """
            return_value = []
            for line in self.values[ML_KEY]:
                list_words_line = line.split(' ')
                for word in list_words_line:
                    if word.startswith('$'):
                        return_value.append(word[1, len(word)])
            return return_value

        def getDemoListEntry(self):
            """
            Gets the path of the entry selected by the users
            :return entry_path:
            """
            full_filename = []
            for line in self.values['-DEMO LIST-']:
                full_filename, line = line, 1
                full_filename = full_filename.split(' ')[1]
                full_filename = full_filename[1:len(full_filename) - 1]
            return full_filename


        def readRuleList(self):
            """
            Reads the values defined on the right multiline element by the user which contain a '=' character and assigns them to a list
            :param values:
            :return rule_list:
            """
            return_value = ''
            for _ in self.inputBufferList:
                if _.startswith('//'):
                    continue
                return_value = return_value + _ + '\n'
            return return_value

        def readColumnList(self):
            """
            Reads the value on the right ML element which start with the character '/' and assigns them to a list
            :param values:
            :return column list filter:
            """
            col_List = self.values['-COLUMN FILTER-'].split(';')
            return_value = []
            for word in col_List:
                # Check for useless characters in the beginning or end of the string
                if word.startswith(' ') or word.startswith('\n') or word.startswith('\t'):
                    word = word[1, len(word)]
                if word.endswith(' ') or word.endswith('\n') or word.endswith('\t'):
                    word = word[1, len(word) - 1]
                return_value.append(word)
            return return_value

        def filterRuleList(self, rule_List):
            """
            Is used to filter out useless elements in a precedently created rule list
            :param rule_List:
            :return rule_list:
            """
            return_value = []
            for i in rule_List:
                if '=' in i:
                    return_value.append(i)
                else:
                    continue
            return return_value

        def listToDict(self, list):
            """
            Is used to convert a list to a dictionary. Useful to sort out the rules with lhs and rhs
            :param __list:
            :return dict:
            """
            dict = {}
            index = 0
            for i in list:
                bufferList = list[index].strip()
                bufferList = list[index].split('=')
                dict[bufferList[0]] = bufferList[1]
                index += 1
            return dict

        def getRuleListDict(self):
            """
            Creates a rule list, filters out useless entries and converts it to a dictionary
            :param values:
            :return dict:
            """
            rule_list = self.readRuleList().split('\n')
            rule_list = self.filterRuleList(rule_list)
            rule_list = self.listToDict(rule_list)
            return rule_list

        def getRowListSQL(self, excelPath):
            """
            Creates a cursor for the SQL table then fetches all the table. It will also pop the first element which is only the row index
            then return the list which contains a row in each entry
            :param tableName:
            :param tableToLoad:
            :param bufferExcelSQL:
            :return list of rows:
            """
            tableToLoad = 'Table1'
            # Creates an SQL database to hold the information read from the excel in memory
            bufferExcelSQL = sqlite3.connect(':memory:')

            if excelPath:
                pd.read_excel(excelPath).to_sql(name=tableToLoad, con=bufferExcelSQL)

            # Creates a cursor for the SQL table and gets every entry from Table1
            cur = bufferExcelSQL.cursor()
            cur.execute('SELECT * FROM Table1')
            list_row = []
            return_value = []
            for row in cur.fetchall():
                # Converts row tuple to list
                for _ in row:
                    list_row.append(_)
                # Pop the first element which is the row index
                list_row.pop(0)
                # Appends the last value to the return_value list, which is a list of rows in the excel file
                return_value.append(list_row)
            return return_value


    def find_in_file(string: str, celex):
        """
        Search through the demo files for a string.
        The case of the string and the file contents are ignored
        :param celex:
        :param string: String to search for
        :return: List of files containing the string
        :rtype: List[str]
        """
        file_list_dict = get_file_list_dict()
        dir_list = file_list_dict.values()

        return_value = []
        # For every path in the directories list
        for file_path in dir_list:
            # Creates a list where each entry is a row from the excel DF
            list_row = celex.getRowListSQL(file_path)
            # For each row in the list
            for row in list_row:
                # For each string or cell in the current row
                for cell in row:
                    if string in cell.lower():
                        # Appends the file name
                        return_value.append(file_path.split('/')[len(file_path.split('/')) - 1])
        return return_value

    def control_variables_window():
        """
        Shows the window used to specify which values to look for when sorting out the entries in
        'value keyword' or 'keyword value' strings
        :return True if variables have been changed:
        """


    def settings_window():
        """
        Show the settings window.
        This is where the folder paths and program paths are set.
        Returns True if settings were changed
        :return: True if settings were changed
        :rtype: (bool)
        """
        layout = [[sg.Text('Valori di controllo', font='DEFAULT 25')],
                  [sg.Text('Inserisci sotto i valori', font='_ 16')],
                  [sg.CB('Attivo',text_color='red'), sg.Input('Valore'), sg.Combo([], )]]

        try:
            global_editor = sg.pysimplegui_user_settings.get('-editor program-')
        except:
            global_editor = 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'
        try:
            global_explorer = sg.pysimplegui_user_settings.get('-explorer program-')
        except:
            global_explorer = 'C:/Windows/explorer.exe'
        try:  # in case running with old version of PySimpleGUI that doesn't have a global PSG settings path
            global_theme = sg.theme_global()
        except:
            global_theme = 'DarkGrey14'

        layout = [[sg.T('Impostazioni di Celex', font='DEFAULT 25')],
                  [sg.T('Percorso al file', font='_ 16')],
                  [sg.Combo(sorted(sg.user_settings_get_entry('-folder names-', [])),
                            default_value=sg.user_settings_get_entry('-demos folder-', get_demo_path()), size=(50, 1),
                            key='-FOLDERNAME-'),
                   sg.FolderBrowse('Esplora file', target='-FOLDERNAME-'), sg.B('Pulisci')],
                  [sg.T('Editor', font='_ 16')],
                  [sg.T('Lascia vuoto per usare quello di default'), sg.T('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')],
                  [sg.In(sg.user_settings_get_entry('-editor program-', 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'), k='-EDITOR PROGRAM-'), sg.FileBrowse()],
                  [sg.T('Esplora file', font='_ 16')],
                  [sg.T('Lascia vuoto per usare il default'), sg.T('C:/Windows/explorer.exe')],
                  [sg.In(sg.user_settings_get_entry('-explorer program-', 'C:/Windows/explorer.exe'), k='-EXPLORER PROGRAM-'), sg.FileBrowse()],
                  [sg.T('Tema', font='_ 16')],
                  [sg.T('Lascia vuoto per usare il default'), sg.T('DarkGrey14')],
                  [sg.Combo([''] + sg.theme_list(), sg.user_settings_get_entry('-theme-', 'DarkGrey14'), readonly=True,
                            k='-THEME-')],
                  [sg.T('Doppio click su un file:'),
                   sg.R('Avvia', 2, sg.user_settings_get_entry('-dclick runs-', False), k='-DCLICK RUNS-'),
                   sg.R('Modifica file', 2, sg.user_settings_get_entry('-dclick edits-', False), k='-DCLICK EDITS-'),
                   sg.R('Nulla', 2, sg.user_settings_get_entry('-dclick none-', False), k='-DCLICK NONE-')],
                  [sg.CB('Usa interfaccia avanzata', default=advanced_mode(), k='-ADVANCED MODE-')],
                  [sg.B('Ok', bind_return_key=True), sg.B('Cancella')],
                  ]

        window_settings = sg.Window('Impostazioni', layout, icon=icon_path)

        settings_changed = False

        while True:

            event, values = window_settings.read()
            # Used to close this window if main is closed
            if Closed.isMainClosed:
                window_settings.close()
                return settings_changed

            if event in ('Cancella', sg.WIN_CLOSED, sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
                break
            if event == 'Ok':
                sg.user_settings_set_entry('-demos folder-', values['-FOLDERNAME-'])
                sg.user_settings_set_entry('-editor program-', values['-EDITOR PROGRAM-'])
                sg.user_settings_set_entry('-theme-', values['-THEME-'])
                sg.user_settings_set_entry('-folder names-', list(
                    set(sg.user_settings_get_entry('-folder names-', []) + [values['-FOLDERNAME-'], ])))
                sg.user_settings_set_entry('-explorer program-', values['-EXPLORER PROGRAM-'])
                sg.user_settings_set_entry('-advanced mode-', values['-ADVANCED MODE-'])
                sg.user_settings_set_entry('-dclick runs-', values['-DCLICK RUNS-'])
                sg.user_settings_set_entry('-dclick edits-', values['-DCLICK EDITS-'])
                sg.user_settings_set_entry('-dclick nothing-', values['-DCLICK NONE-'])
                settings_changed = True
                break
            elif event == 'Pulisci':
                sg.user_settings_set_entry('-folder names-', [])
                sg.user_settings_set_entry('-last filename-', '')
                window_settings['-FOLDERNAME-'].update(values=[], value='')

            window_settings.close()
            return settings_changed

    ML_KEY = '-ML-'  # Multline's key

    def listOneToN(n):
        """
        Just a list of numbers from 1 to n
        """
        num = 0
        return_value = []
        for int in range(1, n):
            num = num + 1
            return_value.append(num)
        return return_value

    # --------------------------------- Create the window ---------------------------------
    def make_window():
        """
        Creates the main window
        :return: The main window object
        :rtype: (Window)
        """
        # Fixes the taskbar icon problem on windows
        if running_windows():
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u'Celex')

        theme = get_theme()
        if not theme:
            theme = sg.OFFICIAL_PYSIMPLEGUI_THEME
        sg.theme(theme)
        # First the window layout...2 columns

        find_tooltip = "Trova nella cartella\n" \
                       "Inserisci una stringa da cercare nella cartella"
        filter_tooltip = "Filtra file\n" \
                         "Inserisci una stringa per filtrare i nomi"
        find_re_tooltip = "Trova nella cartela con regex(Avanzato)\n" \
                          "Inserisci una REGEX(REGular EXpression) per filtrare i risultati"

        left_col = sg.Column([
            [sg.Listbox(values=get_file_list(), select_mode=sg.SELECT_MODE_EXTENDED, size=(50, 20),
                        bind_return_key=True, key='-DEMO LIST-')],
            [sg.Text('Filtro (F1):', tooltip=filter_tooltip),
             sg.Input(size=(25, 1), focus=True, enable_events=True, key='-FILTER-', tooltip=filter_tooltip),
             sg.T(size=(15, 1), k='-FILTER NUMBER-')],
            [sg.Button('Avvia (Tutti i file)'), sg.B('Modifica file'), sg.B('Pulisci'), sg.B('Apri in cartella')],
            [sg.Text('Trova (F2):', tooltip=find_tooltip),
             sg.Input(size=(25, 1), enable_events=True, key='-FIND-', tooltip=find_tooltip),
             sg.T(size=(15, 1), k='-FIND NUMBER-')],
        ], element_justification='l', expand_x=True, expand_y=True)

        lef_col_find_re = sg.pin(sg.Col([
            [sg.Text('Trova (F3):', tooltip=find_re_tooltip),
             sg.Input(size=(25, 1), key='-FIND RE-', tooltip=find_re_tooltip), sg.B('Trova REGEX')]], k='-RE COL-'))

        right_col = [
            [sg.Text('Parole da sostituire:', font='Default 10', pad=(0, 0), justification='left', grab=True)],
            [sg.Multiline(write_only=False, key=ML_KEY, tooltip='Le linee che iniziano con'
                                                                '"//" verranno ignorate, e sono considerate commenti'
                          , expand_y=True, expand_x=True, default_text=('//Inserisci qui la lista dei valori di controllo (Uno per linea).\n'
                                                                                            '//Esempio:\n'
                                                                                            'C45 = PC456T\n'
                                                                                            'C40 = PC40T67\n'))],
            [sg.Text('Colonne:', font='Default 10', pad=(0, 0), justification='left', tooltip='Colonne di interesse da usare come filtro')],
            [sg.Input('Es: Commessa; Pz; Misure Finite; Misure', key='-COLUMN FILTER-')],
            [sg.Text('Riga inizio tabella:', tooltip='Se insicuri lasciare valore di default'), sg.Combo(listOneToN(100), default_value=1, key='-START LINE-', readonly=True)],
            [sg.Button('Guida'), sg.B('Impostazioni'), sg.Button('Esci')],
            [sg.T('Progetto sviluppato con amore',
                  font='Default 8', pad=(0, 0))],
            [sg.T('Per quella grandissima troia di', font='Default 8', pad=(0, 0))],
            [sg.T('Alice', font='Default 8', pad=(0, 0))],
        ]

        options_at_bottom = sg.pin(sg.Column([[sg.CB('Mostra solo il primo risultato nel file', default=True, enable_events=True,
                                                     k='-FIRST MATCH ONLY-'),
                                               sg.CB('Ignora maiuscolo', default=True, enable_events=True,
                                                     k='-IGNORE CASE-'),
                                               sg.CB('Attendi completamento', default=False, enable_events=True,
                                                     k='-WAIT-', visible=False),
                                               ]],
                                             pad=(0, 0), k='-OPTIONS BOTTOM-', expand_x=True, expand_y=False),
                                   expand_x=True, expand_y=False)

        choose_folder_at_top = sg.pin(
            sg.Column([[sg.T('Clicca Impostazioni per cambiare la cartella selezionata'),
                        sg.Combo(sorted(sg.user_settings_get_entry('-folder names-', [])),
                                 default_value=sg.user_settings_get_entry('-demos folder-', ''),
                                 size=(50, 30), key='-FOLDERNAME-', enable_events=True, readonly=True),
                        ]], pad=(0, 0),
                      k='-FOLDER CHOOSE-'))
        # ----- Full layout -----

        layout = [[sg.Text('Celex', font=('Calibri', 30), text_color='#C4DFE6', pad=(10, 0))],
                  [choose_folder_at_top, sg.FolderBrowse('Esplora file', target='-FOLDERNAME-'), sg.B('Pulisci', key='-CLEAN FOLDERNAME-')],
                  # [sg.Column([[left_col],[ lef_col_find_re]], element_justification='l',  expand_x=True, expand_y=True), sg.Column(right_col, element_justification='c', expand_x=True, expand_y=True)],
                  [sg.Pane([sg.Column([[left_col], [lef_col_find_re]], element_justification='l', expand_x=True,
                                      expand_y=True),
                            sg.Column(right_col, element_justification='l', expand_x=True, expand_y=True,)],
                           orientation='h', relief=sg.RELIEF_SUNKEN, k='-PANE-', show_handle=False)],
                  [options_at_bottom], ]

        # --------------------------------- Create Window ---------------------------------
        window = sg.Window('Celex', layout, finalize=True, icon=icon_path, resizable=True,
                           use_default_focus=False, size=(1200, 800))
        window.set_min_size(window.size)
        window.bring_to_front()
        window['-DEMO LIST-'].expand(True, True, True)
        window[ML_KEY].expand(True, True, True)
        window['-PANE-'].expand(True, True, True)

        window.bind('<F1>', '-FOCUS FILTER-')
        window.bind('<F2>', '-FOCUS FIND-')
        window.bind('<F3>', '-FOCUS RE FIND-')
        window['-FOLDER CHOOSE-'].update(visible=True)

        if not advanced_mode():
            window['-RE COL-'].update(visible=False)
            window['-OPTIONS BOTTOM-'].update(visible=False)

        # sg.cprint_set_output_destination(window, ML_KEY)
        return window


    # --------------------------------- Main Program Layout ---------------------------------

    def main():
        """
        The main program that contains the event loop.
        It will call the make_window function to create the window.
        """

        find_in_file.file_list_dict = None

        old_typed_value = None

        file_list_dict = get_file_list_dict()
        file_list = get_file_list()
        window = make_window()
        window['-FILTER NUMBER-'].update(f'{len(file_list)} file')

        counter = 0
        while True:
            event, values = window.read()
            celex = Celex(values)

            counter += 1
            if event in (sg.WINDOW_CLOSED, 'Exit'):
                break
            if event == 'Guida':
                webbrowser.open('https://github.com/BomboBombone/celex', new=1)

            if event == '-DEMO LIST-':  # if double clicked (used the bind return key parm)
                if sg.user_settings_get_entry('-dclick runs-'):
                    event = 'Avvia (Tutti i file)'
                elif sg.user_settings_get_entry('-dclick edits-'):
                    event = 'Modifica file'
            if event == 'Modifica file':
                editor_program = get_editor()
                for file in values['-DEMO LIST-']:
                    # Takes the entry selected by the user and sets full_filename to the path of such file
                    full_filename, line = file, 1
                    full_filename = full_filename.split(' ')[1]
                    full_filename = full_filename[1:len( full_filename ) - 1]
                    if line is not None:
                        if using_local_editor():
                            execute_command_subprocess(editor_program, full_filename)
                        else:
                            try:
                                sg.execute_editor(full_filename, line_number=int(line))
                            except:
                                execute_command_subprocess(editor_program, full_filename)

            elif event == 'Avvia (Tutti i file)':
                celex_excel = Celex.Excel(celex.getDemoListEntry())

                celex.inputBufferList = celex.ignoreComments()

                # Get all the rules and column filters
                rule_list_dict = celex.getRuleListDict()
                columns_filter = celex.readColumnList()

                list_row = celex.getRowListSQL(celex.getDemoListEntry()[0])
                for row in list_row:
                    # Creates a to_separate list where each entry is a list of strings to separate in different cells
                    to_separate = []
                    for entry in row:
                        if ' ' in entry or '\n' in entry:
                            to_separate.append(entry)

                    # Creates a buffer list to loop through the entries in
                    # the to_separate list and performs a check for keywords
                    splitList = []
                    index = 0
                    isLastOneUsed = False
                    # For every string that needs to be split
                    for entry in to_separate:
                        # For every word in each of these strings
                        for word in entry.split(' '):
                            # For each keyword pulled from the ML element
                            for keyWord in celex.getKeyWords():
                                if isLastOneUsed:
                                    continue
                                for filter in columns_filter:
                                    splitList = celex.checkKeyWord(index, entry, splitList, word, keyWord, int, isLastOneUsed)[0]
                                    isLastOneUsed = celex.checkKeyWord(index, entry, splitList, word, keyWord, int, isLastOneUsed)[1]
                print(splitList)


            # elif event == 'Avvia':

            # elif event == 'Sostituisci':

            elif event == '-FILTER-':
                new_list = [i for i in file_list if values['-FILTER-'].lower() in i.lower()]
                window['-DEMO LIST-'].update(new_list)
                window['-FILTER NUMBER-'].update(f'{len(new_list)} file')
                window['-FIND NUMBER-'].update('')
                window['-FIND-'].update('')
                window['-FIND RE-'].update('')
            elif event == '-FOCUS FIND-':
                window['-FIND-'].set_focus()
            elif event == '-FOCUS FILTER-':
                window['-FILTER-'].set_focus()
            elif event == '-FOCUS RE FIND-':
                window['-FIND RE-'].set_focus()
            elif event == '-FIND-' or event == '-FIRST MATCH ONLY-' or event == '-FIND RE-':
                is_ignore_case = values['-IGNORE CASE-']
                old_ignore_case = False
                current_typed_value = str(values['-FIND-'])
                if len(values['-FIND-']) == 1:
                    window[ML_KEY].update('')
                if values['-FIND-']:
                    if find_in_file.file_list_dict is None or old_typed_value is None or old_ignore_case is not is_ignore_case:
                        # New search.
                        old_typed_value = current_typed_value
                        file_list = find_in_file(values['-FIND-'], celex)
                    elif current_typed_value.startswith(old_typed_value) and old_ignore_case is is_ignore_case:
                        old_typed_value = current_typed_value
                        file_list = find_in_file(values['-FIND-'], celex)
                    else:
                        old_typed_value = current_typed_value
                        file_list = find_in_file(values['-FIND-'], celex)
                    window['-DEMO LIST-'].update(sorted(file_list))
                    window['-FIND NUMBER-'].update(f'{len(file_list)} file')
                    window['-FILTER NUMBER-'].update('')
                    window['-FIND RE-'].update('')
                    window['-FILTER-'].update('')
                elif values['-FIND RE-']:
                    window['-ML-'].update('')
                    file_list = find_in_file(values['-FIND RE-'], celex)
                    window['-DEMO LIST-'].update(sorted(file_list))
                    window['-FIND NUMBER-'].update(f'{len(file_list)} file')
                    window['-FILTER NUMBER-'].update('')
                    window['-FIND-'].update('')
                    window['-FILTER-'].update('')
            elif event == 'Trova REGEX':
                window['-ML-'].update('')
                file_list = find_in_file(values['-FIND RE-'], celex)
                window['-DEMO LIST-'].update(sorted(file_list))
                window['-FIND NUMBER-'].update(f'{len(file_list)} file')
                window['-FILTER NUMBER-'].update('')
                window['-FIND-'].update('')
                window['-FILTER-'].update('')
            elif event == 'Impostazioni':
                if settings_window() is True:
                    window.close()
                    window = make_window()
                    file_list_dict = get_file_list_dict()
                    file_list = get_file_list()
                    window['-FILTER NUMBER-'].update(f'{len(file_list)} file')
            elif event == '-CLEAN FOLDERNAME-':
                file_list = get_file_list()
                window['-FOLDERNAME-'].update('')
                window['-FILTER-'].update('')
                window['-FILTER NUMBER-'].update(f'{len(file_list)} file')
                window['-FIND-'].update('')
                window['-DEMO LIST-'].update(file_list)
                window['-FIND NUMBER-'].update('')
                window['-FIND RE-'].update('')
            elif event == '-FOLDERNAME-':
                sg.user_settings_set_entry('-demos folder-', values['-FOLDERNAME-'])
                file_list_dict = get_file_list_dict()
                file_list = get_file_list()
                window['-DEMO LIST-'].update(values=file_list)
                window['-FILTER NUMBER-'].update(f'{len(file_list)} file')
                window['-ML-'].update('')
                window['-FIND NUMBER-'].update('')
                window['-FIND-'].update('')
                window['-FIND RE-'].update('')
                window['-FILTER-'].update('')
            elif event == 'Apri in cartella':
                explorer_program = get_explorer()
                if explorer_program:
                    for file in values['-DEMO LIST-']:
                        file_selected = str(file_list_dict[file])
                        file_path = os.path.dirname(file_selected)
                        if running_windows():
                            file_path = file_path.replace('/', '\\')
                        execute_command_subprocess(explorer_program, file_path)

        window.close()

    def execute_py_file_with_pipe_output(pyfile, parms=None, cwd=None, interpreter_command=None, wait=False,
                                         pipe_output=False):
        """
        Executes a Python file.
        The interpreter to use is chosen based on this priority order:
            1. interpreter_command paramter
            2. global setting "-python command-"
            3. the interpreter running running PySimpleGUI
        :param pyfile: the file to run
        :type pyfile: (str)
        :param parms: parameters to pass on the command line
        :type parms: (str)
        :param cwd: the working directory to use
        :type cwd: (str)
        :param interpreter_command: the command used to invoke the Python interpreter
        :type interpreter_command: (str)
        :param wait: the working directory to use
        :type wait: (bool)
        :param pipe_output: If True then output from the subprocess will be piped. You MUST empty the pipe by calling execute_get_results or your subprocess will block until no longer full
        :type pipe_output: (bool)
        :return: Popen object
        :rtype: (subprocess.Popen) | None
        """

        if pyfile[0] != '"' and ' ' in pyfile:
            pyfile = '"' + pyfile + '"'
        try:
            if interpreter_command is not None:
                python_program = interpreter_command
            else:
                python_program = sg.pysimplegui_user_settings.get('-python command-', '')
        except:
            python_program = ''

        if python_program == '':
            python_program = 'python' if sys.platform.startswith('win') else 'python3'
        if parms is not None and python_program:
            sp = execute_command_subprocess_with_pipe_output(python_program, pyfile, parms, wait=wait, cwd=cwd,
                                                             pipe_output=pipe_output)
        elif python_program:
            sp = execute_command_subprocess_with_pipe_output(python_program, pyfile, wait=wait, cwd=cwd,
                                                             pipe_output=pipe_output)
        else:
            print('execute_py_file - No interpreter has been configured')
            sp = None
        return sp


    def execute_command_subprocess_with_pipe_output(command, *args, wait=False, cwd=None, pipe_output=False):
        """
        Runs the specified command as a subprocess.
        By default the call is non-blocking.
        The function will immediately return without waiting for the process to complete running. You can use the returned Popen object to communicate with the subprocess and get the results.
        Returns a subprocess Popen object.
        :param command: Filename to load settings from (and save to in the future)
        :type command: (str)
        :param *args:  Variable number of arguments that are passed to the program being started as command line parms
        :type *args: (Any)
        :param wait: If True then wait for the subprocess to finish
        :type wait: (bool)
        :param cwd: Working directory to use when executing the subprocess
        :type cwd: (str))
        :param pipe_output: If True then output from the subprocess will be piped. You MUST empty the pipe by calling execute_get_results or your subprocess will block until no longer full
        :type pipe_output: (bool)
        :return: Popen object
        :rtype: (subprocess.Popen)
        """
        try:
            if args is not None:
                expanded_args = ' '.join(args)
                # print('executing subprocess command:',command, 'args:',expanded_args)
                if command[0] != '"' and ' ' in command:
                    command = '"' + command + '"'
                # print('calling popen with:', command +' '+ expanded_args)
                # sp = subprocess.Popen(command +' '+ expanded_args, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, cwd=cwd)
                if pipe_output:
                    sp = subprocess.Popen(command + ' ' + expanded_args, shell=True, stdout=subprocess.PIPE,
                                          stderr=subprocess.PIPE, cwd=cwd)
                else:
                    sp = subprocess.Popen(command + ' ' + expanded_args, shell=True, stdout=subprocess.DEVNULL,
                                          stderr=subprocess.DEVNULL, cwd=cwd)
            else:
                sp = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=cwd)
            if wait:
                out, err = sp.communicate()
                if out:
                    print(out.decode("utf-8"))
                if err:
                    print(err.decode("utf-8"))
        except Exception as e:
            print('** Error executing subprocess **', 'Command:', command)
            print('error:', e)
            sp = None
        return sp


    def execute_py_get_interpreter():
        """
        Returns the command that was specified in the global options that will be used to execute Python files
        when the execute_py_file function is called.
        :return: Full path to python interpreter or '' if nothing entered
        :rtype: (str)
        """

        return sg.pysimplegui_user_settings.get('-python command-', '')


    # Normally you want to use the PySimpleGUI version of these functions
    try:
        execute_py_file = sg.execute_py_file
    except:
        execute_py_file = execute_py_file_with_pipe_output

    try:
        execute_py_get_interpreter = sg.execute_py_get_interpreter
    except:
        execute_py_get_interpreter = execute_py_get_interpreter

    try:
        execute_command_subprocess = sg.execute_command_subprocess
    except:
        execute_command_subprocess = execute_command_subprocess_with_pipe_output

    if __name__ == '__main__':
        icon_path = os.path.dirname(os.path.abspath(__file__)) + '\icon.ico'

    if __name__ == '__main__':
        try:
            version = sg.version
            version_parts = version.split('.')
            major_version, minor_version = int(version_parts[0]), int(version_parts[1])
            if major_version < 4 or minor_version < 32:
                sg.popup('Warning - Your PySimpleGUI version is less then 4.35.0',
                         'As a result, you will not be able to use the EDIT features of this program',
                         'Please upgrade to at least 4.35.0',
                         f'You are currently running version:',
                         sg.version,
                         background_color='red', text_color='white')
        except Exception as e:
            print(f'** Warning Exception parsing version: {version} **  ', f'{e}')
        main()