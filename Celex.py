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
        Only the path to the file
        :return: List of filenames
        :rtype: List[str]
        """
        return_value = []
        for _ in get_file_list_dict().keys():
            return_value.append(_ + ' (' + get_file_list_dict()[_] + ')')
        return return_value


    def get_path_list():
        """
        Returns a list of the paths of the files on the left column in main window
        """
        entry_list = get_file_list()
        return_value = []
        for entry in entry_list:
            return_value_buff = entry.split(' ')[1]
            return_value_buff = return_value_buff[1:len(return_value_buff) - 1]
            return_value.append(return_value_buff)
        return return_value


    def get_demo_path():
        """
        Get the top-level folder path
        :return: Path to list of files using the user settings for this file.  Returns folder of this file if not found
        :rtype: str
        """

        demo_path = sg.user_settings_get_entry('-demos folder-', os.path.dirname(__file__))
        if demo_path == 'C' or demo_path == 'C:':
            demo_path = 'C:/'
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
        user_editor = sg.user_settings_get_entry('-editor program-',
                                                 'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')
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

        class SqLite:
            def __init__(self, db):
                self.db = db

        def __init__(self, values):
            self.values = values
            self.inputBufferList = self.createInputBufferList()

        def checkKeyWord(self, index, entry, word, keyWord, typeToCheck, isLastOneUsed):
            """
            Used to check if the current word being checked contains the keyword, and if it is so
            it checks the last and next word in the current entry of the splitList (List of strings
            that need to be split). It also sets the isLastOneUsed bool, which is used to determine if the next word
            has already been checked in the last iteration in the loop
            :param index:
            :param entry:
            :param splitList:
            :param word:
            :param keyWord:
            :param typeToCheck:
            :param isLastOneUsed:
            :return [string, isLastOneUsed]:
            """
            return_value = ''
            # Convert type
            if typeToCheck == 'Stringa':
                typeToCheck = str
            else:
                typeToCheck = int
            # Check if the keyword is in the string to check
            if not (keyWord in word):
                return [[], False]
            # If it is the first element check only the next one
            if index:
                try:
                    # Check last element first, then next element
                    # If successful assign the value to split list and pop both values
                    if isinstance(entry.split(' ')[index - 1], typeToCheck):
                        return_value = (entry.split(' ')[index - 1] + ' ' + entry.split(' ')[index])
                        isLastOneUsed = False
                    elif isinstance(entry.split(' ')[index + 1], typeToCheck):
                        return_value = (entry.split(' ')[index + 1] + ' ' + entry.split(' ')[index])
                        isLastOneUsed = True
                except IndexError:
                    pass
            else:
                try:
                    if isinstance(entry.split(' ')[index + 1], typeToCheck):
                        return_value = (entry.split(' ')[index + 1] + ' ' + entry.split(' ')[index])
                        isLastOneUsed = True
                except IndexError:
                    pass
            return [return_value, isLastOneUsed]

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
            return sg.user_settings_get_entry('-cv input-')

        def getDemoListEntry(self):
            """
            Gets the path of the entry selected by the users
            :return entry_path:
            """
            full_filename = []
            for line in self.values['-DEMO LIST-']:
                line = line.split(' ')[1]
                line = line[1:len(full_filename) - 1]
                full_filename.append(line)
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
                    word = word[1:len(word)]
                if word.endswith(' ') or word.endswith('\n') or word.endswith('\t'):
                    word = word[1:len(word) - 1]
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
                bufferList = list[index].split('=')
                dict[bufferList[0].strip()] = bufferList[1].strip()
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

        def getRowListDF(self, df, removeIndex: bool):
            """
            Fetches the row list from a df object
            """
            tableToLoad = 'Table1'
            # Creates an SQL database to hold the information read from the excel in memory
            bufferExcelSQL = sqlite3.connect(':memory:')

            df.to_sql(name=tableToLoad, con=bufferExcelSQL)

            # Creates a cursor for the SQL table and gets every entry from Table1
            cur = bufferExcelSQL.cursor()
            cur.execute('SELECT * FROM Table1')
            return_value = []
            for row in cur.fetchall():
                list_row = []
                # Converts row tuple to list
                for _ in row:
                    list_row.append(_)
                if removeIndex:
                    # Pop the first element which is the row index
                    list_row.pop(0)
                # Appends the last value to the return_value list, which is a list of rows in the excel file
                return_value.append(list_row)
            return return_value

        def checkOutPutFolder(self, folder_path):
            """
            Function used to check if there exists a destination folder for the output files, and if not, it creates it
            """
            if not os.path.exists(folder_path):
                os.mkdir(folder_path)

        def getMissingColumns(self, excel):
            return_value = []
            for column in str(sg.user_settings_get_entry('-col filter-')).split(';'):
                try:
                    buffDF = excel[[column]]
                except KeyError:
                    if column.startswith(' '):
                        column = column[1:len(column)]
                    if column.endswith(' '):
                        column = column[0:len(column) - 1]
                    return_value.append(column)
            return return_value


    class Excel(Celex):
        def __init__(self, excel, values):
            super().__init__(values)
            # The excel here is a list of excel file paths
            self.excel = excel

        def getMaterialsDict(self, df_list, materials_list):
            """
            Returns a dictionary with row_number and material
            """
            return_value = {}
            for df in df_list:
                list_row = self.getRowListDF(df, False)
                for row in list_row:
                    row_index = row[0]
                    return_value[row_index] = ''
                    for cell in row:
                        if row.index(cell) == 1:
                            continue
                        for string in str(cell).split(' '):
                            for material in materials_list:
                                if material in string:
                                    return_value[row_index] = material
                    if not return_value[row_index]:
                        return_value[row_index] = ''
            return return_value

        def removeSpaces(self, ex_list, removeLines=True):
            """
            Returns a list of all df objects dropped of all the unnamed columns and empty cells
            """
            return_list = []
            excel_list = ex_list
            # Removes blank columns
            for df in excel_list:
                to_drop = []
                if removeLines:
                    df = self.removeRows(sg.user_settings_get_entry('-start line-'), df)
                for column in df.columns:
                    if str(column).startswith('Unnamed:'):
                        to_drop.append(column)
                df = df.drop(columns=to_drop)
                return_list.append(df)
            # Drops all blank rows
            for df in return_list:
                df_index = return_list.index(df)
                df = df.dropna(how='all')
                list_row = self.getRowListDF(df, False)
                # Rows to drop that don't have the same number of values as the valid columns
                to_drop = []
                row_index = sg.user_settings_get_entry('-start line-')
                valid_values = 0
                try:
                    for cell in list_row[0]:
                        if cell:
                            valid_values += 1
                except IndexError:
                    pass
                for row in list_row:
                    to_drop_values = 0
                    for cell in row:
                        if cell:
                            to_drop_values += 1
                    if to_drop_values == valid_values:
                        continue
                    else:
                        to_drop.append(row[0])
                    row_index += 1
                df = df.drop(to_drop)
                return_list[df_index] = df

            return return_list

        def separateMeasures(self, ex_list):
            """
            Returns a dictionary containing the measures separated
            """
            return_dict = {}
            return_dict['Spessore'] = []
            return_dict['Larghezza'] = []
            return_dict['Lunghezza'] = []
            if self.values['-SPLIT MEASURES-']:
                excel_list = ex_list
                for df in excel_list:
                    for row in self.getRowListDF(df, True):
                        for cell in row:
                            if not cell:
                                continue
                            if ' ' not in str(cell):
                                continue
                            cell_list = str(cell).split(' ')
                            measures_found = False
                            for word in cell_list:
                                if measures_found:
                                    break
                                word_index = cell_list.index(word)
                                if 'x' not in word and 'Ø' not in word and 'L=' not in word:
                                    continue
                                measure_list = []
                                if word == 'x':
                                    measure_list = [cell_list[word_index - 1], cell_list[word_index + 1]]
                                    if cell_list[word_index + 2] == 'x':
                                        measure_list.append(cell_list[word_index + 3])
                                    measures_found = True
                                elif 'x' in word:
                                    measure_list = word.split('x')

                                for measure in measure_list:
                                    i = measure_list.index(measure)
                                    if measure.startswith('Ø'):
                                        measure = measure[1:len(measure)]
                                    elif measure.startswith('L='):
                                        measure = measure[2:len(measure)]
                                    measure_list[i] = measure

                                if measure_list:
                                    if len(measure_list) == 2:
                                        return_dict['Lunghezza'].append(measure_list[0])
                                        return_dict['Larghezza'].append(measure_list[1])
                                        return_dict['Spessore'].append('')
                                    elif len(measure_list) == 3:
                                        return_dict['Lunghezza'].append(measure_list[0])
                                        return_dict['Larghezza'].append(measure_list[1])
                                        return_dict['Spessore'].append(measure_list[2])
                                elif 'Ø' in word:
                                    return_dict['Lunghezza'].append(word[1:len(word)])
                                    return_dict['Spessore'].append('')
                                elif 'L=' in word:
                                    return_dict['Larghezza'].append(word[2:len(word)])
            return return_dict

        def removeRows(self, row_number, df):
            """
            Used to remove row_number-1 of rows from the top of the specified df object
            """
            list_rows = self.getRowListDF(df, True)
            df_length = self.getNofRows(df)
            # Gets a list from row_number to df_length
            listNtoLen = []
            if row_number < 2:
                row_number = 2
            for i in range(1, df_length + 1):
                if i >= row_number:
                    listNtoLen.append(i - 1)
            df = df.filter(items=listNtoLen, axis=0)

            # Changes the column names
            col_dict = {}
            original_cols = []
            new_cols = []
            for column in df.columns:
                original_cols.append(column)
            index = 0
            for row in list_rows:
                if index == row_number - 2:
                    i = 0
                    for column in row:
                        if column is None:
                            column = 'Unnamed: ' + str(i)
                        new_cols.append(column)
                        i += 1
                index += 1
            for column in original_cols:
                i = original_cols.index(column)
                try:
                    col_dict[column] = new_cols[i]
                except IndexError:
                    col_dict[column] = 'Unnamed ' + str(i)
            df = df.rename(columns=col_dict)
            # Returns a df containing all rows from row_number to the end
            return df

        def getExcelList(self):
            """
            Used to get list of df objs
            """
            excel_list = []
            for file_path in self.excel:
                excel_list.append(pd.read_excel(file_path))
            return excel_list

        def filterByColumn(self, columns: list, excel_list: list):
            """
            Takes an excel data frame created by pandas and generates a new representation
            which uses the filter specified in the parameter columns, which needs to be a list.
            Returns a list containing a list of df objs and a missing cols dict
            :param values:
            :param columns:
            :return [list, dict]:
            """
            missingColumns = {}
            for entry in get_file_list_dict().values():
                missingColumns[entry] = []
            return_value = []
            file_columns = excel_list[0].columns

            # For every excel file that has been selected
            index = 0
            for excel in excel_list:
                # Loop to check if every column in the column list exists in the source excel
                columns_buff = []
                for column in columns:
                    # If the column is not in the source file pop it
                    if not (column in file_columns):
                        # If the create missing checkbox is active
                        if self.values['-CREATE MISSING-']:
                            missingColumns[self.excel[index]].append(column)
                    else:
                        columns_buff.append(column)
                # After popping out every not existing column it appends the filtered excel file
                if columns_buff:
                    return_value.append(excel[columns_buff])
                index += 1
            return [return_value, missingColumns]

        def saveToFile(self, file_name: str, df):
            """
            Used to save to disk the DF obj into an excel file inside the user_specified folder
            """
            input_folder = self.values['-FOLDERNAME IN-']
            output_folder = sg.user_settings_get_entry('-output folder-') + '/'
            self.checkOutPutFolder(output_folder)

            for file_index in list1ToN(20):
                if os.path.exists(output_folder + file_name + '_Modificato' + '(' + str(file_index) + ').xlsx'):
                    continue
                else:
                    df.to_excel(output_folder + file_name + '_Modificato' + '(' + str(file_index) + ').xlsx',
                                sheet_name='Foglio1',
                                index=False)
                    break

        def getStartLine(self, excel):
            """
            Used to get the first useful line of the file, therefore excluding all the headers
            """
            pass

        def createColumns(self, columns: list, to_insert: str, df):
            """
            Used to create the specified columns to the corresponding df object.
            Can also specify a to_insert list which will be the values to assign to the missing columns
            Returns the df with the created columns
            """
            # Setup of the list
            to_insert_list = []
            for entry in to_insert.split(';'):
                # Remove white spaces from the beginning and end of the string
                if entry.startswith(' '):
                    entry = entry[1:len(entry)]
                if entry.endswith(' '):
                    entry = entry[0:len(entry) - 1]
                to_insert_list.append(entry)
            # Initial setup of the dictionary
            df_buff = {}
            columns_buff = []
            for column in columns:
                column = column[0:len(column) - 1]
                columns_buff.append(column)
            columns = columns_buff
            # Append the corresponding element to each column
            for column in columns:
                i = columns.index(column)
                df_buff[column] = []
                # Loop for the number of rows in the df object
                for index in list1ToN(self.getNofRows(df)):
                    df_buff[column].append(to_insert_list[i])
            return pd.DataFrame(df_buff)

        def getNofRows(self, df):
            """
            Used to get the number of rows of a specified df obj
            """
            df_length = 0
            df_rows = self.getRowListDF(df, True)
            for row in df_rows:
                df_length += 1
            if not df_length:
                df_length = 1
            return df_length

        def joinDF(self, excel_list_dest, excel_list_source):
            """
            Used to join two df obj lists
            """
            return_value = excel_list_dest
            if excel_list_dest and excel_list_source:
                for excel in excel_list_dest:
                    index = excel_list_dest.index(excel)
                    if excel_list_source[index]:
                        return_value.append(excel.join(excel_list_source[index]))
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


    def insertEntryRow():
        """
        Changes settings to fit the new layout
        """
        # Gets the old settings
        cb_value_list = sg.user_settings_get_entry('-cb value-')
        cv_input_list = sg.user_settings_get_entry('-cv input-')
        cv_type_list = sg.user_settings_get_entry('-cv type-')
        cv_col_list = sg.user_settings_get_entry('-cv col entry-')
        # Sets the new settings
        cb_value_list.append(True)
        cv_input_list.append('')
        cv_type_list.append('Stringa')
        cv_col_list.append('')
        sg.user_settings_set_entry('-cb value-', cb_value_list)
        sg.user_settings_set_entry('-cv input-', cv_input_list)
        sg.user_settings_set_entry('-cv type-', cv_type_list)
        sg.user_settings_set_entry('-cv col entry-', cv_col_list)


    def createEntryRow(index, layout):
        """
        Used to create an entry in the specified layout
        :returns layout:
        """
        if sg.user_settings_get_entry('-cv col combo-'):
            sg.user_settings_get_entry('-cv col entry-', sg.user_settings_get_entry('-cv col combo-')[0])
        else:
            sg.user_settings_set_entry('-cv col combo-', [''])
            sg.user_settings_set_entry('-cv col entry-', sg.user_settings_get_entry('-cv col combo-'))

        entry_row = [sg.CB('Attivo', text_color='red', key='-CB ' + str(index) + '-',
                           default=sg.user_settings_get_entry('-cb value-')[index]),
                     sg.Input(sg.user_settings_get_entry('-cv input-')[index],
                              key='-CV INPUT ' + str(index) + '-',
                              tooltip='Valore di controllo'),
                     sg.Combo(['Stringa', 'Numero'], default_value=sg.user_settings_get_entry('-cv type-')[index],
                              key='-CV TYPE ' + str(index) + '-',
                              tooltip='Stringa = Lettere, '
                                      'numeri e caratteri '
                                      'vari\nNumero = solo '
                                      'numeri'),
                     sg.Combo(sg.user_settings_get_entry('-cv col combo-'),
                              default_value=sg.user_settings_get_entry('-cv col entry-')[index],
                              key='-CV COL ' + str(index) + '-',
                              tooltip='Specifica la colonna nella quale devono essere assegnati questi valori',
                              readonly=True)]
        layout.append(entry_row)
        return layout


    def removeEntry():
        """
        Changes settings to fit new layout
        """
        # Gets the old settings
        cb_value_list = sg.user_settings_get_entry('-cb value-')
        cv_input_list = sg.user_settings_get_entry('-cv input-')
        cv_type_list = sg.user_settings_get_entry('-cv type-')
        cv_col_list = sg.user_settings_get_entry('-cv col entry-')
        # Sets the new settings
        if len(cb_value_list) > 1:
            cb_value_list.pop(len(cb_value_list) - 1)
        if len(cv_input_list) > 1:
            cv_input_list.pop(len(cv_input_list) - 1)
        if len(cv_type_list) > 1:
            cv_type_list.pop(len(cv_type_list) - 1)
        if len(cv_col_list) > 1:
            cv_col_list.pop(len(cv_col_list) - 1)
        sg.user_settings_set_entry('-cb value-', cb_value_list)
        sg.user_settings_set_entry('-cv input-', cv_input_list)
        sg.user_settings_set_entry('-cv type-', cv_type_list)
        sg.user_settings_set_entry('-cv col entry-', cv_col_list)


    def get_num_of_entries():
        """
        Gets the number of entries in the control interface GUI
        """
        try:
            number_of_entries = len(sg.user_settings_get_entry('-cb value-'))
        except TypeError:
            number_of_entries = 0
        return number_of_entries


    def createBaseLayout():
        """
        Used to create the base layout for the control window GUI
        """
        layout = [[sg.Text('Valori di controllo', font='DEFAULT 25')],
                  [sg.Text('Leggi la guida per maggiori informazioni', font='_ 14')],
                  [sg.Button('Aggiungi valore', enable_events=True, key='-ADD ENTRY-'),
                   sg.Button('Rimuovi valore', enable_events=True, key='-REMOVE ENTRY-')]]
        entries_num = get_num_of_entries()
        # If there is less than 2 entries, the remove value button will be removed fom the GUI
        if entries_num == 1:
            layout[2][1] = sg.Button('Rimuovi valore', enable_events=True, key='-REMOVE ENTRY-', visible=False)
        for i in list1ToN(entries_num):
            layout = createEntryRow(i, layout)
        return layout


    def make_control_window(create_entry: bool, del_entry: bool):
        """
        Used to create a window element
        """
        if create_entry:
            insertEntryRow()
        elif del_entry:
            removeEntry()

        layout = createBaseLayout()

        layout_final = [layout,
                        [sg.Column([[sg.Button('Ok'), sg.Button('Cancella')]], justification='right')]
                        ]

        window_control = sg.Window('Valori di controllo', layout_final, icon=icon_path)

        return window_control


    def saveControlSettings(values):
        """
        Used to save the current settings for the control window
        """
        cb_value_list = []
        cv_input_list = []
        cv_type_list = []
        cv_col_list = []
        # Append all the values to buffer lists
        for i in list1ToN(get_num_of_entries()):
            cb_value_list.append(values['-CB ' + str(i) + '-'])
            cv_input_list.append(values['-CV INPUT ' + str(i) + '-'])
            cv_type_list.append(values['-CV TYPE ' + str(i) + '-'])
            cv_col_list.append(values['-CV COL ' + str(i) + '-'])
        # Set the values to settings
        sg.user_settings_set_entry('-cb value-', cb_value_list)
        sg.user_settings_set_entry('-cv input-', cv_input_list)
        sg.user_settings_set_entry('-cv type-', cv_type_list)
        sg.user_settings_set_entry('-cv col entry-', cv_col_list)


    def control_variables_window():
        """
        Shows the window used to specify which values to look for when sorting out the entries in
        'value keyword' or 'keyword value' strings
        :return True if variables have been changed:
        """
        # Add one entry as default if you do not already have one
        if get_num_of_entries():
            window_control = make_control_window(False, False)
        else:
            window_control = make_control_window(True, False)

        while True:
            event, values = window_control.read()

            if event:
                saveControlSettings(values)
            if event in ('Cancella', sg.WIN_CLOSED, sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
                break
            if event == '-ADD ENTRY-':
                window_control.close()
                window_control = make_control_window(True, False)
                continue
            if event == '-REMOVE ENTRY-':
                window_control.close()
                window_control = make_control_window(False, True)
                continue
            if event == 'Ok':
                window_control.close()
                return True
        window_control.close()
        return False


    def settings_window():
        """
        Show the settings window.
        This is where the folder paths and program paths are set.
        Returns True if settings were changed
        :return: True if settings were changed
        :rtype: (bool)
        """
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
                  [sg.T('Percorso di output', font='_ 16', tooltip='Cartella di destinazione dei file')],
                  [sg.Combo(sorted(sg.user_settings_get_entry('-folder names o-', [])),
                            default_value=sg.user_settings_get_entry('-output folder-'), size=(50, 1),
                            key='-FOLDERNAME-'),
                   sg.FolderBrowse('Esplora file', target='-FOLDERNAME-'), sg.B('Pulisci')],
                  [sg.T('Editor', font='_ 16')],
                  [sg.T('Lascia vuoto per usare quello di default'),
                   sg.T('C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')],
                  [sg.In(sg.user_settings_get_entry('-editor program-',
                                                    'C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'),
                         k='-EDITOR PROGRAM-'), sg.FileBrowse()],
                  [sg.T('Esplora file', font='_ 16')],
                  [sg.T('Lascia vuoto per usare il default'), sg.T('C:/Windows/explorer.exe')],
                  [sg.In(sg.user_settings_get_entry('-explorer program-', 'C:/Windows/explorer.exe'),
                         k='-EXPLORER PROGRAM-'), sg.FileBrowse()],
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

            if event in ('Cancella', sg.WIN_CLOSED, sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
                break
            if event == 'Ok':
                sg.user_settings_set_entry('-output folder-', values['-FOLDERNAME-'])
                sg.user_settings_set_entry('-editor program-', values['-EDITOR PROGRAM-'])
                sg.user_settings_set_entry('-theme-', values['-THEME-'])
                sg.user_settings_set_entry('-folder names o-', list(
                    set(sg.user_settings_get_entry('-folder names o-', []) + [values['-FOLDERNAME-'], ])))
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
                window_settings['-FOLDERNAME-'].update(values=[])

        window_settings.close()
        return settings_changed


    ML_KEY = '-ML-'  # Multline's key


    def list1ToN(n):
        """
        Just a list of numbers from 1 to n
        """
        num = 0
        return_value = []
        for int in range(1, n + 1):
            return_value.append(num)
            num = num + 1
        return return_value

    def list0toN(n):
        """
        Just a list of numbers from 1 to n
        """
        num = 0
        return_value = []
        for int in range(0, n + 1):
            return_value.append(num)
            num = num + 1
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
            [sg.Text('Seleziona i file:')],
            [sg.Listbox(values=get_file_list(), select_mode=sg.SELECT_MODE_EXTENDED, size=(50, 20),
                        bind_return_key=True, key='-DEMO LIST-')],
            [sg.Text('Filtro (F1):', tooltip=filter_tooltip),
             sg.Input(size=(25, 1), focus=True, enable_events=True, key='-FILTER-', tooltip=filter_tooltip),
             sg.T(size=(15, 1), k='-FILTER NUMBER-')],
            [sg.Button('Avvia (Tutti i file)', enable_events=True, k='-ALL FILES-'), sg.B('Modifica file'),
             sg.B('Pulisci'), sg.B('Apri in cartella')],
            [sg.Text('Trova (F2):', tooltip=find_tooltip, visible=False),
             sg.Input(size=(25, 1), enable_events=True, key='-FIND-', tooltip=find_tooltip, visible=False),
             sg.T(size=(15, 1), k='-FIND NUMBER-')],
        ], element_justification='l', expand_x=True, expand_y=True)

        lef_col_find_re = sg.pin(sg.Col([
            [sg.Text('Trova (F3):', tooltip=find_re_tooltip),
             sg.Input(size=(25, 1), key='-FIND RE-', tooltip=find_re_tooltip), sg.B('Trova REGEX')]], k='-RE COL-',
            visible=False))

        column_filter_in = sg.pin(sg.Column([
            [sg.Text('Colonne:', font='Default 10', pad=(0, 0), justification='left',
                     tooltip='Colonne di interesse da usare come filtro'),
             sg.Input(sg.user_settings_get_entry('-col filter-'), key='-COLUMN FILTER-')]
        ])
        )

        column_filter_out = sg.pin(sg.Column([
            [sg.Input(sg.user_settings_get_entry('-fill input-'), tooltip='Lascia vuoto per creare colonne vuote.\n'
                                                                          'Se più colonne nuove vengono create'
                                                                          'ogni colonna sarà riempita con un elemento '
                                                                          'nella lista',
                      key='-FILL INPUT-', visible=False)]], justification='l')
        )

        column_filter = [column_filter_in, column_filter_out]

        right_col = [
            [sg.Text('Parole da sostituire:', font='Default 10', justification='left', grab=True)],
            [sg.Multiline(write_only=False, key=ML_KEY, tooltip='Le linee che iniziano con'
                                                                '"//" verranno ignorate, e sono considerate commenti'
                          , expand_y=True, expand_x=True,
                          default_text=(sg.user_settings_get_entry('-ml key-')))],
            [sg.Column([column_filter])],
            [sg.Text('Riga inizio tabella:', tooltip='Se insicuri lasciare valore di default'),
             sg.Combo(list1ToN(100), default_value=sg.user_settings_get_entry('-start line-'), key='-START LINE-',
                      readonly=False),
             sg.Button('Valori di controllo'),
             sg.Button('Materiali', k='-MATERIALS BUTTON-', visible=True)],
        ]

        options_at_bottom = sg.pin(
            sg.Column([[sg.CB('Mostra solo il primo risultato nel file', default=True, enable_events=True,
                              k='-FIRST MATCH ONLY-', visible=False),
                        sg.CB('Ignora maiuscolo', default=True, enable_events=True,
                              k='-IGNORE CASE-', visible=False),
                        sg.CB('Attendi completamento', default=False, enable_events=True,
                              k='-WAIT-', visible=False),
                        sg.CB('Crea colonne mancanti', tooltip='Decidi se creare le colonne mancanti '
                                                               'all\'interno del file di origine', default=False,
                              k='-CREATE MISSING-', enable_events=True),
                        sg.CB('Separa misure', k='-SPLIT MEASURES-',
                              tooltip='Separa le stringhe del tipo "100x60x70" in celle separate', default=True),
                        sg.CB('Filtra materiali', k='-MATERIALS LIST-', default=True,
                              tooltip='Filtra i materiali specificati ed assegnali ad una nuova colonna',
                              enable_events=True)
                        ]],
                      pad=(0, 0), k='-OPTIONS BOTTOM-', expand_x=True, expand_y=False),
            expand_x=True, expand_y=False)

        extra_options = sg.pin(
            sg.Column([[sg.Button('Guida'), sg.B('Impostazioni'), sg.Button('Esci')]],
                      element_justification='right', expand_x=True, expand_y=False)
            , expand_x=True, expand_y=False
        )

        choose_folder_at_top = sg.pin(
            sg.Column([[sg.T('Clicca Impostazioni per cambiare la cartella di destinazione'),
                        sg.Combo(sorted(sg.user_settings_get_entry('-folder names-')),
                                 default_value=sg.user_settings_get_entry('-demos folder-', ''),
                                 size=(50, 30), key='-FOLDERNAME IN-', enable_events=True, readonly=False),
                        ]], pad=(0, 0),
                      k='-FOLDER CHOOSE-'))
        # ----- Full layout -----

        layout = [[sg.Text('Celex', font=('Calibri', 30), text_color='#C4DFE6', pad=(10, 0))],
                  [choose_folder_at_top,
                   sg.FolderBrowse('Esplora file', target='-FOLDERNAME IN-', enable_events=True, key='-FN BROWSE IN-'),
                   sg.B('Pulisci', key='-CLEAN FOLDERNAME IN-'), sg.Text('©Bombo')],
                  [sg.Pane([sg.Column([[left_col], [lef_col_find_re]], element_justification='l', expand_x=True,
                                      expand_y=True),
                            sg.Column(right_col, element_justification='l', expand_x=True, expand_y=True, )],
                           orientation='h', relief=sg.RELIEF_SUNKEN, k='-PANE-', show_handle=False)],
                  [options_at_bottom], [extra_options]]

        # --------------------------------- Create Window ---------------------------------
        window = sg.Window('Celex', layout, finalize=True, icon=icon_path, resizable=True,
                           use_default_focus=False, size=(965, 670))
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
            window['-CREATE MISSING-'].update(visible=False)

        # sg.cprint_set_output_destination(window, ML_KEY)
        return window


    # --------------------------------- Main Program Layout ---------------------------------

    def saveSettings(values):
        """
        Used to save the settings for the current configuration.
        Saves the text in the ml element on the right main window
        """
        celex = Celex(values)
        cv_col_list = []
        sg.user_settings_set_entry('-ml key-', values[ML_KEY])
        sg.user_settings_set_entry('-start line-', values['-START LINE-'])
        sg.user_settings_set_entry('-col filter-', values['-COLUMN FILTER-'])
        sg.user_settings_set_entry('-fill input-', values['-FILL INPUT-'])
        if celex.readColumnList():
            cv_col_list = celex.readColumnList()
            if cv_col_list[len(cv_col_list) - 1] == '':
                cv_col_list.pop(len(cv_col_list) - 1)
        sg.user_settings_set_entry('-cv col combo-', cv_col_list)

    def getMaterialsString():
        """
        Get a string of the materials separated by semicolumns
        """
        materials_list = sg.user_settings_get_entry('-materials list-')
        return_value = ''
        for material in materials_list:
            return_value = return_value + material + ';'
        if return_value == '':
            return_value = 'Es: C40; C45; C70'
        return return_value

    def make_materials_window():
        """
        Creates a materials window element
        """
        layout = [
            [sg.Text('Lista dei materiali', font='DEFAULT 25')],
            [sg.Input(getMaterialsString(), tooltip='Separa ogni materiale con un ";"', k='-MATERIALS-')],
            [
                sg.Column([
                    [sg.Button('Ok', bind_return_key=True), sg.Button('Cancella')]
                ], justification='left')
            ]
        ]
        window_materials = sg.Window('Materiali', layout, icon=icon_path)
        return window_materials


    def saveMaterialsSettings(values):
        materials_list = values['-MATERIALS-'].split(';')
        for material in materials_list:
            i = materials_list.index(material)
            if material == '':
                materials_list.pop(i)
                continue
            if ' ' in material:
                material = material.replace(' ', '')
            if '\n' in material:
                material = material.replace('\n', '')
            materials_list[i] = material
        sg.user_settings_set_entry('-materials list-', materials_list)

    def materials_window():
        window_materials = make_materials_window()
        while True:
            event, values = window_materials.read()

            if event:
                saveMaterialsSettings(values)
            if event in ('Cancella', sg.WIN_CLOSED, sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
                break
            if event == 'Ok':
                window_materials.close()
                return True
        window_materials.close()
        return False

    def start(celex, values, path):
        celex_excel = Excel([path], values)

        celex.inputBufferList = celex.ignoreComments()

        excel_list_original = celex_excel.getExcelList()

        excel_list_original = celex_excel.removeSpaces(excel_list_original)

        # Get all the rules and column filters
        rule_list_dict = celex.getRuleListDict()
        columns_filter = celex.readColumnList()
        excel_list_filtered, missing_columns = celex_excel.filterByColumn(columns_filter, excel_list_original)

        # Creates the missing columns in each df object in excel_list_filget_demo_pathtered
        excel_list_final = []
        if missing_columns[get_path_list()[0]]:
            for excel in excel_list_filtered:
                index = excel_list_filtered.index(excel)
                excel_list_final.append(celex_excel.createColumns(missing_columns[get_path_list()[index]],
                                                                  values['-FILL INPUT-'],
                                                                  excel))

        # Joins each df with its corresponding missing columns df
        excel_list = celex_excel.joinDF(excel_list_filtered, excel_list_final)
        excel_list = celex_excel.removeSpaces(excel_list, removeLines=False)

        # SplitDict is a dict with column:values and SplitList is a list of values
        splitDict = {}

        for entry in excel_list_original:
            entry_index = excel_list_original.index(entry)
            list_row = celex.getRowListDF(entry, True)
            for row in list_row:
                # Creates a to_separate list where each entry is a
                # list of strings to separate in different cells
                to_separate = []
                for cell in row:
                    if not cell:
                        continue
                    if ' ' in str(cell) or '\n' in str(cell):
                        to_separate.append(cell)

                # Creates a buffer list to loop through the entries in
                # the to_separate list and performs a check for keywords

                isLastOneUsed = False

                for column in celex.getMissingColumns(entry):
                    splitDict[column] = []
                # For every string that needs to be split
                for cell in to_separate:
                    # For every word in each of these strings
                    for word in cell.split(' '):
                        index = cell.split(' ').index(word)
                        # For each keyword pulled from the ML element
                        for keyWord in celex.getKeyWords():
                            keyIndex = celex.getKeyWords().index(keyWord)
                            if isLastOneUsed:
                                continue
                            for column in celex.getMissingColumns(entry):
                                splitBuff, isLastOneUsed = celex.checkKeyWord(index, cell,
                                                                              word, keyWord,
                                                                              sg.user_settings_get_entry(
                                                                                  '-cv type-')[keyIndex],
                                                                              isLastOneUsed)
                                if splitBuff:
                                    splitDict[column].append(splitBuff)

        # df_buff is a dj object containing the values with their respective columns
        df_buff = pd.DataFrame()
        for column in splitDict:
            if splitDict[column]:
                df_buff[column] = splitDict[column]

        if excel_list:
            excel_list[0] = excel_list[0].join(df_buff)

        # Creates a measures dictionary if the checkbox is active, else it is empty
        measures_dict = celex_excel.separateMeasures(excel_list_original)
        df_buff = pd.DataFrame()
        for column in measures_dict:
            df_buff[column] = measures_dict[column]
        # Reindex excel_list_original to fit the join function
        df_buff_length = celex_excel.getNofRows(df_buff)
        if excel_list:
            if celex_excel.getNofRows(excel_list[0]) == df_buff_length:
                excel_list[0] = excel_list[0].set_axis(list0toN(df_buff_length - 1), axis='index')
                excel_list[0] = excel_list[0].join(df_buff)

        if values['-MATERIALS LIST-']:
            # Creates a material dict and a df_buff, then joins
            material_dict = celex_excel.getMaterialsDict(excel_list_original,
                                                         sg.user_settings_get_entry('-materials list-'))
            df_buff = pd.DataFrame()
            material_list = []
            for row_num in material_dict:
                material_list.append(material_dict[row_num])
            df_buff['Materiali'] = material_list
            if excel_list:
                excel_list[0] = excel_list[0].join(df_buff)

        # Substitutes all the rule keywords accordingly
        for df in excel_list:
            i = excel_list.index(df)
            for rule in rule_list_dict:
                dest = rule_list_dict[rule]
                df = df.replace(rule, dest)
            excel_list[i] = df

        # Save every df object in the list to a file
        for df in excel_list:
            i = excel_list.index(df)
            name = path.split('/')
            name = name[len(name) - 1]
            name = name.split('.')[0]
            celex_excel.saveToFile(name + '_' + str(i), df)


    def main():
        """
        The main program that contains the event loop.
        It will call the make_window function to create the window.
        """
        # Control variables
        is_fill_input_visible = False
        is_materials_visible = True

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
            if event:
                saveSettings(values)
            if event in (sg.WINDOW_CLOSED, 'Esci'):
                break
            if event == 'Guida':
                webbrowser.open('https://github.com/BomboBombone/celex', new=1)

            if event == 'Valori di controllo':
                control_variables_window()

            if event == '-DEMO LIST-':  # if double clicked (used the bind return key parm)
                if sg.user_settings_get_entry('-dclick runs-'):
                    event = 'Avvia'
                elif sg.user_settings_get_entry('-dclick edits-'):
                    event = 'Modifica file'
            if event == 'Modifica file':
                editor_program = get_editor()
                for file in values['-DEMO LIST-']:
                    # Takes the entry selected by the user and sets full_filename to the path of such file
                    full_filename, line = file, 1
                    full_filename = full_filename.split(' ')[1]
                    full_filename = full_filename[1:len(full_filename) - 1]
                    if line is not None:
                        if using_local_editor():
                            execute_command_subprocess(editor_program, full_filename)
                        else:
                            try:
                                sg.execute_editor(full_filename, line_number=int(line))
                            except:
                                execute_command_subprocess(editor_program, full_filename)

            elif event == '-ALL FILES-':
                for path in get_file_list():
                    celex = Celex(values)
                    path = path.split(' ')[1]
                    path = path[1:len(path) - 1]
                    start(celex, values, path)
            elif event == 'Avvia':
                path = celex.getDemoListEntry()[0]
                start(celex, values, path)

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
            elif event == '-CLEAN FOLDERNAME IN-':
                file_list = get_file_list()
                window['-FOLDERNAME IN-'].update('')
                window['-FOLDERNAME IN-'].update(values=[])
                window['-FILTER-'].update('')
                window['-FILTER NUMBER-'].update(f'{len(file_list)} file')
                window['-FIND-'].update('')
                window['-DEMO LIST-'].update(file_list)
                window['-FIND NUMBER-'].update('')
                window['-FIND RE-'].update('')
            elif event == '-FOLDERNAME IN-':
                sg.user_settings_set_entry('-demos folder-', values['-FOLDERNAME IN-'])
                sg.user_settings_set_entry('-folder names-', list(
                    set(sg.user_settings_get_entry('-folder names-', []) + [values['-FOLDERNAME IN-'], ])))
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
                    for file in celex.getDemoListEntry():
                        file_path = os.path.dirname(file)
                        if running_windows():
                            file_path = file_path.replace('/', '\\')
                        execute_command_subprocess(explorer_program, file_path)

            if event == '-CREATE MISSING-':
                if not is_fill_input_visible:
                    window['-FILL INPUT-'].update(visible=True)
                    is_fill_input_visible = True
                else:
                    window['-FILL INPUT-'].update(visible=False)
                    is_fill_input_visible = False
            if event == '-MATERIALS LIST-':
                if not is_materials_visible:
                    window['-MATERIALS BUTTON-'].update(visible=True)
                    is_materials_visible = True
                else:
                    window['-MATERIALS BUTTON-'].update(visible=False)
                    is_materials_visible = False
            if event == '-MATERIALS BUTTON-':
                materials_window()
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

        try:
            buff = sg.user_settings_get_entry('-demos folder-')
        except FileNotFoundError:
            sg.user_settings_set_entry('-demos folder-', ['C:/Program Files/Celex'])

        # Set addditional user settings
        if not sg.user_settings_get_entry('-ml key-'):
            sg.user_settings_set_entry('-ml key-', '//Inserisci qui la lista dei valori di controllo (Uno per linea).\n'
                                                   '//Esempio:\n'
                                                   'C45 = PC456T\n'
                                                   'C40 = PC40T67\n')
        if not sg.user_settings_get_entry('-col filter-'):
            sg.user_settings_set_entry('-col filter-', 'Es: Commessa; Pz; Misure Finite; Misure')
        if not sg.user_settings_get_entry('-fill input-'):
            sg.user_settings_set_entry('-fill input-', 'Es: 769KF3; Celex')
        sg.user_settings_set_entry('-start line-', 0)

        # Control values window settings
        sg.user_settings_set_entry('-cb value-', [])
        sg.user_settings_set_entry('-cv input-', [])
        sg.user_settings_set_entry('-cv type-', [])
        # Sets the combo list choices to display on the control window
        sg.user_settings_set_entry('-cv col combo-', [])
        # Each entry represents the corresponding entry in the control window
        sg.user_settings_set_entry('-cv col entry-', [])

        # Material window settings
        if not sg.user_settings_get_entry('-materials list-'):
            sg.user_settings_set_entry('-materials list-', [])

        # Set the default output to Desktop/Celex
        default_output_folder = os.path.join(os.environ["HOMEPATH"], "Desktop")
        default_output_folder = 'C:\\' + default_output_folder
        default_output_folder = os.path.join(default_output_folder, "Celex")

        if not os.path.exists(default_output_folder):
            os.mkdir(default_output_folder)

        main()
