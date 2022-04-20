# csvSearch
# À faire: 1) multiple search criteria, 2) dynamic window sizing (sizegrip element?)
# Future: wxPython?
#
revision = '1.0'
import csv
import PySimpleGUI as sg
import operator     # sorted()
import pyperclip    # for copy to clipboard
import sys          # for argv

def terminate_with_error(outstring):
    sg.popup_error(outstring)
    sys.exit(outstring)

separators = {'Comma':',', 'Semicolon':';', 'Tab':'\t', 'Period':'.', 'Space':' '}
encodings = ['None', 'utf-8', 'utf-32', 'ascii' ]   # https://docs.python.org/3/library/codecs.html
std_win_sizes = ['800x500', '500x300', '1200x800']
parameters = {  'file': None, 
                'separator': 'Comma', 
                'encoding': encodings[0], 
                'has_headings': True, 
                'search_column': '0', 
                'search_type': 'starting_with', 
                'sg_theme': 'DarkGrey7',
                'window_size': std_win_sizes[0],    # <width>x<height>
                'tables_width_ratio': '0.625'
                }
csvSearch_help = f'''csvSearch {revision} : search a csv file and copy selected data to clipboard
Program can be called without argument
usage: csvsearch <parameter>=<value> <parameter>=<value>...
<parameter>: {list(parameters.keys())}
<value>: see also graphical user interface
    separator: any character or {list(separators.keys())}
    encoding: {encodings} or any python codec name
    tables_width_ratio: float number between 0 and 1
exemple:
    csvsearch file=myfile.csv separator=Semicolon encoding=utf-8 sg_theme=DarkGrey7
'''
# Arguments provided on command line: extract parameters
#       Syntax is <param>=<value>
nb_of_argv = len(sys.argv)
if len(sys.argv) > 1:
    if sys.argv[1] == '-help':
        sys.exit(csvSearch_help)
    for n in range(1, nb_of_argv):
        arg = sys.argv[n]
        split = arg.find('=')
        if split < 1:
            terminate_with_error(f'csvSearch: invalid argument: {arg}. Use -help for help.')
        param = arg[:split]
        value = arg[split+1:]
        if param not in list(parameters.keys()):
            terminate_with_error(f'csvSearch: unknown parameter: {param}. Use -help for help.')
        if param == 'has_headings':     # Special case: has_heading is boolean
            value = value=='True'
        parameters[param] = value
#    print(parameters)

# no file name provided in cmd line
#       ->use window
sg.theme(parameters['sg_theme'])
filename = parameters['file']
if not filename:      
    layout = [[sg.Text('CSV file:')],
                [sg.In(), sg.FileBrowse()],
                [sg.Text('Separator:'), sg.Combo(list(separators.keys()), default_value=parameters['separator'], key='-SEPARATOR-'),
                    sg.Text('Encoding:'), sg.Combo(encodings, default_value=parameters['encoding'], key='-ENCODING-'),
                    sg.Checkbox('has headings', default=parameters['has_headings'], key='-HEADINGS-')],
                [sg.Text('Window size:'), sg.Combo(std_win_sizes, default_value=parameters['window_size'], key='-WINSIZE-'),
                    sg.Text('Theme:'), sg.Combo(sg.theme_list(), default_value=parameters['sg_theme'], key='-THEME-')],
                [sg.Open(), sg.Cancel()]]
    event, values = sg.Window('csvSearch', layout).read(close=True)
    # print(event, values)
    if event == 'Cancel':
        sys.exit("csvSearch: execution cancelled by the user")
    filename = values[0]
    parameters['separator'] = values['-SEPARATOR-']
    parameters['encoding'] = values['-ENCODING-']
    parameters['has_headings'] = values['-HEADINGS-']       # Attention boolean!
    parameters['window_size'] = values['-WINSIZE-']
    parameters['sg_theme'] = values['-THEME-']

if not filename:
    terminate_with_error("csvSearch: missing file name")
if parameters['encoding'] == 'None':
    parameters['encoding'] = None
if parameters['separator'] in list(separators.keys()):
    parameters['separator'] = separators[parameters['separator']]

# Load data and initialise stuf
sg.theme(parameters['sg_theme'])        # again
alldata = []
try:
    with open(filename, newline='', encoding=parameters['encoding']) as thefile:
        reader = csv.reader(thefile, delimiter=parameters['separator'])
        for row in reader:
            alldata.append(row)
except Exception as e:
    terminate_with_error("csvSearch: error reading file")
if parameters['has_headings']:
    headers = alldata[0]
    data = alldata[1:]
else:       # create headers
    data = alldata
    headers = [f'Col {col}' for col in range(len(data[0]))]
datasize = len(data)
win_size = parameters['window_size']
try:        # window sizing stuf
    split = win_size.find('x')
    frames_width = int(win_size[:split]) - 50
    frames_height = int(win_size[split+1:]) - 50
    search_table_width = int((frames_width)*float(parameters['tables_width_ratio']))
    value_table_width = frames_width - search_table_width
    search_num_rows = int(frames_height / 17)
except Exception as e:
    terminate_with_error('csvSearch: invalid window_size or tables_width_ratio parameter. Use -help for help.')
linedata = []
for header in headers:
    linedata.append([header, ' '])

# Layouts and frames for main window
col1 = sg.Column([[sg.Text('Search:'), sg.Combo(headers, default_value=headers[int(parameters['search_column'])], key='-COLUMN-'), 
                        sg.Combo(['starting_with', 'containing'], default_value=parameters['search_type'], key='-SEARCHTYPE-'), 
                        sg.Input(key='-IN-', enable_events=True, size=(25, None))]])
col2 = sg.Column([
    [sg.Frame('Click to select',[[sg.Table(values=data, headings=headers, 
                                vertical_scroll_only=False,
                                hide_vertical_scroll=False,
                                auto_size_columns=True,
                                justification='left',
                                num_rows=search_num_rows,
                                key='-TABLE-',
                                enable_events=True,
                                expand_x = False,
                                expand_y = False,
                                enable_click_events=True)]],size = (search_table_width,frames_height),)],])
col3 = sg.Column([
    [sg.Frame('Click to copy',[[sg.Table(values=linedata, headings=('Label', 'Value'),
                                vertical_scroll_only=False,
                                hide_vertical_scroll=False,
                                auto_size_columns=True,
                                justification='left',
                                row_height=20,
                                key='-VALUETABLE-',
                                expand_x = True,
                                expand_y = True,
                                enable_click_events=True,
                                enable_events=True)]],size = (value_table_width,frames_height),)],])
col4 = sg.Column([[sg.Button('Exit')]])
layout = [[col1],
          [col2, col3],
          [col4]] 

# Load window and loop on events
window = sg.Window('csvSearch: '+filename, layout, resizable = False)
search = data
while True:
    event, values = window.read()
    # print(event, values)
    if event in (sg.WIN_CLOSED,'Exit'):
        break
    if event == '-IN-':     # Search
        search = []
        index = headers.index(values['-COLUMN-'])
        if values['-SEARCHTYPE-'] == 'starting_with':
            for line in data:
                if line[index].lower().startswith(values['-IN-'].lower()):
                    search.append(line)
        else:
            for line in data:
                if line[index].lower().find(values['-IN-'].lower()) >= 0:
                    search.append(line)
            
        window['-TABLE-'].update(search)
    elif isinstance(event, tuple):
        # TABLE CLICKED Event has value in format ('-TABLE=', '+CLICKED+', (row,col))
        row_num_clicked = event[2][0]
        if row_num_clicked is not None:
            if event[0] == '-TABLE-':
                if row_num_clicked == -1:           # Header was clicked ->sort, for fun
                    col_num_clicked = event[2][1]
                    search = sorted(search, key=operator.itemgetter(col_num_clicked))
                    window['-TABLE-'].update(search)
                else:
                    linedata = []
                    for x in range(len(headers)):
                        linedata.append([headers[x], search[row_num_clicked][x]])
                    window['-VALUETABLE-'].update(linedata)
            elif event[0] == '-VALUETABLE-':    # value table was clicked
                pyperclip.copy(linedata[row_num_clicked][1])    # copy value to clipboard
window.close()