import PySimpleGUI as sg
import sys
from excelCompare import xlsxFileCompae

if __name__ == '__main__':
    if len(sys.argv) == 1:
        event, values = sg.Window('选择2个Excel文件进行比较',
                                  [[sg.Text('请选择excel文件 1.')],
                                   [sg.In(), sg.FileBrowse(
                                       file_types=(("excel files", "*.xlsx"),))],
                                   [sg.Text('请选择excel文件 2.')],
                                   [sg.In(), sg.FileBrowse(
                                       file_types=(("excel files", "*.xlsx"),))],
                                      [sg.Ok(button_text='开始对比'), sg.Cancel()]
                                   ]).read(close=True)
        # print(values)
        file1, file2 = values[0], values[1]
    elif len(sys.argv) == 3:
        file1, file2 = sys.argv[1], sys.argv[2]



if not (file1 and file2):
    sg.popup("Cancel", "No filename supplied")
    raise SystemExit("Cancelling: no filename supplied")

else:
    # sg.popup('The filename you chose was', fname)
    if file1.endswith('.xlsx') and file2.endswith('.xlsx'):
        xlsxFileCompae(file1, file2, True)
        
