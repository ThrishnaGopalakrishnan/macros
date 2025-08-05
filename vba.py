import xlwings as xw
import os

def vba_code(xlsm_path, export_dir):
    wb = xw.Book(xlsm_path)
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    for component in wb.api.VBProject.VBComponents:
        name = component.Name
        type_ = component.Type
        if type_ in [1, 2, 3]:  # 1=Module, 2=Class, 3=Form
            filename = f"{name}.bas" if type_ == 1 else f"{name}.cls"
            with open(os.path.join(export_dir, filename), 'w', encoding='utf-8') as f:
                f.write(component.CodeModule.Lines(1, component.CodeModule.CountOfLines))

    wb.close()

# Replace this with your .xlsm file path
vba_code('min_required_operating_system_version_ modeling (2).xlsm', 'vba_modules')
