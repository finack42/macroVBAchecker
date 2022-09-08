import os
import subprocess

from win32com.client import Dispatch
import re
import sys


attention_string = "WARNING: Suspicious keywords were found:\n"
endOfMacro_string = "\n====================== END OF MACRO ======================\n"
noKeywords_string = "INFO: There aren't any suspicous keywords!\n"
horizontal_line = "\n------------------------------------------------------------------------------------------\n"


def check_macro(macro, doc_name, mod_name):
    checkdoc_string = "INFO: Check of document " + doc_name + " and module " + mod_name + " is starting..."
    with open(fileNameKeywords, "a") as outputfile:
        outputfile.write(checkdoc_string)
    print(checkdoc_string)
    p = re.compile("dde|msiexec|winmgmts|Win32_Process|wscript|base64|StrReverse|shell|"
                   "reverse|pwd|password|cred|credentials|http|ssh|ftp|powershell|cmd|&H|\\\\x", re.IGNORECASE)
    m = p.findall(macro)

    if m:
        print(attention_string)
        print("\n".join(m))
        print(endOfMacro_string)
        with open(fileNameKeywords, "a") as outputfile:
            outputfile.writelines(["\n" + attention_string, "\n".join(m), endOfMacro_string])
    else:
        print(noKeywords_string)
        with open(fileNameKeywords, "a") as outputfile:
            outputfile.write("\n" + noKeywords_string)


if __name__ == '__main__':

    # wbpath = 'C:\\Users\\A7808\\PycharmProjects\\MacroChecker\\example.xlsm'

    wbfile = os.path.join(sys.path[0], sys.argv[1])
    wbfile_ext = sys.argv[1].split('.')[1]
    fileNameKeywords = sys.argv[1] + "_keywords.txt"
    fileNameOle = sys.argv[1] + "_OLE.txt"
    no_of_modules = 0
    xl = Dispatch("Excel.Application")
    xl.Visible = 0
    wb = xl.Workbooks.Open(wbfile)

    if wbfile_ext == 'xlsm':

        m_source_code = []
        file = open(fileNameKeywords, 'w+')

        # vbcode = wb.VBProject.VBComponents("Module3").CodeModule
        for module in wb.VBProject.VBComponents:
            # print(module.CodeModule.Lines(1, module.CodeModule.CountOfLines))
            # 1 = VBA Code (https://bettersolutions.com/vba/visual-basic-editor/extensibility-object-model.htm)
            if module.Type == 1:
                module_name = module.CodeModule.Name
                vbcode_object = wb.VBProject.VBComponents(module_name).CodeModule
                # print(vbcode_object.Lines(1, vbcode_object.CountOfLines))
                if vbcode_object.CountOfLines == 0:
                    print("INFO: Module \"%s\" is empty and won't be checked..." % module_name)
                    break
                m_source_code.append(vbcode_object.Lines(1, vbcode_object.CountOfLines))
                check_macro(m_source_code[no_of_modules], wb.Name, module_name)
                no_of_modules += 1
    else:
        bas_file = open(wbfile, 'r')
        check_macro(bas_file.read(), sys.argv[1], '\"bas-File\"')

    # To check if a module name exists (throws an error if not existing)
    # vbcode = wb.VBProject.VBComponents("Module3").Name
    # print(vbcode.Lines(1, vbcode.CountOfLines))

    total_macros = horizontal_line + "There is a total number of " + str(
        no_of_modules) + " VBA Macros in document \"" + wb.Name + "\".\n"
    bas_file_resume = "This is a single bas-file called \"" + wb.Name + "\"."
    if wbfile_ext == "xlsm":
        print(total_macros)
        with open(fileNameKeywords, "a") as myfile:
            myfile.write(total_macros)

    wb.Close(False)

    with open(fileNameOle, 'w', encoding='utf-8') as f:
        process = subprocess.Popen(['oleid', sys.argv[1]], stdout=f, universal_newlines=True)

    # while True:
    #     output = process.stdout.readline()
    #     print(output.strip())
    #     return_code = process.poll()
    #     if return_code is not None:
    #         # Process has finished, read rest of the output
    #         for output in process.stdout.readlines():
    #             print(output.strip())
    #         break
