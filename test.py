import customtkinter
import Docs_maker as DM

with open('test.txt', 'r', encoding='UTF-8') as code:
    code_compil = ''
    for line in code.readlines():
        code_compil += line.replace('    ', '\t')
    exec(compile(code_compil, "<string>", "exec"))

