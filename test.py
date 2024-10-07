import customtkinter
import Docs_maker as DM
from github import Github
import datetime

def update():
    g = Github()
    user = g.get_user('uvaprol')
    repo = user.get_repo('Auto_Docs')
    content = str(((repo.get_contents('test.txt')).decoded_content).decode('utf-8'))
    content = content.replace('\\r', '')
    content = content.replace('\\n', '\n')
    content = content.replace('    ', '\t')
    new_code = content
    print(new_code)


with open('seting_update.txt', 'r') as s:
    try:
        seting = s.readline().split('-')
        date = datetime.date.today() - datetime.date(int(seting[0]), int(seting[1]), int(seting[2]))
        if date.days >= 7:
            update()
            with open('seting_update.txt', 'w') as s:
                s.write(str(datetime.date.today()))
    except:
        update()
        with open('seting_update.txt', 'w') as s:
            s.write(str(datetime.date.today()))





with open('test.txt', 'r', encoding='UTF-8') as code:
    code_compile = ''
    for line in code.readlines():
        line = line.replace('    ', '\t')
        code_compile += line + '\n'
    exec(compile(code_compile, "<string>", "exec"))



