
# from settings import config as conf
import tkinter as tk
from tkinter import ttk
from tkinter import TclError
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
from docx import Document
import re
import datetime

class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.initUI()

    def initUI(self):
        # self.parent.title(conf().title)
        # self.parent.iconbitmap(conf().icon)
        # self.parent.iconphoto(False, tk.PhotoImage(file=conf().appIcon))
        # self.parent.overrideredirect(False)
        # getattr(self.parent, conf().iconify)()
        # self.parent.minsize(*conf().minsize)
        # self.parent.maxsize(*conf().maxsize)
        # self.parent.geometry(conf().geometry) # ширина=500, высота=400, x=300, y=200
        # self.parent.resizable(*conf().resizable) # размер окна может быть изменён только по горизонтали
        self.parent.protocol('WM_DELETE_WINDOW', self.close) # обработчик закрытия окна
        # self.parent.protocol('WM_TAKE_FOCUS', self.focus_)
        # getattr(self.parent, conf().tkraise)()
        self.parent.grab_set()
        self.buildWindow()

    def buildWidget(self, target: any, resource: any, widget: str, name: str, **settings):
        try:
            setattr(self, name, getattr(resource, widget)(master=target, cnf=settings))
        except TclError:
            setattr(self, name, getattr(resource, widget)(master=target, **settings))
        return getattr(self, name)

    def postWidget(self, widget: any, **settings):
        widget.pack(cnf=settings)

    def build(self, target: any, resource: any, widget: str, name: str, settingsWidget: dict, settingsPack: dict):
        self.postWidget(self.buildWidget(target=target, resource=resource, widget=widget, name=name, **settingsWidget), **settingsPack)
        return getattr(self, name)

    def buildWithLabel(self, label: any, widget: str, name: str):
        pass

    def buildWindow(self):
        frameMain = self.build(
            target = self.parent,
            resource = tk,
            widget = 'Frame',
            name = 'frameMain',
            settingsWidget = dict(
                bg = 'white',
                width = 1,
                height = 1
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'both',
                expand = True
                ))

        frameTop = self.build(
            target = frameMain,
            resource = tk,
            widget = 'Frame',
            name = 'frameTop',
            settingsWidget = dict(
                bg = 'black',
                # padx = 5,
                width = 1,
                height = 1
                ),
            settingsPack = dict(
                side = 'top',
                fill = 'both',
                expand = True
                ))

        table = self.build(
            target = frameTop,
            resource = ttk,
            widget = 'Treeview',
            name = 'table',
            settingsWidget = dict(
                columns = ('key', 'value'),
                show = 'headings',
                selectmode = 'browse'
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'both',
                expand = True
                ))

        table.heading('key', text='свойство')
        table.heading('value', text='значение')

        style = ttk.Style()
        style.theme_use('classic')
        # style.configure('Treeview.Heading', background="black", foreground="white")

        scrollbarVertical = self.build(
            target = frameTop,
            resource = ttk,
            widget = 'Scrollbar',
            name = 'scrollY',
            settingsWidget = dict(
                orient = tk.VERTICAL, 
                command = table.yview
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'y',
                expand = False
                ))

        table.configure(yscroll=scrollbarVertical.set)

        table.bind('<<TreeviewSelect>>', self.getTableValues)
        table.bind('<Button-3>', self.selectionRemove)

        table.drop_target_register(DND_FILES)
        table.dnd_bind('<<Drop:DND_Files>>', lambda e: self.getDocProperties(self.getPath(e.data)))

        self.itemTag = None
        self.path = None

        frameBottom = self.build(
            target = frameMain,
            resource = tk,
            widget = 'Frame',
            name = 'frameBottom',
            settingsWidget = dict(
                bg = 'black',
                # padx = 5,
                width = 1,
                height = 1
                ),
            settingsPack = dict(
                side = 'bottom',
                fill = 'x',
                expand = False
                ))

        editorLabel = self.build(
            target = frameBottom,
            resource = tk,
            widget = 'LabelFrame',
            name = 'labelList',
            settingsWidget = dict(
                text = 'Редактирование'
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'both',
                expand = True
                ))

        editorProperty = self.build(
            target = editorLabel,
            resource = tk,
            widget = 'Entry',
            name = 'editorProperty',
            settingsWidget = dict(
                justify = 'left'
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'x',
                expand = True
                ))

        editorProperty.bind('<KeyRelease>', self.getEditorValues)

        editorValue = self.build(
            target = editorLabel,
            resource = tk,
            widget = 'Entry',
            name = 'editorValue',
            settingsWidget = dict(
                justify = 'left'
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'x',
                expand = True
                ))

        editorValue.bind('<KeyRelease>', self.getEditorValues)

        editorAdd = self.build(
            target = editorLabel,
            resource = tk,
            widget = 'Button',
            name = 'editorDel',
            settingsWidget = dict(
                text = 'Добавить',
                bg = 'white',
                width = len('Добавить'),
                height = 1,
                activebackground = 'black',
                activeforeground = 'white',
                justify = 'center',
                command = self.addTableValues
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'x',
                expand = False
                ))

        editorDel = self.build(
            target = editorLabel,
            resource = tk,
            widget = 'Button',
            name = 'editorDel',
            settingsWidget = dict(
                text = 'Удалить',
                bg = 'white',
                width = len('Удалить'),
                height = 1,
                activebackground = 'black',
                activeforeground = 'white',
                justify = 'center',
                command = self.delTableValues
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'x',
                expand = False
                ))

        self.propertyModeVar = tk.BooleanVar()

        changeProperty = self.build(
            target = editorLabel,
            resource = tk,
            widget = 'Checkbutton',
            name = 'propertyMode',
            settingsWidget = dict(
                text = 'Режим',
                variable = self.propertyModeVar,
                onvalue = True,
                offvalue = False,
                command = self.changeMode
                ),
            settingsPack = dict(
                side = 'left',
                fill = 'x',
                expand = False
                ))

    def close(self):
        if not self.propertyModeVar.get():
            for property, value in [self.table.item(line)['values'] for line in self.table.get_children()]:

                if isinstance(value, int):
                    self.properties[property] = int(value)
                if isinstance(value, str):
                    if '#' in value and re.match('^[A-Za-zА-Яа-я0-9Ёё]', value):
                        # print(re.search(r'[\d*]+[\D]', value)[0])
                        match = re.search(r'\d+#', value)
                        if match:
                            self.properties[property] = int(match[0][:-1])*match[0][-1]
                    elif re.match('\d+-\d+-\d+T\d+:\d+:\d+Z', value):
                        self.properties[property] = datetime.datetime.strftime(datetime.datetime.strptime(value, '%Y-%m-%dT%H:%M:%SZ'), '%d.%m.%y')
                    
                    elif value == 'True' or value == 'False':
                        self.properties[property] = bool(value)
                    else:
                        self.properties[property] = value
        else:
            for property, value in [self.table.item(line)['values'] for line in self.table.get_children()]:
                if property == 'author':
                    self.properties.author = value
                elif property == 'category':
                    self.properties.category = value
                elif property == 'comments':
                    self.properties.comments = value
                elif property == 'content_status':
                    self.properties.content_status = value
                elif property == 'created': # datetime
                    if value != 'None':
                        self.properties.created = datetime.datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
                    # else:
                    #     self.properties.created = value
                elif property == 'identifier':
                    self.properties.identifier = value
                elif property == 'keywords':
                    self.properties.keywords = value
                elif property == 'language':
                    self.properties.language = value
                elif property == 'last_modified_by':
                    self.properties.last_modified_by = value
                elif property == 'last_printed': # datetime
                    if value != 'None':
                        self.properties.last_printed = datetime.datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
                    # else:
                    #     self.properties.last_printed = None
                elif property == 'modified': # datetime
                    if value != 'None':
                        self.properties.modified = datetime.datetime.strptime(value, '%Y-%m-%d %H:%M:%S')
                    # else:
                    #     self.properties.modified = value
                elif property == 'revision' and int(value) > 0: # int
                    self.properties.revision = int(value)
                elif property == 'revision' and int(value) <= 0: # int
                    self.properties.revision = 1
                elif property == 'subject':
                    self.properties.subject = value
                elif property == 'title':
                    self.properties.title = value
                elif property == 'version':
                    self.properties.version = value

        if self.path:
            path, ext = os.path.splitext(self.path)
            path = re.sub(r'\[\d+\]', '', path)
            propertiesCore = self.doc.core_properties
            revision = propertiesCore.revision
            propertiesCore.revision = revision + 1
            self.doc.save(f'{path}[{revision + 1}]{ext}')
        if hasattr(self, 'closeFlag'):
            self.closeFlag += 1
            [self.table.delete(self.table.item(line)['tags'][0]) for line in self.table.get_children()]
            if self.closeFlag == 2:
                self.parent.quit()
        else:
            self.parent.quit()

    def focus_(self):
        # self.parent.quit()
        # self.parent.deiconify()
        pass

    def getPath(self, data):
        files = data.split('} {')
        if len(files) != 1:
            return('')
        else:
            file = files[0]
            if os.path.exists(file):
                return(file)
            else:
                fileWithoutSlashes = file.replace('\\', '')
                fileWithoutBrackets = file.replace('{', '').replace('}', '')
                if os.path.exists(fileWithoutBrackets):
                    return(fileWithoutBrackets)
                elif os.path.exists(fileWithoutSlashes):
                    return(fileWithoutSlashes)
                elif os.path.exists(file[1:-1]):
                    return(file[1:-1])
                else:
                    return('')

    def getDocProperties(self, path: str):
        self.path = path
        if path.endswith('.doc') or path.endswith('.docx'):
            self.doc = Document(path)
            self.properties = self.doc.custom_properties
            for child in self.properties._element:
                print(child.get('name'), child[0].text)
                key, value = child.get('name'), child[0].text
                tag = key.replace(' ', '_') if key != None else None
                if not tag in (self.table.item(line)['tags'][0] for line in self.table.get_children()) and tag != None:
                    self.table.insert(parent='', index='end', iid=tag, values=(key, value), tags=tag)
        self.closeFlag = 0

    def changeMode(self):
        if hasattr(self, 'doc'):
            [self.table.delete(self.table.item(line)['tags'][0]) for line in self.table.get_children()]
            self.itemTag = None
            self.closeFlag = 0
            if self.propertyModeVar.get():
                self.properties = self.doc.core_properties
                for property in ('author', 'category', 'comments', 'content_status', 'created', 'identifier', 'keywords', 'language', 'last_modified_by', 'last_printed', 'modified', 'revision', 'subject', 'title', 'version'):
                    key, value = property, getattr(self.properties, property)
                    tag = key.replace(' ', '_') if key != None else None
                    self.table.insert(parent='', index='end', iid=tag, values=(key, value), tags=tag)
            else:
                self.properties = self.doc.custom_properties
                for child in self.properties._element:
                    key, value = child.get('name'), child[0].text
                    tag = key.replace(' ', '_') if key != None else None
                    self.table.insert(parent='', index='end', iid=tag, values=(key, value), tags=tag)

    def getTableValues(self, event):
        item = self.table.item(self.table.selection())
        tags = item['tags']
        values = item['values']
        self.itemTag = tags[0] if len(tags) != 0 else None
        if len(values) == 2:
            property, value = values
        else:
            property, value = None, None
        if property != None or value != None:
            self.editorProperty.delete(0, 'end')
            self.editorProperty.insert(0, property)
            self.editorValue.delete(0, 'end')
            self.editorValue.insert(0, value)
            self.editorValue.focus()

    def delTableValues(self):
        if self.itemTag:
            self.table.selection_remove(self.itemTag)
            self.table.delete(self.itemTag)
            self.itemTag = None

    def addTableValues(self):
        tags = [self.table.item(line)['tags'][0] for line in self.table.get_children()]
        key, value = self.editorProperty.get(), self.editorValue.get()
        tag = key.replace(' ', '_')
        if len(key) != 0 and len(value) != 0 and not tag in tags and not self.propertyModeVar.get():
            self.table.insert(parent='', index='end', iid=tag, values=(key, value), tags=tag)

    def selectionRemove(self, event):
        if self.itemTag:
            self.table.selection_remove(self.itemTag)
            self.editorProperty.delete(0, 'end')
            self.editorValue.delete(0, 'end')

    def getEditorValues(self, event):
        if self.itemTag != None:
            self.table.set(self.itemTag, 0, self.editorProperty.get())
            self.table.set(self.itemTag, 1, self.editorValue.get())

if __name__ == '__main__':
    # root = tk.Tk()
    root = TkinterDnD.Tk()
    MainApplication(root).pack(side='top', fill='both', expand=True)
    root.mainloop()