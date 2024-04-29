import tkinter as tk
from tkinter import messagebox, filedialog
import os
from excel import excel_encrypt_file
from word import word_encrypt_file

class App:
    def __init__(self, root):
        self.root = root
        self.root.title('数据加密工具')

        self.label = tk.Label(root, text='Excel/Word 数据加密', font=('Helvetica', 18, 'bold'))
        self.file_entry = tk.Entry(root, width=40)
        self.choose_button = tk.Button(root, text='选择文件', command=self.choose_file)
        self.process_button = tk.Button(root, text='处理文件', command=self.process_file)

        
        self.label.grid(row=0, column=0, columnspan=2, pady=10, padx=20)
        self.file_entry.grid(row=1, column=0, columnspan=2, pady=10, padx=20)
        
        
        self.choose_button.grid(row=3, column=0, pady=10, padx=20, sticky='w')
        self.process_button.grid(row=3, column=1, pady=10, padx=20, sticky='e')


    def choose_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
    
    def get_file_type(self,file_path):
        _, file_extension = os.path.splitext(file_path)
        if file_extension in ['.xls', '.xlsx', '.csv']:
            return 'Excel'
        elif file_extension in ['.doc', '.docx']:
            return 'Word'
        else:
            return 'Unknown' 
    
    ...
    # 在这里，我们将 process_file 方法进行修改，使其可以根据选择的识别模式来调用相应的处理方法。
    def process_file(self):
        file_path = self.file_entry.get()
        base_name = os.path.splitext(file_path)[0]
        extension = os.path.splitext(file_path)[1]
        new_file_path = base_name + "-mask" + extension
        file_type = self.get_file_type(file_path)

        if file_type == "Word":
            word_encrypt_file(file_path,new_file_path)
            messagebox.showinfo('Info', f'File {file_path} is a {file_type} file and has been processed!')
        elif file_type == "Excel":
            excel_encrypt_file(file_path,new_file_path)
            messagebox.showinfo('Info', f'File {file_path} is a {file_type} file and has been processed!')
        else:
            messagebox.showinfo('错误', f'File {file_path} is a {file_type} 不支持!')

        

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
