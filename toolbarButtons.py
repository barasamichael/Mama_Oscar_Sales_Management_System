import tkinter as tk
from PIL import Image, ImageTk
class ToolbarIcon:
    """ 
     -----USAGE----- 
    Example = Example(parent,  
                         filename = [filepath],
                         button_command = [command,
                         icon_side = [side]]) 
   =====ILLUSTRATIONS=====
   def main():

        root = Tk()
        toolbar = Frame(root)
        
        app = Example(toolbar,'/sdcard/unnamed_1.png',None,LEFT)
        app = Example(toolbar,'/sdcard/unnamed.png',None,RIGHT)
        
        root.mainloop()  
         
    if __name__ == '__main__':
        main()
    """

    def __init__(self, parent ,filename1= None,button_command = None,icon_side = tk.LEFT):
        super().__init__()

        self.initUI(parent,filename1,icon_side,button_command)


    def initUI(self,parent,filename1,icon_side,button_command):
        
        self.img = Image.open(filename1)
        eimg = ImageTk.PhotoImage(self.img)

        exitButton = tk.Button(parent, image=eimg, relief=tk.FLAT,
            command=button_command)
        exitButton.image = eimg
        exitButton.pack(side=icon_side, padx=2, pady=2)

if __name__ == '__main__':
    pass