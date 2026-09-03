import tkinter as tk


# def show_popup(message):
    # root = tk.Tk()
    # root.withdraw()
    # messagebox.showinfo("Information Message", message)
    # root.destroy()
def show_popup(message, duration=3):
    """Show a desktop popup with the given message."""
    root = tk.Tk()      #Tkinter needs a root window to manage its GUI.Tkinter creates a visible main window.
    root.withdraw()     #hides the root window.

    popup = tk.Toplevel(root)   #creates popup window  with root as its parent window
    #Toplevel is a Tkinter class used to create an additional window.
    #The popup needs a Tkinter application/root associated with it.
    #root is the parent/application and popup is a child window.
    popup.title("Information Message")

    # Get screen dimensions
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()

    # Calculate center position
    width=450
    height=150
    x = (screen_width - width) // 2         #x and y are coordinates of windows where popup will appear
    y = (screen_height - height) // 2

    # Set popup size and position
    popup.geometry(f"{width}x{height}+{x}+{y}")    
    # popup.geometry("450x150")               #dimensions of popup window in pixels

    label = tk.Label(popup, text=message)     #Creates a Label inside popup and display the contents of message.label is child of popup
    #A label is used to display text or other simple information.
    label.pack(pady=40)                       #pady means vertical padding.adds 30 pixels of vertical space around the label.
    
    
    root.after(duration * 1000, root.destroy)   #Schedule "root.destroy" to run after 3000 milliseconds(=3sec).
    root.mainloop()
    
    # root.destroy()        #Completely close/destroy the window immediately and its Tkinter resources.
    #root.destrot()  executes the ftn and 
