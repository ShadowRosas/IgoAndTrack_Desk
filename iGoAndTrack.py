import tkinter as tk
from timer_module import TimerApp

def main():
    root = tk.Tk()
    app = TimerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
