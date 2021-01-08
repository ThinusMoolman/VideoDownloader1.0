from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from pytube import YouTube
import auto_py_to_exe

Folder_Path = ""


def openFileLocation():
    global Folder_Path
    Folder_Path = filedialog.askdirectory()
    if len(Folder_Path) > 1:
        pathError.config(text=Folder_Path, fg="green")

    else:
        pathError.config(text="Please Choose Folder!!", fg="red")


def bar():
    import time
    progress = ttk.Progressbar(root, orient=HORIZONTAL, length=100, mode='determinate')
    progress['value'] = 20
    root.update_idletasks()
    time.sleep(1)

    progress['value'] = 40
    root.update_idletasks()
    time.sleep(1)

    progress['value'] = 50
    root.update_idletasks()
    time.sleep(1)

    progress['value'] = 60
    root.update_idletasks()
    time.sleep(1)

    progress['value'] = 80
    root.update_idletasks()
    time.sleep(1)
    progress['value'] = 100
    progress.grid()


def DownloadVideo():
    choice = chooseChoices.get()
    url = UrlEntry.get()

    if len(url) > 1:
        UrlNotFoundError.config(text="")
        yt = YouTube(url)

        if choice == choices[0]:
            bar()
            select = yt.streams.filter(progressive=True).first()

        elif choice == choices[1]:
            bar()
            select = yt.streams.filter(progressive=True, file_extension='mp4').last()

        elif choice == choices[2]:
            bar()
            select = yt.streams.filter(only_audio=True).first()

        else:
            UrlNotFoundError.config(text="Paste Link again!!", fg="red")

    # download function
    select.download(Folder_Path)

    UrlNotFoundError.config(text="Download Completed!", fg="Green")


root = Tk()
root.title("Youtube Video Downloader")
root.geometry("390x400")
root.config(bg="grey")
root.columnconfigure(0, weight=1)

EnterUrlLabel = Label(root, text="Enter the URL of the Video", font=("Agency FB", 25), fg="white", bg="grey")
EnterUrlLabel.grid()

EnterUrl = StringVar()
UrlEntry = Entry(root, width=50, textvariable=EnterUrl, bg="black", fg="white")
UrlEntry.grid()

UrlNotFoundError = Label(root, text="Url not found", fg="red", font=("Agency FB", 20), bg="grey")
UrlNotFoundError.grid()

saveLabel = Label(root, text="Save the Video File", font=("Agency FB", 25, "bold"), bg="grey", fg="white")
saveLabel.grid()

saveEntry = Button(root, width=20, bg="Black", fg="white", text="Choose Path", command=openFileLocation)
saveEntry.grid()

pathError = Label(root, text="Path not found", fg="red", font=("Agency FB", 20), bg="grey")
pathError.grid()

SelectQuality = Label(root, text="Select Quality", font=("Agency FB", 25), bg="grey", fg="white")
SelectQuality.grid()

choices = ["720p", "144p", "Only Audio"]
chooseChoices = ttk.Combobox(root, values=choices)
chooseChoices.grid()

DownloadButton = Button(root, text="Download", width=20, bg="Black", fg="white", command=DownloadVideo)
DownloadButton.grid()

MadeBy = Label(root, text="Made by Thinus Moolman", font=("Agency FB", 25), bg="grey", fg="white")
MadeBy.grid()
root.mainloop()
