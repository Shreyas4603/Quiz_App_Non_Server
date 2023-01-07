import time

import customtkinter as ct

from tkinter import filedialog
import shutil
import os

# colour pallete
bg = "#24293E"
accent = "#8EBBFF"
text = "#F4F5FC"
secondary = "#434b6e"

ct.set_appearance_mode("dark")

create_window = ct.CTk()
width = 450
height = 600
create_window.geometry("450x350")
create_window.resizable(0, 0)
create_window.title("Create quiz")

create_window.configure(fg_color=bg)
filename_arr = [""]


def check_folder():  # to check if questions and profile folder exist
    global p_path
    path = os.getcwd()  # current directory we r working in
    s_path = path + "\Questions"
    p_path = path + "\Profiles"
    if (os.path.exists(s_path) and os.path.exists(p_path)):
        print("Exist")
        return
    else:
        os.mkdir(s_path)
        os.mkdir(p_path)


def check_exist_profile():
    check_folder()
    global c  # to make sure file is selected ,flag var
    c = True
    print("p path", p_path)
    print("ext = ", ext)
    c = ext.isalpha()

    pth = p_path + "\\" + profile_name.get() + "." + ext
    print("Path profile check = ", pth)
    check = os.path.exists(pth)
    if c == False:
        print("c = ", c)
        file_name.configure(text="Select file")
    elif check == True:
        file_name.configure(text="Profile exist")

    return check


def create_profile(prof_name):
    # print("name = ",prof_name.get())
    global ext

    # check_folder()  # check if question n profile exist
    try:
        source = filename_arr[0]
        ext = source.split(".")[-1]
        p_e = check_exist_profile()
        print("Profile exist ", p_e)
        if (p_e == False and c == True):
            f_name = source.split("Questions/")[-1]
            print(f_name)

            destination = "E:\Quiz App\Profiles\\"
            shutil.copy2(source, destination)

            src = destination + f_name
            dst = destination + str(prof_name.get()) + "." + ext

            print("src = ", src)
            print("dst = ", dst)

            os.rename(src, dst)
    except Exception as e:
        file_name.configure(text=e)
    if c == True:
        file_name.configure(text="Profile complete")
        create_window.update()
        time.sleep(1)
        create_window.destroy()


def open_file():
    filename = ""
    filename = filename + str(filedialog.askopenfilename())
    filename_arr[0] = filename  # stores directory address of the question file selected
    # print(filename)
    # file=filedialog.askdirectory(title="Select image files")
    name = filename.split("/")[-1]
    file_name.configure(text=name)
    print(filename)




# UI


create_quiz = ct.CTkLabel(master=create_window,
                          text="Create Quiz", font=("Momcake", 60, "bold"), fg_color=secondary, bg_color="transparent",
                          text_color=text,
                          corner_radius=8)
create_quiz.pack(pady=10)  # create quiz title

enter_p_name = ct.CTkLabel(create_window, text="Enter profile name", font=("Gill Sans MT", 20))
enter_p_name.place(x=148, y=126)

profile_name = ct.StringVar()
profile_entry = ct.CTkEntry(master=create_window, textvariable=profile_name, font=("Consolas", 20), width=300,
                            border_width=3, border_color=secondary, justify="left")
profile_entry.pack(pady=(80, 10))

select_file = ct.CTkButton(master=create_window, text="Open file", font=("Gill Sans MT", 20), corner_radius=5, command=open_file,
                           fg_color=secondary, width=100)
select_file.place(x=76, y=201)

file_name = ct.CTkLabel(master=create_window, text="", font=("Gill Sans MT", 20), corner_radius=5, width=192, height=35)
file_name.place(x=180, y=201)

done = ct.CTkButton(master=create_window, text="Create", font=("Gill Sans MT", 25), corner_radius=30, width=104,
                    command=lambda :create_profile(profile_name), fg_color="#2c8c00", hover_color="#2eab49")
done.place(x=182, y=243)



create_window.mainloop()
