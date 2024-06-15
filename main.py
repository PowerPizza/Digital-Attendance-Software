import tkinter
import os
from tkinter import *
from tkcalendar import Calendar
from datetime import datetime
import pandas as pd
from tkinter import messagebox
from tkinter import filedialog
import matplotlib.pyplot as plt


root = tkinter.Tk()
root.title("Attendance Manager")
root.geometry("500x400")
root.state("zoomed")
icon_ = PhotoImage(master=root, file="softwareIconPng.png")
root.iconphoto(False, icon_)

if "all_classes" not in os.listdir():
    os.mkdir("all_classes")
print("Starting software....")

# ----------------- CONSTANTS OR DRY FUNCTIONS ------------------
int_to_month = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN", 7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
# ----------------------------- END ----------------------

# ---------------------- BUTTON FUNCTIONS -----------------------

def onCreateClass():
    opt_create_class.config(state="disabled")
    create_class_win = Toplevel(root)
    create_class_win.transient(root)
    create_class_win.title("Create Class")
    create_class_win.geometry("350x200")
    create_class_win.resizable(False, False)

    def on_close_create_class():
        opt_create_class.config(state="normal")
        create_class_win.destroy()
    create_class_win.protocol("WM_DELETE_WINDOW", on_close_create_class)

    class_name_var = StringVar()

    fr1 = Frame(create_class_win)
    lbl_class_name = Label(fr1, text="Class Name", font=("Helvetica", 16))
    lbl_class_name.pack(side=LEFT)
    class_name_ent = Entry(fr1, textvariable=class_name_var, font=("Helvetica", 14))
    class_name_ent.pack(side=LEFT, fill=X)
    fr1.pack(fill=X, pady=5)

    def on_save_class():
        try:
            open(f"all_classes/{class_name_var.get()}.csv", "rb").close()
            messagebox.showerror("Failed", "Class creation failed, class already exists!")
        except BaseException:
            raw_df = pd.DataFrame(columns=["Student Names"])
            raw_df.to_csv(f"all_classes/{class_name_var.get()}.csv", index=False)
            on_close_create_class()
            messagebox.showinfo("Done", "Class creation successful!")
    btn_save_class = Button(create_class_win, text="Save", bg="#d1f5b0", font=("Helvetica", 12), command=on_save_class)
    btn_save_class.pack(fill=X, padx=20, pady=10)
    create_class_win.mainloop()

def onTakeAttendance():
    if not len(os.listdir("all_classes")):
        messagebox.showwarning("Empty Class List", "No class currently exists please create a class first.")
        return

    root.withdraw()
    selected_class = StringVar(root, "Select Class")
    selected_date = datetime.now().date()
    df_of_class = pd.DataFrame()

    def on_back():
        if messagebox.askyesno("Back", "Have you saved changes?"):
            root.deiconify()
            root.state("zoomed")
            attendance_win.destroy()

    attendance_win = Toplevel()
    attendance_win.title("Take Attendance")
    attendance_win.geometry("1000x600")
    attendance_win.state("zoomed")

    attendance_win.protocol("WM_DELETE_WINDOW", on_back)

    main_canvas = Canvas(attendance_win, bg="#FFFFFF")

    header_opt_frame = Frame(main_canvas, bg="#FFFFFF", highlightbackground="#000000", highlightthickness=2)

    def onSelectClass(*eve):
        if selected_class.get() == "Select Class":
            return
        else:
            nonlocal df_of_class
            df_of_class = pd.read_csv(f"all_classes/{selected_class.get()}.csv")

    menu_select_class = OptionMenu(header_opt_frame, selected_class, *map(lambda x: x.split('.')[0], os.listdir("all_classes")))
    menu_select_class.config(font=('Helvetica', 16))
    menu_select_class.pack(padx=20, pady=20, side=LEFT)
    menu_select_class.bind("<Configure>", onSelectClass)

    def on_select_date():
        date_select_win = Toplevel(attendance_win)
        date_select_win.transient(attendance_win)
        date_select_win.title("Select Date")

        cal_1 = Calendar(date_select_win, year=selected_date.year, month=selected_date.month, day=selected_date.day)
        cal_1.pack()
        def on_date_change(*eve):
            nonlocal selected_date
            selected_date = cal_1.selection_get()
            lbl_select_date.config(text=str(selected_date))
            date_select_win.destroy()
        cal_1.bind("<<CalendarSelected>>", on_date_change)
        date_select_win.mainloop()

    frame_date_select = Frame(header_opt_frame, bg="#FFFFFF")
    opt_select_date = Button(frame_date_select, text="Select Date", font=('Helvetica', 16), command=on_select_date)
    opt_select_date.pack()
    lbl_select_date = Label(frame_date_select, text=str(selected_date), bg="#FFFFFF", font=('Helvetica', 14, "bold"), fg="#5c5c5c")
    lbl_select_date.pack()
    frame_date_select.pack(padx=20, pady=20, side=LEFT)

    def load_dataframe_as_table():
        fr_table_ = Frame(frame_take_attendance, bg="#FFFFFF")

        lbl_1 = Label(fr_table_, text="Student Names", font=("Helvetica", 16, "bold"), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
        lbl_1.grid(row=0, column=0)
        lbl_2 = Label(fr_table_, text="Present", font=("Helvetica", 16, "bold"), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
        lbl_2.grid(row=0, column=1)
        lbl_3 = Label(fr_table_, text="Absent", font=("Helvetica", 16, "bold"), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
        lbl_3.grid(row=0, column=2)
        lbl_4 = Label(fr_table_, text="Skip", font=("Helvetica", 16, "bold"), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
        lbl_4.grid(row=0, column=3)

        def create_entry_row(text_):
            selected_APS_var = IntVar()
            row_idx_cpy = row_idx-1

            if str(selected_date) in df_of_class.columns:
                if df_of_class.loc[row_idx_cpy, str(selected_date)] == "P":
                    selected_APS_var.set(1)
                elif df_of_class.loc[row_idx_cpy, str(selected_date)] == "A":
                    selected_APS_var.set(2)
                elif df_of_class.loc[row_idx_cpy, str(selected_date)] == "SKIP":
                    selected_APS_var.set(3)

            def onSelect():
                if selected_APS_var.get() == 1:
                    df_of_class.loc[row_idx_cpy, str(selected_date)] = "P"
                elif selected_APS_var.get() == 2:
                    df_of_class.loc[row_idx_cpy, str(selected_date)] = "A"
                elif selected_APS_var.get() == 3:
                    df_of_class.loc[row_idx_cpy, str(selected_date)] = "SKIP"

            lbl_std_name = Label(fr_table_, text=text_, font=("Helvetica", 16), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
            lbl_std_name.grid(row=row_idx, column=0, sticky=NSEW)
            present_chk = Radiobutton(fr_table_, value=1, variable=selected_APS_var, command=onSelect, font=("Helvetica", 16), bg="#bdffce", highlightthickness=1, highlightbackground="#000000")
            present_chk.grid(row=row_idx, column=1, ipady=5, sticky=EW)
            absent_chk = Radiobutton(fr_table_, value=2, variable=selected_APS_var, command=onSelect, font=("Helvetica", 16), bg="#ffc6bd", highlightthickness=1, highlightbackground="#000000")
            absent_chk.grid(row=row_idx, column=2, ipady=5, sticky=EW)
            skip_chk = Radiobutton(fr_table_, value=3, variable=selected_APS_var, command=onSelect, font=("Helvetica", 16), bg="#9e9e9e", highlightthickness=1, highlightbackground="#000000")
            skip_chk.grid(row=row_idx, column=3, ipady=5, sticky=EW)

        row_idx = 1
        for student in df_of_class["Student Names"].values:
            create_entry_row(student)
            row_idx += 1

        fr_table_.bind("<Configure>", lambda e: frame_take_attendance.config(scrollregion=frame_take_attendance.bbox("all")))
        frame_take_attendance.create_window((0, 0), window=fr_table_, anchor="nw")

    def on_click_go():
        nonlocal df_of_class

        if str(selected_date) in df_of_class.columns:
            if not messagebox.askyesno("Already Taken", "Attendance of this date is already taken do you want to make some changes ?"):
                return

        load_dataframe_as_table()
        def on_scroll_wheel(*eve):
            if eve[0].delta < 0:
                frame_take_attendance.yview_scroll(1, "units")
            else:
                frame_take_attendance.yview_scroll(-1, "units")

        attendance_win.bind("<MouseWheel>", on_scroll_wheel)

    btn_go = Button(header_opt_frame, text="Go!", font=('Helvetica', 16), command=on_click_go)
    btn_go.pack(padx=20, pady=20, side=RIGHT)

    header_opt_frame.pack(fill=X)

    frame_take_attendance = Canvas(main_canvas, bg="#FFFFFF")
    scroll_ = Scrollbar(frame_take_attendance, orient="vertical", command=frame_take_attendance.yview)
    frame_take_attendance.configure(yscrollcommand=scroll_.set)
    scroll_.pack(side="right", fill="y")
    frame_take_attendance.pack(fill=BOTH, expand=True, side=LEFT)

    frame_other_options = Frame(main_canvas, bg="#fff9e3")
    lbl_options = Label(frame_other_options, text="Options", font=("Helvetica", 20), bg="#ff6969")
    lbl_options.pack(fill=X, ipadx=50)
    fr_func_opts = Frame(frame_other_options, bg="#fff9e3")

    def on_add_student():
        std_name_var = StringVar()
        add_std_win = Toplevel(attendance_win)
        add_std_win.transient(attendance_win)
        add_std_win.title("Add Student")
        add_std_win.geometry("300x150")

        fr1 = Frame(add_std_win)
        lbl_name_entry = Label(fr1, text="Name", font=("Helvetica", 16))
        lbl_name_entry.pack(side=LEFT, padx=2)
        name_entry = Entry(fr1, textvariable=std_name_var, highlightthickness=1, highlightbackground="#000000", font=("Helvetica", 14))
        name_entry.pack(side=LEFT, fill=X, expand=True)
        fr1.pack(fill=X)

        def on_save_name():
            if std_name_var.get() not in df_of_class["Student Names"]:
                df_of_class.loc[len(df_of_class.index), "Student Names"] = std_name_var.get()
                df_of_class.to_csv(f"all_classes/{selected_class.get()}.csv", index=False)
                load_dataframe_as_table()
                add_std_win.destroy()
            else:
                messagebox.showerror("Already exists", "Given name is already exists.")
        btn_save_name = Button(add_std_win, text="Save", bg="#d1f5b0", command=on_save_name)
        btn_save_name.pack(ipadx=20, anchor=NE)

        add_std_win.mainloop()
    add_student_btn = Button(fr_func_opts, text="Add Student +", font=("Helvetica", 16), relief="ridge", bg="#d2ffc2", command=on_add_student)
    add_student_btn.pack(padx=20, fill=X, pady=2)

    def onPlotTodayGraph():
        if str(datetime.now().date()) in df_of_class.columns:
            presents = len(df_of_class[df_of_class[str(datetime.now().date())] == "P"])
            absents = len(df_of_class[df_of_class[str(datetime.now().date())] == "A"])
            skips = len(df_of_class[df_of_class[str(datetime.now().date())] == "SKIP"])

            plt.figure("Today's attendance graph")
            bars = plt.bar(["Present", "Absent", "Skip"], [presents, absents, skips])
            bars[0].set_color("green")
            bars[1].set_color("red")
            bars[2].set_color("gray")
            plt.annotate(f"{presents}/{len(df_of_class['Student Names'])}", (bars[0].get_x() + bars[0].get_width()/4, bars[0].get_height()+0.03))
            plt.annotate(f"{absents}/{len(df_of_class['Student Names'])}", (bars[1].get_x() + bars[1].get_width()/4, bars[1].get_height()+0.03))
            plt.annotate(f"{skips}/{len(df_of_class['Student Names'])}", (bars[2].get_x() + bars[2].get_width()/4, bars[2].get_height()+0.03))
            plt.title("Today's Attendance Graph")
            plt.xlabel("Present/Absent/Skip")
            plt.ylabel("No. of students")
            plt.show()
        else:
            messagebox.showerror("Can't Plot", "Attendance of today is not taken yet!")
    graph_today_btn = Button(fr_func_opts, text="Plot Graph (Today's)", font=("Helvetica", 16), relief="ridge", bg="#d2ffc2", command=onPlotTodayGraph)
    graph_today_btn.pack(padx=20, fill=X, pady=2)

    def onPlotMonthlyGraph():
        df_monthly = pd.DataFrame(df_of_class["Student Names"], columns=["Student Names"])
        for dates_ in df_of_class.columns[1:]:
            if datetime.strptime(dates_, "%Y-%m-%d").date().month == datetime.now().month:
                df_monthly[dates_] = df_of_class[dates_]

        df_counts = pd.DataFrame(columns=["presentCount", "absentCount", "skipCount"])
        for x in range(len(df_monthly.index)):
            present_count = len(df_monthly.iloc[x][df_monthly.iloc[x] == "P"])
            absent_count = len(df_monthly.iloc[x][df_monthly.iloc[x] == "A"])
            skip_count = len(df_monthly.iloc[x][df_monthly.iloc[x] == "SKIP"])
            df_counts.loc[df_monthly.loc[x, "Student Names"]+f" {x+1}"] = [present_count, absent_count, skip_count]
        df_counts.plot(kind="bar")
        plt.xlabel("Students")
        plt.ylabel("Days")
        plt.xticks(rotation=270)
        plt.title("This month's attendance Graph")
        plt.show()
    graph_month_btn = Button(fr_func_opts, text="Plot Graph (This Month's)", font=("Helvetica", 16), relief="ridge", bg="#d2ffc2", command=onPlotMonthlyGraph)
    graph_month_btn.pack(padx=20, fill=X, pady=2)

    def onPlotYearlyGraph():
        month_wise_columns = {1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [], 11: [], 12: []}
        for dates_ in df_of_class.columns[1:]:
            if datetime.strptime(dates_, "%Y-%m-%d").date().year == datetime.now().year:
                month_wise_columns[datetime.strptime(dates_, "%Y-%m-%d").date().month].append(dates_)

        df_yearly_attendance = pd.DataFrame(index=list(int_to_month.values()), columns=["presentCount", "absentCount", "skipCount"])
        for month in month_wise_columns:
            if len(month_wise_columns[month]):
                present_count = sum(df_of_class[df_of_class[month_wise_columns[month]] == "P"].count())
                absent_count = sum(df_of_class[df_of_class[month_wise_columns[month]] == "A"].count())
                skip_count = sum(df_of_class[df_of_class[month_wise_columns[month]] == "SKIP"].count())
                df_yearly_attendance.loc[int_to_month[month]] = [present_count, absent_count, skip_count]
        df_yearly_attendance.plot(kind="bar")
        plt.title("This year's attandance graph")
        plt.xlabel("Months")
        plt.ylabel("Attendance count")
        plt.show()
    graph_year_btn = Button(fr_func_opts, text="Plot Graph (This Year's)", font=("Helvetica", 16), relief="ridge", bg="#d2ffc2", command=onPlotYearlyGraph)
    graph_year_btn.pack(padx=20, fill=X, pady=2)

    def on_save():
        if not df_of_class.empty:
            df_of_class.to_csv(f"all_classes/{selected_class.get()}.csv", index=False)
            messagebox.showinfo("Saved", f"Attendance of {selected_date} saved successfully.")
    save_btn = Button(fr_func_opts, text="Save", font=("Helvetica", 16), relief="ridge", bg="#d2ffc2", command=on_save)
    save_btn.pack(padx=20, fill=X, pady=2)

    def on_export():
        if df_of_class.empty:
            return
        file_ = filedialog.asksaveasfilename(filetypes=[("excel file", "*.xlsx"), ("csv file", "*.csv")], defaultextension="csv")
        if file_:
            if file_.split(".")[-1] == "xlsx":
                df_of_class.to_excel(file_, index=False)
            elif file_.split(".")[-1] == "csv":
                df_of_class.to_csv(file_, index=False)
            messagebox.showinfo("File Saved", "Export successful!")
    export_btn = Button(fr_func_opts, text="Export", font=("Helvetica", 16), relief="ridge", bg="#d2ffc2", command=on_export)
    export_btn.pack(padx=20, fill=X, pady=2)

    back_btn = Button(fr_func_opts, text="Back", font=("Helvetica", 16), relief="ridge", bg="#ffca9e", command=on_back)
    back_btn.pack(padx=20, fill=X, pady=2)

    fr_func_opts.pack(side=LEFT, fill=Y)
    frame_other_options.pack(fill=BOTH, side=LEFT)

    main_canvas.pack(fill=BOTH, expand=True)
    attendance_win.mainloop()


def onClickAttendanceReg():
    if not len(os.listdir("all_classes")):
        messagebox.showwarning("Empty Class List", "No class currently exists please create a class first.")
        return

    selected_class = StringVar(root, "Select Class")
    df_selected_class = pd.DataFrame()

    root.withdraw()
    attendance_reg_win = Toplevel()
    attendance_reg_win.title("Attendance Register")
    attendance_reg_win.geometry("1000x600")
    attendance_reg_win.state("zoomed")

    def on_back():
        root.deiconify()
        root.state("zoomed")
        attendance_reg_win.destroy()
    attendance_reg_win.protocol("WM_DELETE_WINDOW", on_back)

    fr_table_ = None  # to clear all data of table destroy and recreate it...
    def load_dataframe_as_table():
        nonlocal fr_table_
        if fr_table_:
            fr_table_.destroy()
        fr_table_ = Frame(table_canvas, bg="#FFFFFF")

        for x in range(len(df_selected_class.columns)):
            lbl_1 = Label(fr_table_, text=df_selected_class.columns[x], font=("Helvetica", 16, "bold"), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
            lbl_1.grid(row=0, column=x)

        for x in range(len(df_selected_class.index)):
            col_idx = 0
            for col in df_selected_class.loc[x]:
                lbl_std_name = Label(fr_table_, text=col, font=("Helvetica", 16), bg="#FFFFFF", highlightthickness=1, highlightbackground="#000000")
                lbl_std_name.grid(row=x+1, column=col_idx, sticky=NSEW)
                col_idx += 1

        fr_table_.bind("<Configure>", lambda e: table_canvas.config(scrollregion=table_canvas.bbox("all")))
        table_canvas.create_window((0, 0), window=fr_table_, anchor="nw")

    register_main_canvas = Canvas(attendance_reg_win, bg="#FFFFFF")

    def onClassSelect(*eve):
        if selected_class.get() == "Select Class":
            return
        else:
            nonlocal df_selected_class
            df_selected_class = pd.read_csv(f"all_classes/{selected_class.get()}.csv")
            load_dataframe_as_table()

    fr_top_options = Frame(register_main_canvas, bg="#FFFFFF")
    back_btn = Button(fr_top_options, text="Back", font=("Helvetica", 16), relief="ridge", bg="#ffca9e", command=on_back)
    back_btn.pack(padx=2, pady=20, side=LEFT)

    menu_select_class = OptionMenu(fr_top_options, selected_class, *map(lambda x: x.split('.')[0], os.listdir("all_classes")))
    menu_select_class.config(font=('Helvetica', 16))
    menu_select_class.pack(padx=20, pady=20, side=LEFT)
    menu_select_class.bind("<Configure>", onClassSelect)
    fr_top_options.pack(fill=X)

    fr_register_canvas = Canvas(register_main_canvas, bg="#FFFFFF")
    fr_table_n_yscroll = Frame(fr_register_canvas)
    table_canvas = Canvas(fr_table_n_yscroll, bg="#FFFFFF")
    table_canvas.pack(fill=BOTH, expand=True, side=LEFT)

    scroll_y = Scrollbar(fr_table_n_yscroll, orient="vertical", command=table_canvas.yview)
    table_canvas.configure(yscrollcommand=scroll_y.set)
    scroll_y.pack(side=RIGHT, fill=Y)
    fr_table_n_yscroll.pack(fill=BOTH, expand=True)

    scroll_x = Scrollbar(fr_register_canvas, orient="horizontal", command=table_canvas.xview)
    table_canvas.configure(xscrollcommand=scroll_x.set)
    scroll_x.pack(side=BOTTOM, fill=X)
    fr_register_canvas.pack(fill=BOTH, expand=True)

    register_main_canvas.pack(fill=BOTH, expand=True)
    attendance_reg_win.mainloop()

# ----------------------------- END ----------------------

canvas_home = Canvas(root, bg="#b3d5ff")
lbl1 = Label(canvas_home, text="Welcome", font=("Helvetica", 28, "bold"), fg="#512eff", bg="#b3d5ff")
lbl1.pack(pady=5)

frame_btns = Frame(canvas_home, bg="#fff4e0")
opt_create_class = Button(frame_btns, text="Create Class", font=("Helvetica", 18), bg="#8a78ff", fg="#feffd6", activebackground="#8a78ff", activeforeground="#feffd6", relief="groove", command=onCreateClass)
opt_create_class.pack(padx=20, pady=10, fill=X)
opt_take_attendance = Button(frame_btns, text="Take Attendance", font=("Helvetica", 18), bg="#8a78ff", fg="#feffd6", activebackground="#8a78ff", activeforeground="#feffd6", relief="groove", command=onTakeAttendance)
opt_take_attendance.pack(padx=20, pady=10, fill=X)
opt_attendance_reg = Button(frame_btns, text="Attendance Register", font=("Helvetica", 18), bg="#8a78ff", fg="#feffd6", activebackground="#8a78ff", activeforeground="#feffd6", relief="groove", command=onClickAttendanceReg)
opt_attendance_reg.pack(padx=20, pady=10, fill=X)

def onClickAbout():
    messagebox.showinfo("About", "Project Name : Attendance Manager\nDescription : Its a prototype software which can be use to automate and digitalize the attendance work of schools.\nMade By : scihack/powerpizza")
opt_about = Button(frame_btns, text="About", font=("Helvetica", 18), bg="#8a78ff", fg="#feffd6", activebackground="#8a78ff", activeforeground="#feffd6", relief="groove", command=onClickAbout)
opt_about.pack(padx=20, pady=10, fill=X)
frame_btns.pack(pady=40)
canvas_home.place(relx=0, rely=0, relwidth=1, relheight=1)

root.mainloop()