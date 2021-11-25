import easygui as g

g.msgbox(msg="欢迎使用成绩整理与分析系统    copyright©HazelNut", title="成绩整理与分析系统")
    
while 1:
    msg ="选择操作类型"
    title = "成绩整理与分析系统"
    choices = ["数据汇总", "数据分析","退出"]
    choice = g.buttonbox(msg, title, choices)

    if choice == choices[0]:
        import getres
        g.msgbox("数据汇总完成！")
    elif choice == choices[1]:
        import paint_gui
    else:
        sys.exit(0)           # user chose Cancel