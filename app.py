from file_model.file_model import FileModel

if __name__ == '__main__':
    FileModel().formats(print_title="SPEED CHOICE")
    print("You have two options.\n"
          "Either go (f)ast which will:\n\n"
          "\t1. Load each workbook\n"
          "\t2. Assume the first column (A) contains your index column\n"
          "\t3, Will iterate over every single sheet and spit out the desired textfile\n\n"
          "\nOr, you can go (s)low which will enable you to select specific sheets and set any column you choose as the "
          "index")
    choice = str(input("Do you want to go (f)ast or (s)low?: "))
    while choice not in 'f F s S'.split():
        choice = str(input("Just type 'f' for fast or 's' for slow: "))
    if choice == 'f':
        FileModel().create_dataframes_fast()
    elif choice == 's':
        FileModel().create_dataframes()