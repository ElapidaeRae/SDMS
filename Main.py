import datetime, openpyxl, os, sv_ttk
from tkinter import *
from tkinter import messagebox
from tkinter import ttk

# import pandas as pd
# import matplotlib.pyplot as plt
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

Users = {'Adam': ['Adam', 'Adam'],
         'Kelly': ['Kelly', 'K3lly'],
         'DaisyS89': ['P34Ches', 'Daisy'],
         'DaveN23': ['d4v3', 'Dave'],
         'Boss': ['ssoB', 'Boss']}

CarBrands = ['Abarth', 'Alfa Romeo', 'Alpine', 'Aston Martin', 'Audi', 'BMW', 'Bentley', 'Bugatti', 'Cadillac',
             'Chevrolet', 'Chrysler', 'Citroën', 'Cupra', 'DS', 'Dacia', 'Ferrari', 'Fiat', 'Ford', 'Genesis',
             'Honda', 'Hyundai', 'Infiniti', 'Isuzu', 'Iveco', 'Jaguar', 'Jeep', 'Kia', 'Lada', 'Lamborghini',
             'Land Rover', 'Lexus', 'Lotus', 'MG', 'Maserati', 'Mazda', 'McLaren', 'Mercedes-Benz', 'Mini',
             'Mitsubishi', 'Nissan', 'Peugeot', 'Polestar', 'Porsche', 'Renault', 'Rolls-Royce', 'SAAB', 'Seat',
             'Smart', 'Ssangyong', 'Subaru', 'Suzuki', 'Tesla', 'Toyota', 'Vauxhall', 'Volkswagen', 'Volvo', 'Škoda']

loggedInUser = ''


def loginSuccess(user):
    """
The function that runs upon login success.
loggedInUser is global as it is used elsewhere.
this is very janky.
    :param user:
    """
    global loggedInUser
    root.deiconify()
    root.focus()
    loginW.withdraw()
    loggedInUser = Users.get(user)[1]


def login(username, password):
    """
Takes the inputted username and password in the login window and checks them against the Users dictionary.
messageboxes are used to convey login failure.
    :param username:
    :param password:
    """
    if Users.get(username.get()) is not None:
        if Users.get(username.get())[0] == password.get():
            loginSuccess(username.get())
        else:
            messagebox.showerror(title='Login Failure', message='Incorrect Credentials')
    else:
        messagebox.showerror(title='Login Failure', message='Incorrect Credentials')


def CurrentDate(day, month, year):
    """
Sets the strings in the spinboxes to the current date with datetime.
Quite self-explanatory, but very useful.
    :param day:
    :param month:
    :param year:
    """
    today = datetime.date.today()
    day.set(today.day)
    month.set(today.month)
    year.set(today.year)


def subwindowclose(window):
    """
Used to replace the default window close protocol.
    :param window:
    """
    window.destroy()
    root.deiconify()


windowBG = '#1c1c1c'
# This makes a folder in the local %appdata% folder to put the Excel file in
carsFilePath = f'{os.getenv("APPDATA")}/SDMS/Cars.xlsx'
dirFilePath = f'{os.getenv("APPDATA")}/SDMS/'
# If the folder doesn't exist, make it
if not os.path.exists(carsFilePath):
    os.makedirs(dirFilePath)
    carsData = openpyxl.Workbook()
    dataSheet = carsData.active
    header = ['Registration', 'Brand', 'Colour', 'Buy Price', 'Buy Date', 'Buyer', 'Sale Price', 'Sale Date', 'Seller',
              'Margin', 'Buyer Compensation', 'Seller Compensation']
    dataSheet.append(header)
    carsData.save(carsFilePath)
    messagebox.showinfo(title='File Created', message='Excel Data File Created')


# Buyer Window


def buyerwindowstart():
    """
Making the buyer window. Defined as a function so that it can be linked to a button and new instances
of it can be created after it is destroyed.
    """
    # Creates the window and defines settings before withdrawing the root window
    buyerWindow = Toplevel(root)
    buyerWindow.geometry('700x400')
    buyerWindow.config(bg=windowBG)
    buyerWindow.resizable(FALSE, FALSE)
    buyerWindow.title('Log A Bought Car')
    root.withdraw()
    buyerWindow.focus()
    # frame for the frames to go in
    buyerFrame = Frame(buyerWindow, bg=windowBG)
    # frame for the entries to go in
    buyCarInfo = ttk.LabelFrame(buyerFrame, text='Car Information')
    # the buyBrand Combobox takes its list from the CarBrands list at the top
    buyBrand = ttk.Combobox(buyCarInfo, values=CarBrands, state='readonly')
    buyBrandLabel = ttk.Label(buyCarInfo, text='Car Brand', )
    buyColour = ttk.Entry(buyCarInfo)
    buyColourLabel = ttk.Label(buyCarInfo, text='Car Colour')
    buyBrand.grid(row=0, column=1, padx=10)
    buyBrandLabel.grid(row=0, column=0, padx=10, pady=20)
    buyColour.grid(row=0, column=3, padx=10)
    buyColourLabel.grid(row=0, column=2, padx=10)

    # every entry is paired with a label telling the user what goes where
    buyReg = ttk.Entry(buyCarInfo)
    buyRegLabel = ttk.Label(buyCarInfo, text='Registration')
    buyReg.grid(row=1, column=1, padx=10, sticky='ew')
    buyRegLabel.grid(row=1, column=0, padx=10, pady=20)

    buyCarInfo.grid(row=0, column=0, padx=20, pady=20, sticky='news')
    buyBuyerInfo = ttk.LabelFrame(buyerFrame, text='Buyer Information')

    buyPrice = ttk.Entry(buyBuyerInfo)
    buyPriceLabel = ttk.Label(buyBuyerInfo, text='Buy Price')
    buyPrice.grid(row=0, column=1, padx=10, pady=20)
    buyPriceLabel.grid(row=0, column=0, padx=10, pady=20)

    buyDateLabel = ttk.Label(buyBuyerInfo, text='Buy Date', )
    buyDateDay = ttk.Spinbox(buyBuyerInfo, width=6, from_=1, to=31)
    buyDateMonth = ttk.Spinbox(buyBuyerInfo, width=5, from_=1, to=12)
    buyDateYear = ttk.Spinbox(buyBuyerInfo, width=5, from_=1980, to=datetime.date.today().year)
    # this uses the CurrentDate function defined above
    buyDateCurrent = ttk.Button(buyBuyerInfo, text='Use Current Date',
                                command=CurrentDate(day=buyDateDay, month=buyDateMonth, year=buyDateYear))
    buyDateLabel.grid(row=1, column=0, padx=10, pady=10)
    buyDateDay.grid(row=1, column=1, padx=5, pady=10)
    buyDateMonth.grid(row=1, column=2, padx=5)
    buyDateYear.grid(row=1, column=3, padx=5)
    buyDateCurrent.grid(row=0, column=3)
    buyBuyerInfo.grid(row=1, column=0, padx=20, pady=20, sticky='news')
    # lambda makes the command run asynchronously allowing it to work within the mainloop
    buyConfirm = ttk.Button(buyerFrame, text='Confirm',
                            command=lambda: buyerLogData(buyReg, buyBrand, buyColour, buyPrice, buyDateDay,
                                                         buyDateMonth, buyDateYear, buyerWindow))
    buyConfirm.grid(row=2, column=0, sticky='ew')

    buyerFrame.pack()
    # the window closure protocol is replaced so when the window closes the root window reappears
    buyerWindow.protocol('WM_DELETE_WINDOW', lambda: subwindowclose(buyerWindow))


def buyerLogData(reg, brand, colour, buyprice, buyday, buymonth, buyyear, window):
    """
Takes all the parameters below that have been entered into the buyer window and logs them in the spreadsheet
    :param reg:
    :param brand:
    :param colour:
    :param buyprice:
    :param buyday:
    :param buymonth:
    :param buyyear:
    :param window:
    """
    # opening the Excel spreadsheet
    carWorkbook = openpyxl.load_workbook(carsFilePath)
    dataSheet = carWorkbook.active
    # combining the outputs from the date spinboxes to be in the format yyyy-mm-dd
    buyDate = f'{buyyear.get()}-{buymonth.get()}-{buyday.get()}'
    # assembling the data into a list to be written as one row
    datalist = [reg.get().capitalize(), brand.get(), colour.get().upper(), int(buyprice.get()), buyDate, loggedInUser]
    # print(f'1: {reg.get()[4]}')
    # print(f'2: {len(reg.get())}')
    # entry validation so that nothing is left empty
    if reg.get() is not None or brand.get() is not None or colour.get() is not None or buyprice.get() is not None:
        # checking the registration is valid, no custom plates though
        if reg.get()[4] == ' ' and len(reg.get()) == 8:
            dataSheet.append(datalist)
            carWorkbook.save(carsFilePath)
            messagebox.showinfo(title='Car Logged',
                                message=f'Your {reg.get().capitalize()} plate {brand.get()} has been successfully logged.')

            subwindowclose(window)
        else:
            messagebox.showerror(title='Formatting error', message='Registration entered is either '
                                                                   'invalid or a custom plate.')
    else:
        messagebox.showerror(title='Data Not Entered', message='Not all fields have been filled.')


# Seller Window


def sellerWindowStart():
    """
Making the seller window
    """
    sellerWindow = Toplevel(root)
    sellerWindow.geometry('580x260')
    sellerWindow.config(bg=windowBG)
    sellerWindow.resizable(FALSE, FALSE)
    sellerWindow.title('Log A Sold Car')
    root.withdraw()
    sellerWindow.focus()

    sellCarInfo = ttk.LabelFrame(sellerWindow, text='Seller Information')
    sellRegLabel = ttk.Label(sellCarInfo, text='Registration')
    sellReg = ttk.Entry(sellCarInfo)
    sellPriceLabel = ttk.Label(sellCarInfo, text='Sell Price')
    sellPrice = ttk.Entry(sellCarInfo)
    sellDateLabel = ttk.Label(sellCarInfo, text='Sell Date')
    sellDateDay = ttk.Spinbox(sellCarInfo, width=6, from_=0, to=31)
    sellDateMonth = ttk.Spinbox(sellCarInfo, width=5, from_=0, to=12)
    sellDateYear = ttk.Spinbox(sellCarInfo, width=5, from_=1980, to=datetime.date.today().year)
    sellDateCurrent = ttk.Button(sellCarInfo, text='Use Current Date',
                                 command=lambda: CurrentDate(day=sellDateDay, month=sellDateMonth, year=sellDateYear))
    sellConfirm = ttk.Button(sellerWindow, text='Confirm',
                             command=lambda: sellerLogData(sellReg, sellPrice, sellDateDay, sellDateMonth, sellDateYear,
                                                           sellerWindow))
    sellRegLabel.grid(row=0, column=0, pady=10, padx=10)
    sellReg.grid(row=0, column=1, pady=10, padx=10)
    sellPriceLabel.grid(row=0, column=2, pady=10, padx=10)
    sellPrice.grid(row=0, column=3, pady=10, padx=10)
    sellDateLabel.grid(row=1, column=0, padx=5, pady=10)
    sellDateDay.grid(row=1, column=1, padx=5, pady=10)
    sellDateMonth.grid(row=1, column=2, padx=5, pady=10)
    sellDateYear.grid(row=1, column=3, padx=5, pady=10)
    sellDateCurrent.grid(row=2, column=2, padx=5, pady=10)
    sellCarInfo.grid(row=0, column=0, padx=10, pady=10)
    sellConfirm.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

    sellerWindow.protocol('WM_DELETE_WINDOW', lambda: subwindowclose(sellerWindow))



def sellerLogData(reg, price, sellday, sellmonth, sellyear, window):
    """
Takes the data from sellerWindow and logs it in the spreadsheet
    :param reg:
    :param price:
    :param sellday:
    :param sellmonth:
    :param sellyear:
    :param window:
    """
    carWorkbook = openpyxl.load_workbook(carsFilePath)
    datasheet = carWorkbook.active
    sellDate = f'{sellyear.get()}-{sellmonth.get()}-{sellday.get()}'
    data = [int(price.get()), sellDate, loggedInUser]
    # print(data)
    # Checking the registration column of the spreadsheet to see if it contains reg
    for cell in datasheet['A']:
        # checking if it's None or else it breaks
        if cell.value is not None:
            if reg.get() in cell.value:
                # if reg is found, write to the spreadsheet
                datasheet.cell(row=cell.row, column=7, value=data[0])
                datasheet.cell(row=cell.row, column=8, value=data[1])
                datasheet.cell(row=cell.row, column=9, value=data[2])
                buyprice = datasheet['D' + str(cell.row)].value
                margin = int(price.get()) - buyprice
                # calculates the buyer and seller's compensation to be written
                if margin > 0:
                    buyerComp = (margin / 100) * 2
                    sellerComp = (margin / 100) * 5
                else:
                    buyerComp = 0
                    sellerComp = 0
                dataSheet.cell(row=cell.row, column=10, value=margin)
                dataSheet.cell(row=cell.row, column=11, value=buyerComp)
                dataSheet.cell(row=cell.row, column=12, value=sellerComp)
                carWorkbook.save(carsFilePath)
                # notifies the user that the logging was successful
                messagebox.showinfo(title='Car Logged',
                                    message=f'Your {reg.get().capitalize()} plate has been successfully logged as sold for £{price.get()}.')
                subwindowclose(window)
    if reg.get() not in dataSheet['A']:
        messagebox.showerror(title='Registration Not Found', message='The registration could not be found.')


# Statistics & Settings Window


def statsWindowStart():
    statsWindow = Toplevel(root)
    statsWindow.geometry('600x400')
    statsWindow.config(bg=windowBG)
    statsWindow.resizable(FALSE, FALSE)
    root.withdraw()
    carWorkbook = openpyxl.load_workbook(carsFilePath)
    datasheet = carWorkbook.active
    lifetimeComp = 0
    monthlyComp = 0

    # for cell in datasheet['E']:
    #     if cell.value is not None:
    #         # date stored as yyyy-mm-dd
    #         date=cell.value.split('-')
    #         # if date[1]==datetime.date.today().month:

    # Stats for a Buyer
    for cell in datasheet['F']:
        if cell.value is not None:
            if cell.value == loggedInUser:
                # print(datasheet['K' + str(cell.row)].value)
                buyComp = datasheet['K' + str(cell.row)].value
                if buyComp is not None:
                    lifetimeComp = buyComp + lifetimeComp
                    if datasheet['E' + str(cell.row)].value.split('-')[1] == datetime.date.today().month:
                        monthlyComp = monthlyComp + buyComp

    # Stats for a Seller
    for cell in datasheet['I']:
        if cell.value is not None:
            if cell.value == loggedInUser:
                # print(datasheet['L' + str(cell.row)])
                sellComp = datasheet['L' + str(cell.row)].value
                if sellComp is not None:
                    lifetimeComp = lifetimeComp + sellComp
                    # print(lifetimeComp)
                    # totals the commission for the month
                    if datasheet['E' + str(cell.row)].value.split('-')[1] == datetime.date.today().month and datasheet['E' + str(cell.row)].value.split('-')[0] == datetime.date.today().year:
                        monthlyComp = monthlyComp + sellComp
    # Attempted Graph Code, Non-Functional in current state
    #
    # statsGraphFrame=ttk.Frame(statsWindow)
    # graphData1={'Month':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
    #       'Commission':[1515,1415,715,1625,2268,1532,980,585,1519,1628,1264,1000]}
    # graphdataframe1=pd.DataFrame(graphData1)
    # Graph = plt.Figure(figsize=(6, 5), dpi=100)
    # axes = Graph.add_subplot(111)
    # lineYear1 = FigureCanvasTkAgg(Graph, statsGraphFrame)
    # lineYear1.get_tk_widget().grid(row=0,column=0)
    # #graphdataframe1 = graphdataframe1[['Month', 'Commission']].groupby('Month').sum()
    # plt.xticks([0,1,2,3,4,5,6,7,8,9,10,11],graphData1['Month'])
    # graphdataframe1.plot(kind='line', legend=True, ax=axes)
    # axes.set_title('Commissions by Month')
    # statsGraphFrame.grid(row=0,column=0)

    statsLabelFrame = ttk.Frame(statsWindow)
    lifeCompLabel = ttk.Label(statsLabelFrame, text='Lifetime Compensation:')
    lifeCompValue = ttk.Label(statsLabelFrame, text=str(lifetimeComp))
    monthCompLabel = ttk.Label(statsLabelFrame, text='Month\'s Compensation:')
    monthCompValue = ttk.Label(statsLabelFrame, text=str(monthlyComp))
    lifeCompLabel.grid(row=0, column=0, padx=10, pady=10)
    lifeCompValue.grid(row=0, column=1, padx=10, pady=10)
    monthCompLabel.grid(row=1, column=0, padx=10, pady=10)
    monthCompValue.grid(row=1, column=1, padx=10, pady=10)
    statsLabelFrame.grid(row=1, column=0)

    statsWindow.protocol('WM_DELETE_WINDOW', lambda: subwindowclose(statsWindow))


# Main Window


root = Tk()
root.geometry('220x280')
root.resizable(FALSE, FALSE)
root.title('S.D.M.S')
root.withdraw()
root.config(bg=windowBG)
icon = PhotoImage(file='sdmsLogo.png')
root.iconphoto(True, icon)

rootFrame = ttk.Frame(root)

BuyerButton = ttk.Button(rootFrame, text='Log A Bought Car', padding=10, command=buyerwindowstart)
SellerButton = ttk.Button(rootFrame, text='Log A Sold Car', padding=10, command=sellerWindowStart)
StatsButton = ttk.Button(rootFrame, text='View Statistics', padding=10, command=statsWindowStart)
BuyerButton.grid(row=0, column=0, sticky='ew', pady=20)
SellerButton.grid(row=1, column=0, sticky='ew', pady=20)
StatsButton.grid(row=2, column=0, sticky='ew', pady=20)

rootFrame.grid(row=0, column=0, padx=40)

# Login Window


loginW = Toplevel(root)
loginW.geometry('520x380')
loginW.config(bg=windowBG)
loginW.resizable(FALSE, FALSE)
loginW.title('S.D.M.S Login')
loginFrame = ttk.Frame(loginW)
loginFrame.pack(padx=20, pady=20)

loginLabel = ttk.Label(loginFrame, text='S.D.M.S', font=('Calibri', 24))
loginUserEntry = ttk.Entry(loginFrame, font=('Calibri', 16))
loginUserLabel = ttk.Label(loginFrame, text='Username', font=('Calibri', 16))
loginPassEntry = ttk.Entry(loginFrame, show='*', font=('Calibri', 16))
loginPassLabel = ttk.Label(loginFrame, text='Password', font=('Calibri', 16))
loginButton = ttk.Button(loginFrame, text='Login',
                         command=lambda: login(username=loginUserEntry, password=loginPassEntry))

loginLabel.grid(row=0, column=0, columnspan=2, pady=30)
loginUserLabel.grid(row=1, column=0, padx=10)
loginUserEntry.grid(row=1, column=1, pady=5)
loginPassLabel.grid(row=2, column=0, padx=10)
loginPassEntry.grid(row=2, column=1, pady=5)
loginButton.grid(row=3, column=0, columnspan=2, pady=20)

loginW.protocol('WM_DELETE_WINDOW', lambda: root.destroy())
# binds the enter key to the login button for convenience
loginW.bind('<Return>', lambda event=None: loginButton.invoke())

sv_ttk.use_dark_theme()
root.mainloop()
