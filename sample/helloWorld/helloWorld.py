import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

@xw.func
def joke(x):
    wb = xw.Book.caller()
    jokes = ['one','two','three']
    for i, joke in jokes:
        if i == x: return joke


if __name__ == "__main__":
    xw.Book("helloWorld.xlsm").set_mock_caller()
    main()
