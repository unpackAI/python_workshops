import xlwings as xw


@xw.func
def hello(name):
    return f"Hello {name}!"


def get_or_create_sheet(name, wb):
    """Get a sheet and create one if needed (at the end)"""
    sheet_names = [ws.name for ws in wb.sheets]
    if name in sheet_names:
        return wb.sheets[name]
    return wb.sheets.add(name, after=wb.sheets[-1])


def main():
    # wb = xw.Book("rating_anime.xlsx")
    wb = xw.Book.caller()
    sheet = get_or_create_sheet("toto", wb)

    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye!"
    else:
        sheet["A1"].value = "Hello xlwings!"


# if __name__ == "__main__":
#     main()
