def num_rows(file):
    return len(file)

# return number of columns by counting number of columns in index row (Usually is header)
def num_columns(file):
    return len(file[0])