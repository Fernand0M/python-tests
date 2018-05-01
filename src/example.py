from contracts import ContractsBook

contracts_book = ContractsBook('/data/data.xlsx')
data = contracts_book.data

for contract in data:
    msg = f'{contract["name"]} {contract["dob"]} {contract["address"]}'
    print(msg)
