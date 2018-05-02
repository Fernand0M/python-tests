from __future__ import print_function
from mailmerge import MailMerge
from contracts import ContractsBook

template = "test.docx"
contracts_book = ContractsBook('/data/data.xlsx')
data = contracts_book.data


for contract in data:
    msg = f'{contract["name"]} {contract["dob"]} {contract["address"]}'
    document = MailMerge(template)
    document.merge(
        name=contract["name"],
        IS=contract["dob"],
        date=contract["address"]
    )
    document.write('testoutput' + contract["name"] + '.docx')

    print(msg)





