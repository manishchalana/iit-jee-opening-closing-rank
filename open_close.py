import pandas as pd
import PyPDF2
pdfFileObj = open('openingclosingranks2019.pdf', 'rb')
import sys

pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
num_cols = 8

category ='OBC-NCL'
seat_pool = 'Gender-Neutral'

final_df = None

for i in range(pdfReader.numPages):
    pageObj = pdfReader.getPage(i)
    content = pageObj.extractText()
    orig_content = pageObj.extractText()
    content = content.replace(" \n", " ")
    content = content.replace("Round \nNo", "Round No")
    content = content.replace("Openning \nRank", "Openning Rank")
    content = content.replace("Closing \nRank", "Closing Rank")
    content = content.replace('Allotted \nQuota', 'Allotted Quota')
    content = content.replace("Bachelor of \nTechnology", "Bachelor of Technology")
    content = content.replace("Arunachal Pradesh ", "Arunachal Pradesh\n")
    content = content.replace("(IIIT) Dharwad ", "(IIIT) Dharwad\n")

    content = content.split("\n")

    num_rows = len(content)//num_cols
    content = [content[j*num_cols:(j+1)*num_cols] for j in range(num_rows)]


    assert sum([line[0]=='7' for line in content])==len(content)-1

    df = pd.DataFrame(content[1:], columns=content[0])


    df['Page Number'] = i


    df = df[(df['Category']==category) & (df['Seat Pool']==seat_pool)]

    if final_df is None:
        final_df = df.copy()
    else:
        final_df = pd.concat([final_df, df])


pdfFileObj.close()

# final_df['Preparatory list'] = final_df.apply(lambda x: str(x['Closing Rank'])[-1]=="P", axis=1)

# final_df['New Rank'] = final_df.apply(lambda x: print(x.index), axis=1)
final_df['Closing Rank'] = final_df.apply(lambda x: str(x['Closing Rank'])[:-1] if str(x['Closing Rank'])[-1]=="P" else str(x['Closing Rank']), axis=1)

final_df['Closing Rank'] = pd.to_numeric(final_df['Closing Rank'])
final_df['Openning Rank'] = pd.to_numeric(final_df['Openning Rank'])
final_df['Round No'] = pd.to_numeric(final_df['Round No'])

final_df = final_df.sort_values('Closing Rank')
final_df.to_excel(category + "_" + seat_pool + '.xlsx', index=False)
