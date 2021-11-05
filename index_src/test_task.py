import sqlite3
import pandas as pd
import numpy as np
import openpyxl
from docxtpl import DocxTemplate
import docx


con = sqlite3.connect('test.db')




def run_time():
    sql_query = """SELECT * FROM testidprod
    where partner is NULL and state is NULL and bs=0 and (factor = 1 or factor = 2);"""
    dff = pd.read_sql(sql_query, con)
    dff = dff.drop(columns=['id','mir', 'raw', 'hash', 'meta', 'partner', 'state', 'bs', 'country'])
    dff = dff.reindex(columns=['factor', 'year', 'res'])
    year_add_dff = pd.DataFrame({'factor':[1, 1, 2, 2],
                          'year':[2006, 2020, 2006, 2020],
                         'res':[np.NaN, np.NaN, np.NaN, np.NaN]})


    df1 = dff.append(year_add_dff, ignore_index=True)
   # df1 = df1.sort_values(by='year')

    factor_add_df = df1[1: 16]
    factor_add_df.loc[:,'factor'] = 6
    factor_add_df.loc[:,'year'] = range(2006, 2021)

    #l = len(main_df)



    # factor_add_df = main_df.head(int(l/2))
    # factor_add_df['factor'] = [6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6]
    # # factor_add_df['year'] = range(2006, 2021)
    # # #

    #factor_add_df2 = factor_add_df.groupby(['world'])

    combined_df = [df1, factor_add_df]


    sums = pd.concat(combined_df, sort=False)
    sums = sums.groupby(['factor', 'year']).sum('res').rename(columns={'res': 'world'})

    #add_factor = sums.copy(deep=False)
    # sums1 = sums.append(add_factor, ignore_index=True)
    # factor_add_df = pd.DataFrame({'factor': [6, 6],
    #                             'year': [2009, 2010],
    #                             'res': [np.NaN, np.NaN]})
    # sums1 = sums.append(factor_add_df, ignore_index=True)



    #sums.loc[2007, 'world'] = 'ffffffffffffffffffffff'



    b = sums['world'].loc[2]
    c = sums['world'].loc[1]
    d = b.div(c)


    q = d.array
    sums.loc[6,'world'] = q

    # w = sums.loc[sums['factor'] == 6, ['world']]
    # cagr_sum = 0
    #
    # for idx, value  in enumerate(w):
    #     if not value:
    #         continue
    #     if idx+1 == len(w):
    #         continue
    #     if not w[idx+1]:
    #         continue
    #     current_cagr =  ((w[idx+1] / value) ** (1 / 100) - 1) * 100
    #     cagr_sum += current_cagr

    df = sums
    sums = sums.T


    sums.to_excel('report.xlsx')




    # sums[
    #     'CAGR1'] = ((sums['2020'] / sums['2006']) ** (1 / 100) - 1) * 100
    # print(sums)

    # pandas data frame


    document = docx.Document()
    styles = document.styles

   # document.add_table(3, 3, style=None)





    #docx.styles.style._TableStyle[]
    #document.styles = 'Table'



    document.save('report.docx')

    doc = DocxTemplate('report.docx')
    a = 5
   # context = {'cagr': df, 's_year': 2006, 'e_year': 2020, 'direction': 'grew' }
    #context =  df
    context =  {'s_year': 2006, 'e_year': 2020, 'direction': 'grew' }


    doc.render(context)
    doc.save('report.docx')



    print(doc)


    print(df1)
    print(sums)



if __name__ == '__main__':
    run_time()




