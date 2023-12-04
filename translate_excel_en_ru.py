import pandas as pd
import sqlite3
from translate import Translator
from tqdm import tqdm
import os
global tr_count, reuse_count, progress_count


def get_counter(name):
    if name == "tr_count":
        return tr_count
    return reuse_count


def drop_db():
    con = sqlite3.connect('translation_en-ru.db')
    cur = con.cursor()
    cur.execute('DROP TABLE IF EXISTS translation')
    con.commit()
    con.close()


def to_translate(file, new_file):
    df = pd.read_excel(file, header=None)
    con = sqlite3.connect('translation_en-ru.db')
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS translation (
            en  text,
            ru  text);
            """)
    con.commit()

    # translator_zh_en = Translator(from_lang="zh", to_lang="en")
    translator_en_ru = Translator(from_lang="en", to_lang="ru")
    global tr_count, reuse_count, progress_count
    tr_count, reuse_count, progress_count = 0, 0, 0
    df_ru = df.copy()

    for row in tqdm(range(df.shape[0])):
        for col in range(df.shape[1]):
            # print(row, col)
            progress_count += 100.0 / (df.shape[0] * df.shape[1])
            # print(progress_count)
            if type(df.iloc[row, col]) is str:
                cell = df.iloc[row, col].replace('"', '``')
                c = cur.execute(f'SELECT count(*) FROM translation WHERE en = "{cell}"')
                if c.fetchone()[0] < 1:  # or ex == 1:
                    tr_count += 1
                    # print('ask to translation')
                    # cell_en = translator_zh_en.translate(cell)
                    cell_ru = translator_en_ru.translate(cell)
                    if cell_ru.find('MYMEMORY WARNING') >= 0:
                        return print(f'{cell_ru}')
                    cur.execute("INSERT INTO translation (en, ru) VALUES (?,?)",
                                (cell, cell_ru))
                    # df_en.iloc[row, col] = cell_en
                    df_ru.iloc[row, col] = cell_ru
                    con.commit()
                else:
                    reuse_count += 2
                    # print('Reuse cell:', cell)
                    s = cur.execute(f'SELECT en, ru  FROM translation WHERE en = "{cell}"').fetchone()
                    df.iloc[row, col] = s[0]
                    df_ru.iloc[row, col] = s[1]
    con.close()

    with pd.ExcelWriter(new_file) as writer:
        df.to_excel(writer, sheet_name='en', header=None, index=False)
        # df_en.to_excel(writer, sheet_name='en', header=None, index=False)
        df_ru.to_excel(writer, sheet_name='ru', header=None, index=False)

    # print("Saved!!!")
    # e_new.insert(0, str(tr_count))
    # e_reuse.insert(0, str(reuse_count))


if __name__ == '__main__':
    # file = os.path.abspath('UM20220817041 zh.xlsx')
    file = os.path.abspath(r"c:\Шелопин\Gettering\2023 10 24 Для Договора\3 Gettering Diffusion\DOA-400L Horizontal DOA Equipment FUM (LP Diffusion)20230906.xlsx")
    new_file = os.path.abspath(r"c:\Шелопин\Gettering\2023 10 24 Для Договора\3 Gettering Diffusion\DOA-400L Horizontal DOA Equipment FUM (LP Diffusion)20230906_en_ru.xlsx")
    to_translate(file, new_file)
    print('New translations :', tr_count)
    print("Translation reuse:", reuse_count)
