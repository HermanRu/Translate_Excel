import pandas as pd
import sqlite3
from translate import Translator
from tqdm import tqdm

global tr_count, reuse_count, progress_count


def get_counter(name):
    if name == "tr_count":
        return tr_count
    return reuse_count


def drop_db():
    con = sqlite3.connect('translation_zh_en.db')
    cur = con.cursor()
    cur.execute('DROP TABLE IF EXISTS translation')
    con.commit()
    con.close()


def to_translate(file, new_file):
    df = pd.read_excel(file, header=None)
    con = sqlite3.connect('translation_zh_en.db')
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS translation (
            zh  text,
            en  text,
            ru  text);
            """)
    con.commit()

    translator_zh_en = Translator(from_lang="zh", to_lang="en")
    translator_en_ru = Translator(from_lang="en", to_lang="ru")
    global tr_count, reuse_count, progress_count
    tr_count, reuse_count, progress_count = 0, 0, 0
    df_en, df_ru = df.copy(), df.copy()

    for row in tqdm(range(df.shape[0])):
        for col in range(df.shape[1]):
            progress_count += 100.0 / (df.shape[0] * df.shape[1])
            # print(progress_count)
            if type(df.iloc[row, col]) is str:
                cell = df.iloc[row, col].replace('"', '`')
                c = cur.execute(f'SELECT count(*) FROM translation WHERE zh = "{cell}"')
                if c.fetchone()[0] < 1:  # or ex == 1:
                    tr_count += 2
                    cell_en = translator_zh_en.translate(cell)
                    cell_ru = translator_en_ru.translate(cell_en)
                    if cell_ru.find('MYMEMORY WARNING') >= 0 or cell_en.find('MYMEMORY WARNING') >= 0:
                        return print(f'{cell_en}\n{cell_ru}')
                    cur.execute("INSERT INTO translation (zh, en, ru) VALUES (?,?,?)",
                                (cell, cell_en, cell_ru))
                    df_en.iloc[row, col] = cell_en
                    df_ru.iloc[row, col] = cell_ru
                    con.commit()
                else:
                    reuse_count += 2
                    # print('Reuse cell:', cell)
                    s = cur.execute(f'SELECT en, ru  FROM translation WHERE zh = "{cell}"').fetchone()
                    df_en.iloc[row, col] = s[0]
                    df_ru.iloc[row, col] = s[1]
    con.close()

    with pd.ExcelWriter(new_file) as writer:
        df.to_excel(writer, sheet_name='zh', header=None, index=False)
        df_en.to_excel(writer, sheet_name='en', header=None, index=False)
        df_ru.to_excel(writer, sheet_name='ru', header=None, index=False)

    # print("Saved!!!")
    # e_new.insert(0, str(tr_count))
    # e_reuse.insert(0, str(reuse_count))


if __name__ == '__main__':
    file = 'book_zh.xlsx'
    new_file = 'book_zh_en_ru.xlsx'
    to_translate(file, new_file)
    print('New translations :', tr_count)
    print("Translation reuse:", reuse_count)
