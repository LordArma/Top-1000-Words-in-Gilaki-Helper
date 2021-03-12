import sqlite3
from shutil import copyfile
import pandas as pd
import os
import json
from json2xml import json2xml
from json2xml.utils import readfromjson
import asyncio
from PIL import Image
from pyppeteer import launch
from datetime import datetime
import glob
import shutil
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_DIRECTION, WD_TABLE_ALIGNMENT
import pyexcel as pe


db_dir = 'words.db'
release_dir = "./Top-1000-Words-in-Gilaki"
start_range = 1
end_range = 1001


def init():
    if not os.path.exists(release_dir):
        os.makedirs(release_dir)

    if not os.path.exists(release_dir + '/CSV'):
        os.makedirs(release_dir + '/CSV')

    if not os.path.exists(release_dir + '/docs'):
        os.makedirs(release_dir + '/docs')

    if not os.path.exists(release_dir + '/Excel'):
        os.makedirs(release_dir + '/Excel')

    if not os.path.exists(release_dir + '/Flash Card'):
        os.makedirs(release_dir + '/Flash Card')

    if not os.path.exists(release_dir + '/JSON'):
        os.makedirs(release_dir + '/JSON')

    if not os.path.exists(release_dir + '/PDF'):
        os.makedirs(release_dir + '/PDF')

    if not os.path.exists(release_dir + '/SQLite'):
        os.makedirs(release_dir + '/SQLite')

    if not os.path.exists(release_dir + '/Word'):
        os.makedirs(release_dir + '/Word')

    if not os.path.exists(release_dir + '/XML'):
        os.makedirs(release_dir + '/XML')

    copyfile("./templates/favicon.png", release_dir + "/docs/favicon.png")
    copyfile("./templates/index.html",  release_dir + "/docs/index.html")
    copyfile("./templates/LICENSE",  release_dir + "/LICENSE")


def normalize(path : str = release_dir + "/SQLite/Top 1000 Words in Gilaki.sqlite"):
    try:
        conn = sqlite3.connect(db_dir)
        cursor = conn.cursor()
        print("Connected to SQLite successfully .")
        print("Normalization started...")
        sqlite_update_query = "UPDATE tbl_words SET glk_word = trim(glk_word), glk_example = trim(glk_example), " \
                              "fa_word = trim(fa_word), fa_example = trim(fa_example) "
        count = cursor.execute(sqlite_update_query)
        conn.commit()
        print("Records striped successfully.", cursor.rowcount)
        cursor.close()

        print("Releasing SQLite...")
        copyfile(db_dir, release_dir + "/SQLite/Top 1000 Words in Gilaki.sqlite")
        print("SQLite Released Successfully.")

    except sqlite3.Error as error:
        print("Failed to stripe rows.", error)
    finally:
        if (conn):
            conn.close()
            print("The SQLite connection is closed.")

    print("Sorting database started...")
    copyfile("./templates/database.db", path)
    try:
        conn = sqlite3.connect(db_dir)
        cursor = conn.cursor()

        sqlite_select_query = "SELECT * from tbl_words order by glk_word"
        cursor.execute(sqlite_select_query)
        records = cursor.fetchall()

        id = 0
        conn2 = sqlite3.connect(path)
        for row in records:
            id += 1
            cursor2 = conn2.cursor()
            r1 = row[1]
            r2 = row[2]
            r3 = row[3]
            r4 = row[4]
            r1 = r1.replace("'", "ٰ")
            r2 = r2.replace("'", "ٰ")
            r3 = r3.replace("'", "ٰ")
            r4 = r4.replace("'", "ٰ")
            sqlite_insert_query = f"""INSERT INTO tbl_words
                                      (id, glk_word, glk_example, fa_word, fa_example) 
                                       VALUES 
                                      ({id}, "{r1}", "{r2}", "{r3}", "{r4}")"""

            count = cursor2.execute(sqlite_insert_query)
            conn2.commit()
        cursor.close()

    except sqlite3.Error as error:
        print("Failed to read data from sqlite table", error)
    finally:
        if (conn):
            conn.close()
            print("Database sorted Successfully.")


def make_csv():
    print("Making CSV started...")
    tmp_output_dir = release_dir + "/CSV/Top 1000 Words in Gilaki.csv"
    conn = sqlite3.connect(db_dir, isolation_level=None, detect_types=sqlite3.PARSE_COLNAMES)
    db_df = pd.read_sql_query("SELECT * FROM tbl_words", conn)
    db_df.to_csv(tmp_output_dir, index=False)
    print("CSV made Successfully.")


def make_json():
    print("Making JSON started...")
    tmp_min_output_dir = release_dir + "/JSON/Top 1000 Words in Gilaki.min.json"
    conn = sqlite3.connect(db_dir)
    conn.row_factory = sqlite3.Row
    db = conn.cursor()
    rows = db.execute("SELECT * from tbl_words").fetchall()
    conn.commit()
    conn.close()
    js = json.dumps([dict(ix) for ix in rows], ensure_ascii=False).encode('utf8')
    f = open(tmp_min_output_dir, "w")
    f.write(js.decode())
    f.close()
    print("Minified JSON made Successfully.")
    tmp_output_dir = release_dir + "/JSON/Top 1000 Words in Gilaki.json"
    js = json.dumps([dict(ix) for ix in rows], indent=4, ensure_ascii=False).encode('utf8')
    f = open(tmp_output_dir, "w")
    f.write(js.decode())
    f.close()
    print("Beautified JSON made Successfully.")


def update_docs():
    print("Updating docs...")
    copyfile(release_dir + "/JSON/Top 1000 Words in Gilaki.min.json", release_dir + "/docs/Top 1000 Words in Gilaki.min.json")
    print("docs updated Successfully.")


def make_xml():
    print("Making XML started...")
    tmp_output_dir = "database.xml"
    data = readfromjson(release_dir + "/JSON/Top 1000 Words in Gilaki.min.json")
    data = json2xml.Json2xml(data, wrapper="wordlist", pretty=True, attr_type=False).to_xml()
    data = data.replace("<item>", "<word>")
    data = data.replace("</item>", "</word>")
    data = data.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8"?>')
    f = open(tmp_output_dir, "w")
    f.write(data)
    f.close()
    print("XML made Successfully.")
    print("Releasing XML...")
    copyfile(tmp_output_dir, release_dir + "/XML/Top 1000 Words in Gilaki.xml")
    print("XML Released Successfully.")
    os.remove(tmp_output_dir)


def readtemplate(path : str) -> str:
    f = open(f"{path}", "r")
    return f.read()


def readdb(path : str = release_dir + "/SQLite/Top 1000 Words in Gilaki.sqlite"):
    print("Making HTML flashcards started...")

    word = readtemplate("./templates/word.html")
    meaning = readtemplate("./templates/meaning.html")
    full = readtemplate("./templates/full.html")

    try:
        conn = sqlite3.connect(path)
        cursor = conn.cursor()

        sqlite_select_query = "SELECT * from tbl_words"
        cursor.execute(sqlite_select_query)
        records = cursor.fetchall()

        for row in records:
            flash = word
            flash = flash.replace("{glk_word}", row[1])
            createflash(f"{row[0]}-word.html", flash)

        for row in records:
            flash = meaning
            flash = flash.replace("{glk_word}", row[1])
            flash = flash.replace("{glk_example}", row[2])
            createflash(f"{row[0]}-meaning.html", flash)

        for row in records:
            txt = ""
            flash = full
            flash = flash.replace("{glk_word}", row[1])
            txt += row[1] + "\n"
            flash = flash.replace("{fa_word}", row[3])
            txt += row[3] + "\n"
            flash = flash.replace("{glk_example}", row[2])
            txt += row[2] + "\n"
            flash = flash.replace("{fa_example}", row[4])
            txt += row[4] + "\n"
            createflash(f"{row[0]}-full.html", flash)
            createinfo(f"{row[0]}.txt", txt)

        cursor.close()

    except sqlite3.Error as error:
        print("Failed to read data from sqlite table", error)
    finally:
        if (conn):
            conn.close()
            print("HTML flashcards made Successfully.")


def crop_center(pil_img, crop_width, crop_height):
    img_width, img_height = pil_img.size
    return pil_img.crop(((img_width - crop_width) // 2,
                         (img_height - crop_height) // 2,
                         (img_width + crop_width) // 2,
                         (img_height + crop_height) // 2))


def makejpg(typ):
    for i in range(start_range, end_range, 1):
        async def main():
            browser = await launch(headless=True, options={'args': ['--no-sandbox']})
            page = await browser.newPage()
            await page.setViewport({'width': 600, 'height': 600})
            await page.goto(f"file:///home/arma/PycharmProjects/top-1000-words-in-gilaki-helper/TEMP/{i}-{typ}.html")
            await page.screenshot({'path': f"TEMP/{i}-{typ}.jpg"})
            await browser.close()

        asyncio.get_event_loop().run_until_complete(main())

        thumb_width = 600

        im = Image.open(f"TEMP/{i}-{typ}.jpg")
        im_thumb = crop_center(im, thumb_width, thumb_width)
        im_thumb.save(f"TEMP/{i}-{typ}.jpg", quality=100)


def createflash(name, html):
    f = open(f"TEMP/{name}", "w")
    f.write(html)
    f.close()

def createinfo(name, text):
    f = open(f"TEMP/{name}", "w")
    f.write(text)
    f.close()


def make_flash_html():
    readdb()


def make_flash_jpg():
    print("Making JPG flashcards started...")
    print("Making JPG word flashcards started...")
    makejpg("word")
    print("Making JPG meaning flashcards started...")
    makejpg("meaning")
    print("Making JPG full flashcards started...")
    makejpg("full")
    print("JPG flashcards made Successfully.")
    print("Releasing JPG flashcards...")
    files = glob.iglob(os.path.join("./TEMP/", "*.jpg"))
    for file in files:
        if os.path.isfile(file):
            shutil.copy2(file, release_dir + "/Flash Card/")
    print("JPG flashcards Released Successfully.")


def make_gif():
    print("Making GIF flashcards started...")

    def makegif(num):
        os.system(f'convert -delay 200 -loop 0 ./TEMP/{num}-word.jpg ./TEMP/{num}-meaning.jpg ./TEMP/{num}-full.jpg ./TEMP/{num}-animation.gif')

    for i in range(start_range, end_range, 1):
        makegif(i)

    print("GIF flashcards made Successfully.")
    print("Releasing GIF flashcards...")
    files = glob.iglob(os.path.join("./TEMP/", "*.gif"))
    for file in files:
        if os.path.isfile(file):
            shutil.copy2(file, release_dir + "/Flash Card/")
            os.remove(file)
    print("GIF flashcards Released Successfully.")
    files = glob.iglob(os.path.join("./TEMP/", "*.jpg"))
    for file in files:
        if os.path.isfile(file):
            os.remove(file)


def make_docx(path : str = release_dir + "/SQLite/Top 1000 Words in Gilaki.sqlite"):
    print("Making DOCX started...")

    document = Document("./templates/database.docx")
    table = document.tables[0]
    table.directions = WD_TABLE_DIRECTION.RTL
    table.allow_autofit = True

    row = table.rows[0].cells

    row[0].text = ''
    row[0].paragraphs[0].add_run('Farsi Example').bold = True
    row[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    row[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    row[1].text = ''
    row[1].paragraphs[0].add_run('Farsi Word').bold = True
    row[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    row[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    row[2].text = ''
    row[2].paragraphs[0].add_run('Gilaki Example').bold = True
    row[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    row[2].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    row[3].text = ''
    row[3].paragraphs[0].add_run('Gilaki Word').bold = True
    row[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    row[3].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    try:
        conn = sqlite3.connect(path)
        cursor = conn.cursor()

        sqlite_select_query = "SELECT * from tbl_words"
        cursor.execute(sqlite_select_query)
        records = cursor.fetchall()

        for r in records:
            row = table.add_row().cells

            row[0].text = ''
            row[0].paragraphs[0].add_run(r[4])
            row[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row[0].alignment = WD_TABLE_ALIGNMENT.RIGHT

            row[1].text = ''
            row[1].paragraphs[0].add_run(r[3])
            row[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row[1].alignment = WD_TABLE_ALIGNMENT.RIGHT

            row[2].text = ''
            row[2].paragraphs[0].add_run(r[2])
            row[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row[2].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row[2].alignment = WD_TABLE_ALIGNMENT.RIGHT

            row[3].text = ''
            row[3].paragraphs[0].add_run(r[1])
            row[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            row[3].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            row[3].alignment = WD_TABLE_ALIGNMENT.RIGHT

        cursor.close()

    except sqlite3.Error as error:
        print("Failed to read data from sqlite table", error)
    finally:
        if (conn):
            conn.close()
            print("DOCX made Successfully.")

    document.save(release_dir + "/Word/Top 1000 Words in Gilaki.docx")


def make_pdf():
    print("Making PDF started...")
    tmp_dir = release_dir + "/Word/Top 1000 Words in Gilaki.docx"
    os.system(f'libreoffice --headless --convert-to pdf --outdir "{release_dir}/PDF/" "{tmp_dir}"')
    print("PDF made Successfully.")


def make_xlsx(path : str = release_dir + "/SQLite/Top 1000 Words in Gilaki.sqlite"):
    print("Making Excel started...")
    sheet = pe.get_sheet(file_name="./templates/database.xlsx")

    try:
        conn = sqlite3.connect(path)
        cursor = conn.cursor()

        sqlite_select_query = "SELECT * from tbl_words"
        cursor.execute(sqlite_select_query)
        records = cursor.fetchall()

        for r in records:
            sheet.row += [r[1], r[2], r[3], r[4]]

        cursor.close()

        sheet.rightToLeft = True
        sheet.save_as(release_dir + "/Excel/Top 1000 Words in Gilaki.xlsx")

    except sqlite3.Error as error:
        print("Failed to read data from sqlite table", error)
    finally:
        if (conn):
            conn.close()
            print("XLSX made Successfully.")


def change_readme():
    print("Updating README started...")
    tmp_dir = "./templates/README.txt"
    readme = readtemplate(tmp_dir)
    d = datetime.now()
    current_time = d.strftime('%d, %b %Y')
    readme = readme.replace("{date}", f"{current_time}")
    f = open(release_dir + "/README.md", "w")
    f.write(readme)
    f.close()
    print("README updated Successfully.")


def push_release():
    os.system(f'cd {release_dir} && git add . && git commit -m "fix typo" && git push')


if __name__ == '__main__':
    init()
    normalize()
    make_csv()
    make_json()
    update_docs()
    make_xml()
    make_flash_html()
    make_flash_jpg()
    make_gif()
    make_docx()
    make_pdf()
    make_xlsx()
    change_readme()
    push_release()

