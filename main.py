from googletrans import Translator
import time
from tqdm import tqdm
import xlrd
import xlwt
import threading

def translate_text(translator, text, dest_language, translation_cache):
    if text is None or text == '':
        return ''

    translated_text = translation_cache.get(text, None)

    if translated_text is None:
        try:
            if isinstance(text, str):
                translation = translator.translate(text, dest=dest_language)
                translated_text = translation.text
                translation_cache[text] = translated_text
            else:
                translated_text = text
        except Exception as e:
            translated_text = text

    return translated_text

def translate_column(col, sheet_read, sheet_write, dest_language, translation_cache, thread_lock):
    translator = Translator()
    for row in range(0, sheet_read.nrows):
        text = sheet_read.cell_value(row, col)
        translated_text = translate_text(translator, text, dest_language, translation_cache)
        sheet_write.write(row, col, translated_text)

    with thread_lock:

        tqdm.write(f"Column {col} translated.")

def translate_excel(file_path,dest_language):
    start_time = time.time()

    wb_read = xlrd.open_workbook(file_path)
    sheet_read = wb_read.sheet_by_index(0)

    wb_write = xlwt.Workbook()
    sheet_write = wb_write.add_sheet("Sheet1")

    translation_cache = {}                      #Noticed multiple instances of same word in the .xls file, so keeping a cache to improve speed


    thread_lock = threading.Lock()              #lock to execute critical sections of the function
    threads = []
 
    for col in range(sheet_read.ncols):         #Creating thread objects for each column
        thread = threading.Thread(target=translate_column, args=(col, sheet_read, sheet_write, dest_language, translation_cache, thread_lock))
        threads.append(thread)
        thread.start()

    for thread in threads:                      #waiting for thread execution to finish
        thread.join()

    translated_file_path = file_path.replace('.xls', f'_translated_{dest_language}.xls')
    wb_write.save(translated_file_path)         #save .xls file

    end_time = time.time()
    elapsed_time = end_time - start_time

    print(f"\nTranslation completed in {elapsed_time:.2f} seconds.")
    print(f"Translated file saved as: {translated_file_path}")

if __name__ == "__main__":
    excel_file_path = 'Order Export.xls'

    translate_excel(excel_file_path,'en')