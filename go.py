import xlrd
import xlwt

def run_quickstart(text):
    # [START translate_quickstart]
    # Imports the Google Cloud client library
    from google.cloud import translate

    # Instantiates a client
    translate_client = translate.Client()

    # The text to translate
    # text = 'Broshtanドレッサーは同期​​してそうだ印象的なコントラストです。そのきれいな、箱型のフレーム、フラッシュマウント引き出しやラウンドテーパー脚がTにミッドセンチュリーモダンデザインの本質をキャプチャし、ドレッサーのプロフィールは非常に現代的かもしれないが、仕上がりは悲惨と独特のソーマークで、非常に素朴ですドラマティックな音色のバリエーションを持つ木目。品質ダブテール構造は、内側からの美しさを反映しています。'
    # The target language
    # target = 'en'
    
    translation = translate_client.translate(
        text,
        target_language="ja")

    # print(u'Translation: {}'.format(translation['translatedText']))
    # [END translate_quickstart]

    print(translation['translatedText'])
    return translation['translatedText']


x = run_quickstart("fuck you")
print(x)
workbook = xlrd.open_workbook("./cnm.xls")
wb2 = xlwt.Workbook()
booksheet = workbook.sheet_by_index(0)
writer = wb2.add_sheet('sheet1' , cell_overwrite_ok = True)    
try:
    for i in range(1,700):
        m = booksheet.row_values(i)
        text = m[5]
        x = run_quickstart(text)
        print(x)
        writer.write(i,0,x)
except:
    wb2.save('rinimabi.xls')

wb2.save('rinimabi.xls')
