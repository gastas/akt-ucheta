import pandas as pd
import streamlit as st
import io
from docxtpl import DocxTemplate
import datetime



def to_docx(df, date, smena, number):
    context = {}
    context['number'] = number
    context['tbl_contents'] = df
    context['smena'] = smena
    context['day'] = date

    doc = DocxTemplate("akt_tpl.docx")
    # подставляем контекст в шаблон
    doc.render(context)
    # сохраняем и смотрим, что получилось 
    #doc.save("generated_docx.docx")
    file_stream = io.BytesIO()
    # Save the .docx to the buffer
    doc.save(file_stream)
    # Reset the buffer's file-pointer to the beginning of the file
    file_stream.seek(0)

    return file_stream


st.set_page_config(
    layout="centered", page_icon="🖱️", page_title="Акт учета времени оказания услуг сотрудниками"
)
st.title("Акт учета времени оказания услуг сотрудниками")

text = st.text_area("ФИО, одно в строку.")
lines = text.split("\n")
lines = sorted(lines)

date = st.date_input('Дата')

date_short = date.strftime("%-d-%m")

date = date.strftime("%-d.%m.%Y")


smena = st.selectbox(
    'Смена',
    ('08:00-20:00', '20:00-08:00', '16:00-04:00', '09:30-21:30', '21:30-09:30' ))

if smena == '08:00-20:00':
    start = '08:00'
    end = '20:00'
elif smena == '16:00-04:00':
    start = '16:00'
    end = '04:00'
elif smena == '21:30-09:30':
    start = '21:30'
    end = '09:30'
elif smena == '09:30-21:30':
    start = '09:30'
    end = '21:30'
else: 
    start = '20:00'
    end = '08:00'



df = pd.DataFrame(lines)

#st.table(df)

df['row_number'] = df.reset_index().index
df['row_number'] += 1

df["Вид услуг"] = "Комплектовщик"
df["Начало"] = start
df["Окончание"] = end
df["Перерыв"] = "1,5"
df["Количество часов"] = "10,5"
df["Подпись"] = " "

number = df['row_number'].iat[-1]

df = st.experimental_data_editor(df)



df = df.to_numpy()




#st.write(number)




    

doc_download = to_docx(df, date, smena, number)
docxname = "vash-perevozchik-{}.docx".format(date_short)

if doc_download:
    st.download_button(
        label="Скачать",
        data=doc_download,
        file_name=docxname,
        mime="docx"
    )



#st.table(df)

