import streamlit as st
import io
import pandas as pd
import re
from docxtpl import DocxTemplate
from docx import Document
from datetime import date

from dadata import Dadata
import requests
import pymorphy3
from pytrovich.enums import NamePart, Gender, Case
from pytrovich.maker import PetrovichDeclinationMaker
from pytrovich.detector import PetrovichGenderDetector


st.set_page_config(
   page_title="Заполнение Word",
   # page_icon="⚖️",
   layout="wide", #centered, wide
   # initial_sidebar_state="expanded",
)

token_dadata = "567ebcc7e05cfd43af075b5d76885366245ee52c"
token_checko = 'au6EiqY5YJcXl5re'
morph = pymorphy3.MorphAnalyzer(lang="ru")
maker = PetrovichDeclinationMaker()
detector = PetrovichGenderDetector()

def initials_and_greeting(fio):
    if len(fio.split()) == 3:
        f, i, o = fio.split()
        nameshort_ceo = f + " " + i[0] + "." + o[0] + "."
        gender = detector.detect(firstname=i, middlename=o)
        if gender == Gender.MALE:
            greeting = "Уважаемый " + i + " " + o + "!"
        elif gender == Gender.FEMALE:
            greeting = "Уважаемая " + i + " " + o + "!"
        else:
            greeting = "Уважаемые господа!"
    else:
        nameshort_ceo = fio
        greeting = "Уважаемые господа!"
    return nameshort_ceo, greeting


# Склонение ФИО - падежи Case.GENITIVE, Case.DATIVE
def sklonenie_fio(fio, case=Case.GENITIVE):

    if len(fio.split()) == 3:

        f, i, o = fio.split()
        gender = detector.detect(firstname=i, middlename=o)
        d_f = maker.make(NamePart.LASTNAME, gender, case, f)
        d_i = maker.make(NamePart.FIRSTNAME, gender, case, i)
        d_o = maker.make(NamePart.MIDDLENAME, gender, case, o)

        # Возвращение склоненного ФИО в формате "Иванов Иван Иванович"
        return f"{d_f} {d_i} {d_o}", f"{d_f} {d_i[0]}.{d_o[0]}."
    else:
        return fio


#Функция, разбивающая текст в список с учетом пунктуации
def split_text(text):
    # Define a regular expression pattern to match words, spaces, punctuation, and double quotation marks
    pattern = r'\b\w+\b|[.,!?;"()-]|\s'

    # Use re.findall to tokenize the string based on the pattern
    tokens = re.findall(pattern, text)

    return tokens


def capitalize_symbols(reference_list, target_list):
    result_list = []

    for ref_text, target_text in zip(reference_list, target_list):
        if len(ref_text) > 2 and ref_text[0].isupper() and ref_text[1].islower():
            # Capitalize only the first symbol in the second list
            result_text = target_text[0].upper() + target_text[1:].lower()
        elif ref_text[:2].isupper():
            # Capitalize the entire second list
            result_text = target_text.upper()
        else:
            result_text = target_text

        result_list.append(result_text)

    return result_list


# Функция проверяет, что слово в списке стоит в кавычках, чтобы не склонять его
def has_quotes_around(idx, lst):
    found_before = False
    found_after = False
    # Check for '"' before the index
    for i in range(idx):
        if lst[i] == '"':
            found_before = True
            break

    # Check for '"' after the index
    for i in range(idx + 1, len(lst)):
        if lst[i] == '"':
            found_after = True
            break
    # Return False if both '"' before and after the index are found
    return not (found_before and found_after)


# Функция для склонения должности руководителя. Примеры написания падежей - nomn, datv, gent
def sklonenie(input_string, case="gent"):
    if input_string:
        result_list = []
        # Декомпозиция текста на список слов и знаков препинания
        list_cleaned = split_text(input_string)
        # Изменение падежей в словах именительного падежа
        for i, word in enumerate(list_cleaned):
            try:
                if morph.parse(word)[0].tag.case == "nomn" and has_quotes_around(i, list_cleaned):  # Проверяет, что элемент списка в именительном падеже и не стоит в кавычках
                    result_list.append(morph.parse(word)[0].inflect({case}).word)
                else:
                    result_list.append(word)
            except:
                result_list.append(word)
        result_list = capitalize_symbols(list_cleaned, result_list)
        result = "".join(result_list)

        return result
    else:
        return "Не получилось завершить склонение"


def get_founders(r_checko):
    to_founders = "Укажите_адресат_Собственники"
    founders = []

    founders_count = len(r_checko["data"]["Учред"]["ФЛ"]) + len(r_checko["data"]["Учред"]["РосОрг"]) + len(r_checko["data"]["Учред"]["ИнОрг"]) + len(r_checko["data"]["Учред"]["ПИФ"]) + len(r_checko["data"]["Учред"]["РФ"])

    if r_checko["data"]["ОКОПФ"]["Код"] == "12300" and founders_count > 1:
        to_founders = "Участникам"
    elif r_checko["data"]["ОКОПФ"]["Код"] == "12300" and founders_count == 1:
        to_founders = "Единственному участнику"
    elif r_checko["data"]["ОКОПФ"]["Код"] == "12247" or r_checko["data"]["ОКОПФ"]["Код"] == "12267":
        to_founders = "Акционерам/Единственному акционеру"

    try:
        for i in r_checko["data"]["Учред"]["ФЛ"]:
            founders.append(i["ФИО"] + ", " + str(i["Доля"]["Процент"]))
    except Exception as e:
        st.write(e)

    try:
        for i in r_checko["data"]["Учред"]["РосОрг"]:
            founders.append(i["НаимСокр"] + f"({i['ИНН']})" + ", " + str(i["Доля"]["Процент"]))
    except Exception as e:
        st.write(e)

    try:
        for i in r_checko["data"]["Учред"]["ПИФ"]:
            founders.append(i["Наим"] + f" под управлением {i['УпрКом']['НаимСокр']} ({i['УпрКом']['ИНН']})" + ", " + str(i["Доля"]["Процент"]))
            # founders.append(i["Наим"])
    except Exception as e:
        st.write(e)

    try:
        for i in r_checko["data"]["Учред"]["ИнОрг"]:
            founders.append(i["НаимПолн"] + f" ({i['Страна']}, {i['РегНомер']})" + ", " + str(i["Доля"]["Процент"]))
    except Exception as e:
        st.write(e)

    try:
        for i in r_checko["data"]["Учред"]["РФ"]:
            founders.append(i["Тип"] + ", " + str(i["Доля"]["Процент"]))
    except Exception as e:
        st.write(e)

    founders_string = " / ".join(founders)
    # founders_string = "test"

    return to_founders, founders_string


def fill_empty_dataframe(empty_df, source_df):
    # Ensure all columns in empty_df are filled with data from source_df or remain blank
    for column in empty_df.columns:
        if column in source_df.columns:
            empty_df[column] = source_df[column]
        else:
            empty_df[column] = "XXX"  # Leave blank if no matching data in source_df
    return empty_df


@st.cache_data
def get_data(INN):
    # Получение информации с Dadata об организации
    with Dadata(token_dadata) as dadata:
        r_data = dadata.find_by_id(name="party", query=INN)

    # Получение информации с Checko об организации
    r_checko = requests.get(f'https://api.checko.ru/v2/company?key={token_checko}&inn={INN}').json()
    # st.write(r_checko) # Удалить

    # Определение имени и должности руководителя
    FIO_CEO = "ФИО_Руководителя"
    CEO_post = "Руководитель"
    try:
        FIO_CEO = r_data[0]['data']['management']['name']
        CEO_post = (r_data[0]['data']['management']['post']
                    .replace("ГЕНЕРАЛЬНЫЙ ДИРЕКТОР", "Генеральный директор")
                    .replace("ДИРЕКТОР", "Директор")
                    .replace("ПРЕДСЕДАТЕЛЬ ПРАВЛЕНИЯ", "Председатель правления"))
    except:
        st.caption(
            "Не удалось найти ФИО руководителя, возможно компания управляется юр. лицом или информация о руководителе скрыта в ЕГРЮЛ")
        try:
            management_company = r_checko['data']['УпрОрг']["НаимСокр"]
            management_company_INN = r_checko['data']['УпрОрг']["ИНН"]
            st.caption(
                f"Компания управляется юридическим лицом - {management_company} (ИНН - {management_company_INN})")
            r_data_management_company = Dadata(token_dadata).find_by_id(name="party", query=management_company_INN)
            FIO_CEO = r_data_management_company[0]['data']['management']['name']
            CEO_post = r_data_management_company[0]['data']['management']['post'].capitalize() + " управляющей компании " + management_company
        except:
            st.caption("Не удалось найти информацию об управляющей компании")

    # Получение сокращенного ФИО, обращения и гендера
    nameshort_ceo, greeting = initials_and_greeting(FIO_CEO)

    to_founders, founders_string = get_founders(r_checko)

    # Компоновка данных
    info = {
        "НаимСокр": r_data[0]['data']['name']['short_with_opf'],
        "НаимПолн": r_data[0]['data']['name']['full_with_opf']
        .replace("ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "Общество с ограниченной ответственностью")
        .replace("НЕПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "Непубличное акционерное общество")
        .replace("ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "Публичное акционерное общество")
        .replace("АКЦИОНЕРНОЕ ОБЩЕСТВО", "Акционерное общество"),
        "НаимПолнБезОПФ": r_data[0]['data']['name']['full'],
        "НаимСокрБезОПФ": r_data[0]['data']['name']['short'],
        # "НаимПолнРод": sklonenie(r_data[0]['data']['name']['full_with_opf'], 'gent'),
        "ОПФ": r_data[0]['data']['opf']['full'],
        "ОКВЭД": r_checko['data']["ОКВЭД"]["Наим"] + f' ({r_checko["data"]["ОКВЭД"]["Код"]})',
        # "ОПФРод": sklonenie(r_data[0]['data']['opf']['full'], "gent"),
        "Собств": founders_string,
        "СобствОбращ": to_founders,
        # "Собств": get_founders(r_checko)[1],
        # "СобствОбращ": get_founders(r_checko)[0],
        "ИНН": INN,
        "КПП": r_data[0]['data']['kpp'],
        "ОГРН": r_data[0]['data']['ogrn'],
        "ЮрАдрес": r_checko['data']["ЮрАдрес"]["АдресРФ"],
        "РуководФио": FIO_CEO,
        "РуководФиоДат": sklonenie_fio(FIO_CEO, Case.DATIVE)[0],
        "РуководФиоСокращ": nameshort_ceo,
        "РуководФиоСокращДат": sklonenie_fio(FIO_CEO, Case.DATIVE)[1],
        "РуководДолжн": CEO_post,
        "РуководДолжнДат": sklonenie(CEO_post, 'datv'),
        "РуководОбращ": greeting,
    }

    info["НаимПолнРод"] = sklonenie(info["НаимПолн"], 'gent')

    return info, r_data, r_checko


# Получение данных по отчетности из ГИРБО
@st.cache_data
def get_fs(INN, token_checko):
    r_checko_fs = requests.get(f'https://api.checko.ru/v2/finances?key={token_checko}&inn={INN}').json()
    df_fs = pd.DataFrame.from_dict(r_checko_fs["data"], orient="columns")
    df_fs = df_fs.fillna(0)/1000
    fs_codes_map = {
  "1100": "Итого внеоборотных активов",
  "1110": "Нематериальные активы",
  "1120": "Результаты исследований и разработок",
  "1130": "Нематериальные поисковые активы",
  "1140": "Материальные поисковые активы",
  "1150": "Основные средства",
  "1160": "Доходные вложения в материальные ценности",
  "1170": "Финансовые вложения",
  "1180": "Отложенные налоговые активы",
  "1190": "Прочие внеоборотные активы",
  "1200": "Итого оборотных активов",
  "1210": "Запасы",
  "1220": "Налог на добавленную стоимость по приобретенным ценностям",
  "1230": "Дебиторская задолженность",
  "1240": "Финансовые вложения (за исключением денежных эквивалентов)",
  "1250": "Денежные средства и денежные эквиваленты",
  "1260": "Прочие оборотные активы",
  "1300": "Итого капитал",
  "1310": "Уставный капитал (складочный капитал, уставный фонд, вклады товарищей)",
  "1320": "Собственные акции, выкупленные у акционеров",
  "1340": "Переоценка внеоборотных активов",
  "1350": "Добавочный капитал (без переоценки)",
  "1360": "Резервный капитал",
  "1370": "Нераспределенная прибыль (непокрытый убыток)",
  "1400": "Итого долгосрочных обязательств",
  "1410": "Долгосрочные заемные средства",
  "1420": "Отложенные налоговые обязательства",
  "1430": "Оценочные обязательства",
  "1450": "Прочие долгосрочные обязательства",
  "1500": "Итого краткосрочных обязательств",
  "1510": "Краткосрочные заемные обязательства",
  "1520": "Краткосрочная кредиторская задолженность",
  "1530": "Доходы будущих периодов",
  "1540": "Оценочные обязательства",
  "1550": "Прочие краткосрочные обязательства",
  "1600": "Баланс (актив)",
  "1700": "Баланс (пассив)",
  "2100": "Валовая прибыль (убыток)",
  "2110": "Выручка",
  "2120": "Себестоимость продаж",
  "2200": "Прибыль (убыток) от продаж",
  "2210": "Коммерческие расходы",
  "2220": "Управленческие расходы",
  "2300": "Прибыль (убыток) до налогообложения",
  "2310": "Доходы от участия в других организациях",
  "2320": "Проценты к получению",
  "2330": "Проценты к уплате",
  "2340": "Прочие доходы",
  "2350": "Прочие расходы",
  "2400": "Чистая прибыль (убыток)",
  "2410": "Налог на прибыль",
  "2411": "Текущий налог на прибыль",
  "2412": "Отложенный налог на прибыль",
  "2421": "Постоянные налоговые обязательства",
  "2430": "Изменение отложенных налоговых обязательств",
  "2450": "Изменение отложенных налоговых активов",
  "2460": "Прочее",
  "2500": "Совокупный финансовый результат периода",
  "2510": "Результат от переоценки внеоборотных активов, не включаемый в чистую прибыль (убыток) периода",
  "2520": "Результат от прочих операций, не включаемый в чистую прибыль (убыток) периода",
  "2530": "Налог на прибыль от операций, результат которых не включается в чистую прибыль (убыток) периода",
  "2900": "Базовая прибыль (убыток) на акцию",
  "2910": "Разводненная прибыль (убыток) на акцию",
  "3100": "Величина капитала на 31 декабря года, предшествующего предыдущему",
  "3200": "Величина капитала на 31 декабря предыдущего года",
  "3210": "Увеличение капитала - всего, за предыдущий год",
  "3211": "Чистая прибыль, за предыдущий год",
  "3212": "Переоценка имущества, за предыдущий год",
  "3213": "Доходы, относящиеся непосредственно на увеличение капитала, за предыдущий год",
  "3214": "Дополнительный выпуск акций, за предыдущий год",
  "3215": "Увеличение номинальной стоимости акций, за предыдущий год",
  "3216": "Реорганизация юридического лица, за предыдущий год",
  "3220": "Уменьшение капитала - всего, за предыдущий год",
  "3221": "Убыток, за предыдущий год",
  "3222": "Переоценка имущества, за предыдущий год",
  "3223": "Расходы, относящиеся непосредственно на уменьшение капитала, за предыдущий год",
  "3224": "Уменьшение номинальной стоимости акций, за предыдущий год",
  "3225": "Уменьшение количества акций, за предыдущий год",
  "3226": "Реорганизация юридического лица, за предыдущий год",
  "3227": "Дивиденды, за предыдущий год",
  "3230": "Изменение добавочного капитала, за предыдущий год",
  "3240": "Изменение резервного капитала, за предыдущий год",
  "3300": "Величина капитала на 31 декабря отчетного года",
  "3310": "Увеличение капитала - всего, за отчетный год",
  "3311": "Чистая прибыль, за отчетный год",
  "3312": "Переоценка имущества, за отчетный год",
  "3313": "Доходы, относящиеся непосредственно на увеличение капитала, за отчетный год",
  "3314": "Дополнительный выпуск акций, за отчетный год",
  "3315": "Увеличение номинальной стоимости акций, за отчетный год",
  "3316": "Реорганизация юридического лица, за отчетный год",
  "3320": "Уменьшение капитала - всего, за отчетный год",
  "3321": "Убыток, за отчетный год",
  "3322": "Переоценка имущества, за отчетный год",
  "3323": "Расходы, относящиеся непосредственно на уменьшение капитала, за отчетный год",
  "3324": "Уменьшение номинальной стоимости акций, за отчетный год",
  "3325": "Уменьшение количества акций, за отчетный год",
  "3326": "Реорганизация юридического лица, за отчетный год",
  "3327": "Дивиденды, за отчетный год",
  "3330": "Изменение добавочного капитала, за отчетный год",
  "3340": "Изменение резервного капитала, за отчетный год",
  "3400": "Капитал всего до корректировок",
  "3401": "Нераспределенная прибыль (непокрытый убыток) до корректировок",
  "3402": "Другие статьи капитала, по которым осуществлены корректировки до корректировок",
  "3410": "Корректировка в связи с изменением учетной политики",
  "3411": "Корректировка в связи с изменением учетной политики",
  "3412": "Корректировка в связи с изменением учетной политики",
  "3420": "Корректировка в связи с исправлением ошибок",
  "3421": "Корректировка в связи с исправлением ошибок",
  "3422": "Корректировка в связи с исправлением ошибок",
  "3500": "Капитал - всего после корректировок",
  "3501": "Нераспределенная прибыль (непокрытый убыток) после корректировок",
  "3502": "Другие статьи капитала, по которым осуществлены корректировки после корректировок",
  "3600": "Чистые активы",
  "4100": "Сальдо денежных потоков от текущих операций",
  "4110": "Поступления - всего",
  "4111": "От продажи продукции, товаров, работ и услуг",
  "4112": "Арендных платежей, лицензионных платежей, роялти, комиссионных и иных аналогичных платежей",
  "4113": "От перепродажи финансовых вложений",
  "4119": "Прочие поступления",
  "4120": "Платежи - всего",
  "4121": "Поставщикам (подрядчикам) за сырье, материалы, работы, услуги",
  "4122": "В связи с оплатой труда работников",
  "4123": "Проценты по долговым обязательствам",
  "4124": "Налога на прибыль организаций",
  "4129": "Прочие платежи",
  "4200": "Сальдо денежных потоков от инвестиционных операций",
  "4210": "Поступления - всего",
  "4211": "От продажи внеоборотных активов (кроме финансовых вложений)",
  "4212": "От продажи акций других организаций (долей участия)",
  "4213": "От возврата предоставленных займов, от продажи долговых ценных бумаг (прав требования денежных средств к другим лицам)",
  "4214": "Дивидендов, процентов по долговым финансовым вложениям и аналогичных поступлений от долевого участия в других организациях",
  "4219": "Прочие поступления",
  "4220": "Платежи - всего",
  "4221": "В связи с приобретением, созданием, модернизацией, реконструкцией и подготовкой к использованию внеоборотных активов",
  "4222": "В связи с приобретением акций других организаций (долей участия)",
  "4223": "В связи с приобретением долговых ценных бумаг (прав требования денежных средств к другим лицам), предоставление займов другим лицам",
  "4224": "Процентов по долговым обязательствам, включаемым в стоимость инвестиционного актива",
  "4229": "Прочие платежи",
  "4300": "Сальдо денежных потоков от финансовых операций",
  "4310": "Поступления - всего",
  "4311": "Получение кредитов и займов",
  "4312": "Денежных вкладов собственников (участников)",
  "4313": "От выпуска акций, увеличения долей участия",
  "4314": "От выпуска облигаций, векселей и других долговых ценных бумаг и др.",
  "4319": "Прочие поступления",
  "4320": "Платежи - всего",
  "4321": "Собственникам (участникам) в связи с выкупом у них акций (долей участия) организации или их выходом из состава участников",
  "4322": "На уплату дивидендов и иных платежей по распределению прибыли в пользу собственников (участников)",
  "4323": "В связи с погашением (выкупом) векселей и других долговых ценных бумаг, возврат кредитов и займов",
  "4329": "Прочие платежи",
  "4400": "Сальдо денежных потоков за отчетный период",
  "4450": "Остаток денежных средств и денежных эквивалентов на начало отчетного периода",
  "4490": "Величина влияния изменений курса иностранной валюты по отношению к рублю",
  "4500": "Остаток денежных средств и денежных эквивалентов на конец отчетного периода",
  "6100": "Остаток средств на начало отчетного года",
  "6200": "Поступило средств - всего",
  "6210": "Вступительные взносы",
  "6215": "Членские взносы",
  "6220": "Целевые взносы",
  "6230": "Добровольные имущественные взносы и пожертвования",
  "6240": "Прибыль от предпринимательской деятельности организации",
  "6250": "Прочие",
  "6300": "Использовано средств - всего",
  "6310": "Расходы на целевые мероприятия",
  "6311": "Социальная и благотворительная помощь",
  "6312": "Проведение конференций, совещаний, семинаров и т. п.",
  "6313": "Иные мероприятия",
  "6320": "Расходы на содержание аппарата управления",
  "6321": "Расходы, связанные с оплатой труда (включая начисления)",
  "6322": "Выплаты, не связанные с оплатой труда",
  "6323": "Расходы на служебные командировки и деловые поездки",
  "6324": "Содержание помещений, зданий, автомобильного транспорта и иного имущества (кроме ремонта)",
  "6325": "Ремонт основных средств и иного имущества",
  "6326": "Прочие",
  "6330": "Приобретение основных средств, инвентаря и иного имущества",
  "6350": "Прочие",
  "6400": "Остаток средств на конец отчетного года"
}
    df_fs["Показатель"] = df_fs.index.map(fs_codes_map)
    return df_fs, r_checko_fs


# Поиск в словаре по названию
def find_item_in_dict(search_dict, field):
    """
    Takes a dict with nested lists and dicts,
    and searches all dicts for a key of the field
    provided.
    """
    fields_found = []

    for key, value in search_dict.items():

        if key == field:
            fields_found.append(value)

        elif isinstance(value, dict):
            results = find_item_in_dict(value, field)
            for result in results:
                fields_found.append(result)

        elif isinstance(value, list):
            for item in value:
                if isinstance(item, dict):
                    more_results = find_item_in_dict(item, field)
                    for another_result in more_results:
                        fields_found.append(another_result)

    return fields_found


def main():
    st.title("Заполнение docx файла")
    delimiters = r'[,\n; ]+'
    INN = list(map(str.strip, re.split(delimiters, str(st.text_input("Введите ИНН", help="Если несколько ИНН - перечислите")))))
    get_data_button = st.toggle("Получить данные по ИНН")
    get_fs_button = st.toggle("Получить данные о фин. отчетности")
    # st.write(sklonenie(INN)) # Проверка как работает склонение
    info = []

    if get_data_button:
        for inn in INN:
            info_, r_data, r_checko = get_data(inn)
            info.append(info_)
        with st.expander("Полученные о компаниях данные"):
            # st.write(info)
            # st.dataframe(info)
            st.dataframe(pd.DataFrame(info).T, use_container_width=True)
            # st.write(find_item_in_dict(r_data[0], "okved"))

    # Изобразить отчетность из чеко

    if get_fs_button:
        df_fs_list = []
        for inn in INN:
            df_fs_, r_checko_fs_ = get_fs(inn, token_checko)
            # if not df_fs_.empty and r_checko_fs_:
            try:
                df_fs_["Организация"] = r_checko_fs_["company"]["НаимСокр"]
                df_fs_["ИНН"] = r_checko_fs_["company"]["ИНН"]
                last_key = list(r_checko_fs_["bo.nalog.ru"]["Отчет"])[-1] # возвращает последний key в словаре
                df_fs_["Ссылка"] = r_checko_fs_["bo.nalog.ru"]["Отчет"][last_key]
                girbo_id = r_checko_fs_["bo.nalog.ru"]["ID"]
                df_fs_["Ссылка ГИРБО"] = f"https://bo.nalog.ru/organizations-card/{girbo_id}"
                df_fs_list.append(df_fs_)
            # else:
            except:
                pass


        df_fs = pd.concat(df_fs_list).fillna(0).sort_index(axis=1)
        filter_mask_pivot = ["Выручка", "Себестоимость продаж", "Чистая прибыль (убыток)", "Баланс (актив)", "Чистые активы",
                             "Основные средства", "Запасы", "Финансовые вложения", "Финансовые вложения (за исключением денежных эквивалентов)", "Долгосрочные заемные средства", "Краткосрочные заемные обязательства"]
        df_fs_pivot = (df_fs[df_fs["Показатель"].isin(filter_mask_pivot)].
                       groupby(["Организация", "ИНН", "Показатель"]).sum().reset_index())

        df_fs_pivot["Показатель"] = pd.Categorical(df_fs_pivot["Показатель"], categories=filter_mask_pivot, ordered=True) # Столбец переводится в категорию, чтобы сортировать так, как задано в списке
        df_fs_pivot = df_fs_pivot.sort_values(by=["Организация", "Показатель"]).reset_index(drop=True)

        # Демонстрация и скачивание отчетности
        with st.expander("Доступная финансовая отчетность, тыс. руб."):
            st.dataframe(df_fs, use_container_width=True)
            st.dataframe(df_fs_pivot, use_container_width=True, hide_index=True, column_config={
                "Ссылка": st.column_config.LinkColumn(),
                "Ссылка ГИРБО": st.column_config.LinkColumn()
            })

            df_fs_buffer = io.BytesIO()
            with pd.ExcelWriter(df_fs_buffer, engine='xlsxwriter', date_format='dd.mm.yy',
                                datetime_format='dd.mm.yy') as writer:
                workbook = writer.book
                worksheets = workbook.worksheets()
                format1 = workbook.add_format(
                        {'align': 'top', 'num_format': '#,##0', 'font_name': 'Roboto', 'font_size': '9', 'text_wrap': True})

                df_fs_pivot.to_excel(writer, sheet_name="Ключевые показатели", index=False)
                df_fs.to_excel(writer, sheet_name="Вся отчетность")

                for worksheet in worksheets:
                    worksheet.set_column(0, 60, 13, format1)

                writer.close()

                st.download_button(
                    label="Выгрузить в Excel фин. отчетность",
                    data=df_fs_buffer,
                    file_name=f'Фин. отчетность.xlsx',
                    mime='application/vnd.ms-excel'
                )



    # File uploader allows user to add file
    uploaded_file = st.file_uploader("Загрузите шаблон *docx*", type=['docx'])
    # uploaded_file2 = st.file_uploader("Загрузите данные переменных *docx*", type=['docx'])

    if uploaded_file is not None:
        # Convert the uploaded file to a bytes buffer
        file_buffer = io.BytesIO(uploaded_file.read())

        # Create a DocxTemplate object using the buffer
        doc = DocxTemplate(file_buffer)
        # full_text = []
        # # Извлечение текста из таблиц
        # for table in document.tables:
        #     for row in table.rows:
        #         for cell in row.cells:
        #             full_text.append(cell.text)
        # # Извлечение текста из параграфов
        # for p in document.paragraphs:
        #     full_text.append(p.text)
        #
        # text = "\n".join(full_text)
        # # st.write(text)
        # matches = re.findall(r"\{\{(.*?)}}", text)
        variables = sorted(doc.get_undeclared_template_variables())
        context = pd.DataFrame(columns=variables)
        df_info = pd.DataFrame(info)

        # df_info.loc[len(df_info)] = [0 * len(df_info.columns)]
        if df_info.empty:
            df_info = pd.concat([df_info, pd.DataFrame(["XXX"])], ignore_index=True)
        df_info["Сегодня"] = date.today().strftime("%d.%m.%Y")
        df_info["ГодТек"] = date.today().strftime("%Y")
        df_info["ГодПред"] = (df_info["ГодТек"].astype(int)-1).astype(str)
        df_info["ГодПроверки"] = df_info["ГодПред"]
        df_info["ГодПроверкиПред"] = (df_info["ГодПроверки"].astype(int) - 1).astype(str)


        df = fill_empty_dataframe(context, df_info)

        render_df = st.data_editor(df, use_container_width=True, height=320, num_rows="dynamic")

        render_map = render_df.to_dict(orient='records')
        # st.write(render_map)

        files = 0
        for record in render_map:
        # Render the template with the context
            doc.render(record)

            # Save the rendered document to a buffer
            rendered_docx_buffer = io.BytesIO()
            doc.save(rendered_docx_buffer)

            # Reset buffer position to the beginning
            rendered_docx_buffer.seek(0)
            files += 1
            # Let the user download the rendered DOCX
            st.download_button(
                label=f"Скачать заполненный файл под записью - {files}",
                data=rendered_docx_buffer,
                file_name=uploaded_file.name,
                key=files,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )



if __name__ == "__main__":
    main()

