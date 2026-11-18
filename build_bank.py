"""Собирает банк тем (Excel) из выдачи + ручные кластеры тем."""
import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

import glob

SRC = sorted(glob.glob('output/yt_research_*.xlsx'))[-1]  # последний прогон
STAMP = SRC.split('_')[-1][:-5]
OUT = f'output/bank_tem_{STAMP}.xlsx'

s = pd.read_excel(SRC, sheet_name='Выдача')
s = s.sort_values('views', ascending=False).drop_duplicates(['niche', 'lang', 'video_id'])
a = s['lang_audio'].fillna('').str[:2]
s = s[(a == s['lang']) | (a == '')]
hits = s[(s.views >= 50000) & (s.views_per_sub.fillna(0) >= 5) & (s.channel_subs.fillna(1e9) < 1_500_000)]
hits = hits.sort_values(['niche', 'lang', 'views_per_sub'], ascending=[True, True, False])
hits = hits[['niche', 'lang', 'title', 'views', 'views_per_sub', 'channel', 'channel_subs',
             'duration_s', 'age_days', 'published', 'url', 'thumbnail', 'query']].copy()
hits['duration_s'] = hits['duration_s'] // 60
hits = hits.rename(columns={'views_per_sub': 'просмотры/подписчики', 'duration_s': 'длительность_мин'})

FAM = [
    ("Невестка/свекровь против пожилого родителя", "ES ★★★, EN ★★, DE ★★",
     "Su nuera la corrió en Nochebuena… sin saber que era la dueña de la fábrica (1M, 8k подп.)\nMi nuera me echó del coche de boda; retiré los 88 coches que yo pagué",
     "Рассказ от лица бабушки/дедушки; невестка унижает → выясняется, что дом/ранчо/бизнес на имени пожилого",
     "Сменить рассказчика: невестка сама рассказывает, как ошиблась; «я подслушала разговор невестки с её матерью»"),
    ("Наследство и жадные родственники", "EN ★★★, ES ★★★, DE ★★★",
     "Family Laughed at Her $1 Inheritance—Next Morning, a Second Lawyer… (3M, 32k)\nSus Hijos Pelearon su Herencia Frente a Ella… Sin Saber que Ya Era del Nieto (1M, 5k)\nSie teilten Opas Erbe auf und überließen mir den blinden Hund. Im Halsband steckte ein Zettel (334k)",
     "Смешное/«пустое» наследство (1$, старый дом, собака, ранчо) оказывается главным; второй адвокат / второе завещание",
     "Предмет-загадка в наследстве (ключ, кассета, рецепт); наследство достаётся сиделке/соседу; «завещание прочитали на поминках вслух»"),
    ("Неблагодарные взрослые дети", "EN ★★★, DE ★★, ES ★★",
     "My Son Called Police to Remove Me From “His” House — The Officer Knew My Name From the Title\nMy Son Texted: “Dad Can’t Join the Cruise — My Wife Wants Only Family” (2.6M, 13k)\nMein Sohn forderte: „Gib mir die Hälfte deiner Rente!“",
     "Родителя выгоняют/исключают → он тихо продаёт дом / отменяет оплату / раскрывает документы",
     "Взгляд внука, который встаёт на сторону деда; «я подслушала, что сын называет меня обузой, и записала это»"),
    ("Унижение на семейном событии → разворот", "EN ★★★, DE ★★★, ES ★★",
     "Familie blamiert sie auf der Hochzeit – dann sagt ein Witwer: „Tun Sie so, als gehörten Sie zu mir“\nHer Family Humiliated Her at the Funeral—Now the True Heiress…\nMi hijastra dijo: “No vengas a mi boda”… cancelé la celebración que yo había pagado",
     "Свадьба / похороны / Рождество / юбилей / встреча выпускников: публичное унижение и публичная расплата",
     "Сезонно: Рождество, Новый год, День благодарения. Сейчас ноябрь, самое время для рождественских сюжетов"),
    ("Скрытое богатство / недооценённый человек", "EN ★★★, DE ★★★, ES ★★★",
     "Humillaron al VAGABUNDO del rancho… sin saber que ERA el nuevo DUEÑO (5.8M, 16k)\nBeim Klassentreffen wollten sie sie demütigen – Doch dann erfuhren sie, wem das Luxushotel gehörte!\nHe cheated on his wife not knowing everything he owned belonged to her",
     "Формула «…sin saber que / ohne zu ahnen / not knowing», герой оказывается владельцем",
     "Тайный владелец: уборщица, водитель, бабушка в автобусе; «я 30 лет скрывала от мужа свою зарплату»"),
    ("Миллионер + бедная героиня / ребёнок", "DE ★★★, EN ★★",
     "„Iss du, mein Schatz“ – sie verzichtete aufs Essen, ohne zu ahnen, dass der CEO sie beobachtete (147k при ~0 подп.)\nKellnerin hilft jeden Morgen altem Mann – doch dann betreten seine Anwälte mit 4 Bodyguards das Café\nThe Millionaire’s Grandma Chose the Delivery Girl (7.4M, 40k)",
     "Акт доброты к незнакомцу → он оказывается миллионером/CEO",
     "В DE это самый массовый шаблон, но он уже перегрет. Лучше перевернуть: миллионер проверяет семью, притворяясь бедным"),
    ("Измена (с табу-поворотом)", "ES ★★★, EN ★★",
     "Lo Que Pasó con Mi Yerno Mientras Mi Hija No Estaba (913k, 5k)\nMi Esposo Me Escribió Desde Noruega… y el Marido De Su Amante Nos Expuso\nHusband Cheated Because I Was Disabled. I Divorced Him — And Crushed His Company",
     "Измена + неожиданный участник (зять, подруга, сестра); месть через деньги/бизнес",
     "«Узнала об измене из посылки / почтальона / умных часов», сюжет вокруг технологий"),
    ("Тайна, раскрытая после смерти", "DE ★★, ES ★★, EN ★★",
     "WAHRE GESCHICHTE: 40 Jahre lang log sie mich an. Erst nach ihrem Tod verstand ich warum (372k, 5k)\nVOLVÍ DEL FUNERAL CON UNA HERENCIA SECRETA… Y OÍ EL PLAN DE MI FAMILIA",
     "Письмо / дневник / второй адрес находят после похорон",
     "«Второй муж/жена на похоронах», «ребёнок, о котором никто не знал», «ключ от банковской ячейки»"),
    ("НОВОЕ: Любимчик и козёл отпущения среди детей", "DE ★★★, EN ★★",
     "Mein Vater sagte: „Das Baby deiner Schwester ist das erste Enkelkind!“ (38x к медиане канала)\nMeine Eltern feierten die Verlobung meiner Schwester an meinem Geburtstag–dann fanden sie meinen Brief (39x)\nMy Parents Said You're Not Invited At My Sister's Wedding — 3 Days Later They Begged Me To Pay $38,000 (279k)",
     "Рассказчик — молодой взрослый, родители предпочитают брата или сестру → публичный разворот; 10–15 мин (DE), 50–60 мин (EN)",
     "Привязать к праздникам (Рождество, день рождения, юбилей); «родители продали мой подарок сестре»"),
    ("НОВОЕ: «Перестал(а) платить за семью»", "ES ★★★, DE ★★, EN ★★",
     "Tiraron mis cosas en bolsas negras… así que dejé de pagar su casa (303x)\nEl Año Que Dejé De Pagar La Navidad (94x, 307k)\nNachdem ich alle Rechnungen bezahlt hatte, warf mich meine Schwiegermutter hinaus (43x)",
     "Герой тайно содержал семью → его выгоняют/унижают → он перестаёт платить, и всё рушится",
     "Сезонный вариант: «в этом году я не оплачу Рождество». Сейчас самое время"),
    ("НОВОЕ: Тайное сообщение / шёпот-предупреждение", "EN ★★★, DE ★★",
     "At my sister's baby shower, my husband texted from across the room: Don't eat anything she hands you (35x)\nNach dem Unfall meines Mannes flüsterte ein alter Patient: „Lauf!“ (68k)\nAt my baby shower, my 11 year old niece texted me from across the room",
     "Короткая фраза-предупреждение в названии, интрига с первой секунды",
     "Варианты: записка в кармане, сообщение от незнакомого номера, ребёнок шепчет на ухо"),
    ("НОВОЕ: Бросил(а) → она стала успешной", "ES ★★★, EN ★★",
     "Abandonó a su esposa en silla de ruedas y ella se ganó la lotería (13.7M)\nAbandonó a su esposa en el campo pero ella regresó para ser exitosa (6M)\nEngañó a su esposa sin imaginar que ella se hará millonaria (11.8M)",
     "Предательство → взлёт героини → встреча через годы",
     "Рассказ от лица того, кто бросил («я ушёл от больной жены — через 5 лет…»)"),
]

PRED = [
    ("НОВОЕ: «Если ты родился в 1945–1965…»", "EN ★★★, ES ★★",
     "If You Were Born 1945–1965, Edgar Cayce Believed You Carry a Special Spiritual Task (201x, 554k)\nWere You Born Between 1945 and 1965? Edgar Cayce Revealed Your Final Spiritual Mission (100x)\nWhat the Final Digit of Your Birth Year Reveals About Your Soul’s Past Life",
     "Провидец + поколение зрителя. Аудитория ниши 55+, адресное обращение к ней",
     "Другие провидцы и поколения: «Ванга о рождённых в 60-х», «что Нострадамус сказал о поколении бумеров»"),
    ("НОВОЕ: Зодиак/асцендент как «судьба» (DE)", "DE ★★★",
     "Nostradamus prophezeite es: DIESE 5 Sternzeichen erleben 2027 ihr Schicksalsjahr (195x, 283k)\nDein Aszendent zeigt, wer du sein musstest — alle 12 im Vergleich (115x)\nDein Element sagt mehr über dich als dein Sternzeichen (79x)",
     "Мягкая эзотерика + личное обращение «твой знак / твой асцендент»",
     "Серия по 12 знакам на 2027 год, ролики на 20–25 минут"),
    ("НОВОЕ: Прогнозы ясновидящей на месяц или неделю", "ES ★★★, DE ★★",
     "¡NOVIEMBRE 2026 SERÁ IMPACTANTE! GUERRAS y DESASTRES (50x)\nTODOS EN ALERTA: Las Predicciones Mundiales Más Fuertes de Septiembre 2026 (60x)\nNovember 2026: Die Endphase wird nun eingeläutet (339k)",
     "Регулярная рубрика с датой в названии, к ней возвращаются каждый месяц",
     "Формат-конвейер: «Предсказания на декабрь 2026», «на январь 2027» и дальше каждый месяц"),
    ("Симпсоны предсказали 2027", "EN ★★★, ES ★★★, DE ★★",
     "7 Predicciones de Los Simpson para 2027 que dan MIEDO (1.2M при 1k подп.!)\nThe Simpsons Predicted 2027… And It's Already Happening (3.6M, 25k)\nLas Predicciones de Los Simpson Para 2027 Te Van a Volar la Cabeza (4.3M, 67k)",
     "Подборка «N предсказаний на год» + «уже сбывается»",
     "Привязка к месяцу («ноябрьский блэкаут»), к стране (Испания/Германия), к ЧМ-2026, к ИИ. Пик спроса с ноября по январь, тема ограничена по времени"),
    ("Нострадамус / Ванга / Кейси на 2027", "EN ★★★, ES ★★★, DE ★★",
     "Nostradamus predijo estos 7 desastres para 2026 (887k, 3k)\nNOSTRADAMUS Predijo el Colapso de 3 Naciones Antes de que Termine 2026 (492k, 2k)\nEdgar Cayce's Vision About AMERICA After 2027",
     "Провидец + конкретный год + катастрофа/страна, и обязательно «что уже сбылось»",
     "Перевести с 2026 на 2027: «что сбылось в 2026 и что ждёт в 2027». Это готовый ход на ближайшие 2–3 месяца"),
    ("Предсказания про конкретную страну", "DE ★★★, ES ★★★",
     "Irlmaiers letzte Vision: Vor seinem Tod sah er DAS über Deutschland! (81k, 1k)\nMühlhiasls LETZTE PROPHEZEIUNG! Was soll im Winter 2026 geschehen?!\nLa Profecía de Baba Vanga Sobre México… ¿Claudia Sheinbaum?",
     "Местные провидцы: DE — Irlmaier, Mühlhiasl, Madame Burela, Padre Pio; ES — Parravicini, Hildemaro",
     "В DE много маленьких каналов с хитами на Irlmaier/Mühlhiasl, ниша местная и не перегретая. Для Испании: «профеции об Испании» (в выборке почти нет, нужно проверить отдельно)"),
    ("Зодиак + деньги/удача в 2027", "EN ★★★, ES ★★",
     "Chinese Zodiac 2027: 6 LUCKIEST Signs of the Fire Goat Year (1M, 13k)\nBaba Vanga predijo: tres signos del zodíaco recibirán 10 millones en 2026\n6 Luckiest Chinese Animal Zodiac Signs, That Will Be RICHEST in 2027",
     "Провидец × знак зодиака × деньги; год Огненной Козы начинается в феврале 2027",
     "«Если вы родились в эти годы…» (ES: «Si naciste en ESTOS años, 2027 viene por tu dinero»)"),
    ("Буква «М» на ладони", "EN ★★, ES ★★",
     "If You Have an “M” on Your Palm, Nostradamus Says Watch Out in 2026! (184k при ~0 подп.)\n¿Tienes una M en Tu Palma? Esto Es Lo que Nostradamus Descubrió (521k, 6k)",
     "Личная проверка зрителя («посмотри на свою ладонь») + провидец",
     "Другие телесные/личные «признаки»: родинка, дата рождения, первая буква имени"),
    ("Библейские пророчества / конец времён", "EN ★★★, ES ★★, DE ★★",
     "Ezekiel 38 Names the Nations That Will Attack Israel — Look at the Map Today\nThree Days of Darkness: The Signs the Saints Said Would Come First (388k, 2k)\nDie Prophezeiung von Padre Pio über Deutschland",
     "Пророчество + сегодняшние новости / карта",
     "Самая большая и конкурентная часть ниши. Лучше заходить через связку с новостями"),
    ("ИИ + пророчества", "EN ★★, ES ★★",
     "Elon Musk’s Grok AI Was Asked “Where Is America in Bible Prophecy?” (370k, 4k)\nExpertos Preguntaron a la IA Grok Dónde Aparece México en las Profecías Bíblicas",
     "Спросили ИИ о пророчестве, ответ «пугающий»",
     "Свежий гибридный угол: «ИИ проанализировал все предсказания Ванги — вот что совпало»"),
    ("Срочное предупреждение на месяц", "DE ★★, ES ★★",
     "Madame Burela: „Haltet bis Dezember 2026 durch!“ (126k, 1k)\nLA ÚNICA VEZ QUE VANGA LLORÓ: LO QUE VIO PARA OCTUBRE Y NOVIEMBRE\n¡URGENTE! …una clarividente advierte sobre NOVIEMBRE",
     "Привязка к ближайшему месяцу создаёт срочность",
     "Серия «что предсказано на декабрь / январь / февраль»: каждый месяц новый ролик"),
    ("Политика и лидеры (Трамп и др.)", "EN ★★, ES ★★",
     "Baba Vanga’s 2026 Trump Prophecy (1.4M, 8k)\nEdgar Cayce predijo el punto de inflexión de Trump en 2026 (1M, 14k)",
     "Провидец + действующий политик",
     "Добавить местных политиков: Sheinbaum, Milei, Sánchez, Merz"),
]

FORMULAS = [
    ("«Прямая речь» – кто сказал + поворот", "„Dürfen wir in Ihrer Scheune schlafen?“, fragte die Witwe – Seine Antwort veränderte ihr Leben", "Семейные, все языки"),
    ("[Унижение]… sin saber que / ohne zu ahnen / not knowing [скрытая сила]", "Su suegra la corrió en Nochebuena… sin saber que era la dueña de la fábrica", "Семейные, все языки"),
    ("От 1-го лица: «Мой сын сказал X. Я сделал Y…»", "Mi hijo dijo: «Es hora de que te mudes». Vendí la casa mientras trabajaba…", "Семейные: EN/ES, аудитория 50+"),
    ("Конкретная сумма или предмет", "Family Laughed at Her $1 Inheritance / Mansión de $3 / 88 coches", "Семейные"),
    ("Многоточие-обрыв в конце", "…Aber als ich vom Tisch aufstand …", "Семейные"),
    ("[Провидец] + [год] + [число] + страх", "6 Profecías ESCALOFRIANTES de Nostradamus para 2026 (El Mundo NO Está Listo)", "Предсказания"),
    ("«…и это уже сбывается / already happening»", "The Simpsons Predicted 2027… And It's Already Happening", "Предсказания"),
    ("Личная проверка зрителя", "If You Have an “M” on Your Palm… / Si naciste en ESTOS años…", "Предсказания"),
    ("Страна/регион + провидец", "Irlmaiers letzte Vision: … über Deutschland!", "Предсказания, DE/ES"),
    ("КАПС на 1–3 ключевых словах", "NOSTRADAMUS Predijo el Colapso de 3 Naciones", "Оба"),
]

THUMBS = [
    ("Семейные DE/ES/EN (каналы мести)", "«Стена текста»: вся завязка 8–12 строками, белый + жёлтый/оранжевый на тёмном или красном фоне, справа эмоциональное лицо (ИИ). Внизу плашка-клиффхэнгер: «Was sie dann vorhatte, wusste niemand», «LLEGARON A MI PUERTA…», «MOM GASPED». Цифры выделены цветом: 6.600 $ PRO MONAT, $486,000",
     "У слабых роликов тех же каналов такой же стиль, значит это обязательный минимум ниши, а не преимущество. Решает сюжет. Выигрывает текст, где есть конкретная сумма и реплика злодея в кавычках"),
    ("Семейные ES (большие каналы «película»)", "Кинематографичная ИИ-сцена БЕЗ текста: слабый с виду герой (старик в кресле-качалке, бабушка с лотком) + угроза (вооружённые бандиты) + стрелка на скрытую деталь. 6–7 млн просмотров",
     "Вся история читается в одном кадре: контраст «беззащитный против угрозы» обещает разворот «sin imaginar quién era»"),
    ("Семейные EN (Reddit)", "Скриншот поста в стиле соцсети: аватар, галочка, одна фраза-крючок («Don't eat anything she hands you»), счётчики 99+",
     "Работает для формата «реальная история из Reddit» и для крючка «тайное сообщение»"),
    ("Семейные EN (мини-драмы)", "Постер фильма: золотой шрифт с засечками, героиня в центре, богатая семья вокруг, фэнтези-свет",
     "Работает только для формата «мини-драма / фильм». Такой же постер без конкретного крючка («Forgotten Princess») не сработал"),
    ("Слабые семейные превью", "Стоп-кадр из фильма с логотипом «Lifetime», нейтральные лица без конфликта, сцена без текста и без угрозы",
     "Нет конфликта и нет вопроса, который хочется закрыть"),
    ("Предсказания (все языки)", "Крупное лицо провидца (Нострадамус, Ванга, Кейси) с драматичным светом + апокалиптический фон (огонь, горящий город, метеор) + 2–4 огромных слова с ГОДОМ или МЕСЯЦЕМ: «WHY 2027?», «END OF NOVEMBER WILL BE DEVASTATING», «2026 ES EL AÑO»",
     "У слабых роликов того же стиля нет даты или срочности («HIDE YOUR PALM», «HE SAW THIS 500 YEARS AGO», «THE COMING STORM»). Год или месяц на превью — главный рычаг"),
    ("Предсказания: личное обращение", "Огромные цифры поколения «1945 & 1965», «5 STERNZEICHEN 2027» с иконками знаков, вопрос «WER BIST DU WIRKLICH?»",
     "Зритель ищет себя на картинке"),
    ("Предсказания: шок-образ", "Трамп в гробу рядом с Нострадамусом (ES, 7.3M и 834k), Симпсоны в мультяшном кадре с узнаваемыми политиками",
     "Сильный шок, но есть риск: зрители жалуются на кликбейт, а YouTube может ограничить показ"),
    ("Предсказания DE: эзотерика-интервью", "Две говорящие головы + короткий вопрос («LOCKDOWN 2027?»)",
     "Работает для немецких каналов-интервью с медиумами"),
]

COMMENTS = [
    ("Семейные", "Аудитория 55–75 лет пишет свои истории о неблагодарных детях («Мне 62, из 4 детей поздравил только сын», «fui maltratada por mis hijos»)",
     "Темы «неблагодарные дети» и «одинокая старость» дают самое сильное личное вовлечение. В конце ролика спрашивайте: «А с вами такое было?»"),
    ("Семейные", "Массовые просьбы «Segunda parte», «Next part please» (ES/EN); каналы закрепляют «parte 2 ya viene»",
     "Стратегия сериала: клиффхэнгер → часть 2 через 3–5 дней. При этом хвалят и «полные истории, а не на половину» — первая часть должна всё равно давать развязку"),
    ("Семейные", "Усталость от штампов: «no yelling or slapping… or rich and poor storyline. Very unique» (390 лайков)",
     "Возможность: сюжеты без пощёчин и без «миллионера и бедной», где развязка строится на уме героя"),
    ("Семейные", "Спорный персонаж (Nathan) собрал тысячи лайков под комментариями «я бы бросила его»",
     "Морально спорный второстепенный герой = споры в комментариях = продвижение"),
    ("Семейные ES", "Сильная религиозность: «Dios nunca nos abandona», «Amén»",
     "В испанских историях вплетать веру и мотив «Dios jamás abandona». Это работает и в превью"),
    ("Семейные DE", "Обожают животных: под историей про слепого пса десятки комментариев «пёс — лучшее наследство»",
     "Добавлять в сюжеты животное как моральный ориентир («собака сразу почувствовала плохих людей»)"),
    ("Предсказания DE/EN", "Люди пишут свой знак и надежды: «Ich bin Wassermann, hoffe auf Ruhe und Gesundheit»",
     "Ролики по знакам + в начале «напиши свой знак в комментариях». Это готовое вовлечение"),
    ("Предсказания все языки", "В топе скепсис-конспирология: «это не предсказания, а их планы» / «predictive programming» / «no es predicción, es plan ejecutado»",
     "Новый угол: «Симпсоны: предсказание или план?». Рамка «элиты знают заранее» совпадает с мнением аудитории"),
    ("Предсказания EN", "ИИ вызывает сильный страх и много комментариев (4 000 лайков у комментария под «10 Times AI Said…»)",
     "Связки «ИИ + пророчество», «Нострадамус предсказал ИИ?»"),
    ("Предсказания ES", "Закреп «какое пророчество поразило больше всего? Напиши номер 1–6» набрал 3 163 лайка",
     "Нумеровать пророчества в ролике и просить ответить цифрой"),
    ("Оба", "Жалобы: навязчивая музыка, движущийся фон, реклама поверх концовки, кликбейт-превью",
     "Тихий фон, без рекламы в последние 2 минуты, превью должно соответствовать содержанию"),
]

COLS = ['Тема / кластер', 'Где сильна (★ = сила)', 'Примеры хитов (просмотры, подписчики канала)', 'Что работает', 'Новые углы подачи']
summary = pd.read_excel(SRC, sheet_name='Ниши_сводка').drop(columns=['shorts_share'])

with pd.ExcelWriter(OUT, engine='openpyxl') as xw:
    pd.DataFrame(FAM, columns=COLS).to_excel(xw, sheet_name='Темы_семейные', index=False)
    pd.DataFrame(PRED, columns=COLS).to_excel(xw, sheet_name='Темы_предсказания', index=False)
    pd.DataFrame(FORMULAS, columns=['Формула названия', 'Пример', 'Где применять']).to_excel(xw, sheet_name='Формулы_названий', index=False)
    pd.DataFrame(THUMBS, columns=['Где', 'Что на превью хитов', 'Вывод']).to_excel(xw, sheet_name='Превью', index=False)
    pd.DataFrame(COMMENTS, columns=['Ниша', 'Что пишут зрители', 'Что делать']).to_excel(xw, sheet_name='Комментарии_инсайты', index=False)
    cm = pd.read_csv(f'output/comments_{STAMP}.csv')
    cm = cm[cm.is_request | (cm.likes >= 50)].sort_values(['niche', 'lang', 'likes'], ascending=[True, True, False])
    cm.drop(columns=['is_request']).to_excel(xw, sheet_name='Комментарии_топ', index=False)
    summary.to_excel(xw, sheet_name='Сравнение_ниш', index=False)
    hits.to_excel(xw, sheet_name='Хиты_малых_каналов', index=False)
    ch = pd.read_excel(SRC, sheet_name='Каналы').sort_values(['niche', 'lang', 'outliers'], ascending=[True, True, False])
    ch.to_excel(xw, sheet_name='Конкуренты', index=False)
    d = pd.read_excel(SRC, sheet_name='Все_ролики')
    d = d[(d.outlier_x >= 3) & (d.age_days >= 7) & (d.age_days <= 400)]
    d = d.sort_values(['niche', 'lang', 'outlier_x'], ascending=[True, True, False])
    d = d[['niche', 'lang', 'title', 'outlier_x', 'views', 'channel_median', 'channel', 'duration_s',
           'published', 'url', 'thumbnail']].rename(columns={'outlier_x': 'во_сколько_раз_выше_медианы_канала'})
    d['duration_s'] = d['duration_s'] // 60
    d = d.rename(columns={'duration_s': 'длительность_мин'})
    d.to_excel(xw, sheet_name='Выбросы_конкурентов', index=False)
    for ws in xw.book.worksheets:
        for i, col in enumerate(ws.columns, 1):
            w = max(len(line) for c in col for line in str(c.value or '').split('\n'))
            ws.column_dimensions[get_column_letter(i)].width = min(w + 2, 60)
        for row in ws.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical='top')
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.freeze_panes = 'A2'

print(OUT, len(hits))
print(hits.groupby(['niche', 'lang']).size())
