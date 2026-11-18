# Ниши и стартовые поисковые запросы по языкам.
# Добавляй/убирай запросы свободно: каждый запрос search.list = 100 единиц квоты.

LANGS = {
    "en": {"region": "US"},
    "es": {"region": "ES"},
    "de": {"region": "DE"},
}

NICHES = {
    "family_drama": {
        "en": [
            "family drama story",
            "true life stories",
            "reddit family drama stories",
            "my husband cheated story",
            "mother in law story",
            "inheritance family betrayal story",
            "revenge story family",
            "AITA family story",
            "my son kicked me out story",
            "golden child scapegoat story",
            "I stopped paying for my family story",
        ],
        "es": [
            "historias de la vida real",
            "historias familiares dramáticas",
            "historia de traición familiar",
            "historia suegra nuera",
            "historias reales de infidelidad",
            "historias de herencia familia",
            "mi nuera me humilló historia",
            "hijo ingrato historia real",
            "dejé de pagar a mi familia historia",
        ],
        "de": [
            "wahre Lebensgeschichten",
            "Familiendrama Geschichte",
            "Schwiegermutter Geschichte",
            "Geschichten aus dem Leben",
            "Reddit Geschichten deutsch",
            "Erbe Familie Verrat Geschichte",
            "Rache Geschichte Familie",
            "undankbare Kinder Geschichte",
            "Lieblingskind Geschichte Eltern",
        ],
    },
}

SEARCH_MONTHS_BACK = 12        # искать ролики за последние N месяцев
SEARCH_RESULTS = 50            # результатов на запрос (макс 50 за 100 единиц)
TOP_CHANNELS_PER_GROUP = 15    # сколько каналов на (нишу x язык) разбирать глубоко
UPLOADS_PER_CHANNEL = 300      # сколько последних роликов канала выгружать
OUTLIER_THRESHOLD = 3.0        # ролик-выброс: просмотры >= N x медиана канала
SMALL_CHANNEL_SUBS = 50_000    # "маленький" канал для оценки входа в нишу
SHORTS_MAX_SECONDS = 180       # всё, что короче, считается Shorts и отбрасывается
# Только длинные ролики: medium = 4–20 мин, long = 20+ мин (каждый — отдельный поиск, 100 единиц)
SEARCH_DURATIONS = ["medium", "long"]

# Отключённые ниши (чтобы вернуть — перенесите обратно в NICHES):
DISABLED_NICHES = {
    "predictions": {
        "en": [
            "predictions 2027",
            "Baba Vanga predictions",
            "Nostradamus predictions",
            "horoscope 2027",
            "tarot reading prediction",
            "psychic predictions",
            "prophecy end times",
            "future predictions world",
            "zodiac signs luck",
        ],
        "es": [
            "predicciones 2027",
            "profecías",
            "Baba Vanga predicciones",
            "Nostradamus predicciones",
            "horóscopo 2027",
            "tarot predicciones",
            "vidente predicciones",
        ],
        "de": [
            "Vorhersagen 2027",
            "Prophezeiungen",
            "Baba Wanga Vorhersagen",
            "Nostradamus Vorhersagen",
            "Horoskop 2027",
            "Tarot Vorhersage",
            "Hellseher Vorhersagen",
        ],
    },
}
