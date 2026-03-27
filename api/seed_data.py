"""
Default MA + COM reference data for Micro Embroidery.
Loaded on startup if the database tables are empty.
Source: AA M_Program and Com List.xlsx (57 combos, 12 MAs)
"""

DEFAULT_MA_REFERENCE = [
    {"size_normalized": "105x25", "size_display": "105 x 25", "ma_number": "MA76230"},
    {"size_normalized": "110x25", "size_display": "110 x 25", "ma_number": "MA50334"},
    {"size_normalized": "110x35", "size_display": "110 x 35", "ma_number": "MA50310"},
    {"size_normalized": "110x40", "size_display": "110 x 40", "ma_number": "MA50340"},
    {"size_normalized": "120x30", "size_display": "120 x 30", "ma_number": "MA50284"},
    {"size_normalized": "135x30", "size_display": "135 x 30", "ma_number": "MA50330"},
    {"size_normalized": "140x25", "size_display": "140 x 25", "ma_number": "MA53345"},
    {"size_normalized": "140x30", "size_display": "140 x 30", "ma_number": "MA50338"},
    {"size_normalized": "140x35", "size_display": "140 x 35", "ma_number": "MA53344"},
    {"size_normalized": "140x40", "size_display": "140 x 40", "ma_number": "MA50254"},
    {"size_normalized": "150x30", "size_display": "150 x 30", "ma_number": "MA55887"},
    {"size_normalized": "150x35", "size_display": "150 x 35", "ma_number": "MA55451"},
]

DEFAULT_COM_REFERENCE = [
    {"ma_number": "MA76230", "com_number": 1, "fabric_colour": "Rot", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA50334", "com_number": 1, "fabric_colour": "Graphit", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Grau 0108 I"},
    {"ma_number": "MA50334", "com_number": 2, "fabric_colour": "Convoy", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA50334", "com_number": 3, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA50334", "com_number": 4, "fabric_colour": "Deep Sky#4348", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA50334", "com_number": 5, "fabric_colour": "Weiß", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50310", "com_number": 1, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50340", "com_number": 1, "fabric_colour": "Royalblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3544 I"},
    {"ma_number": "MA50340", "com_number": 2, "fabric_colour": "Weiß", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50340", "com_number": 3, "fabric_colour": "Convoy", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA50340", "com_number": 4, "fabric_colour": "Schwarz", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50284", "com_number": 1, "fabric_colour": "Schwarz", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA50284", "com_number": 2, "fabric_colour": "Warngelb", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA50330", "com_number": 1, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53345", "com_number": 1, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53345", "com_number": 2, "fabric_colour": "Weiß", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA53345", "com_number": 3, "fabric_colour": "Weiß", "embroidery_colour": "Blau 3554 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA53345", "com_number": 4, "fabric_colour": "Rot", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA53345", "com_number": 5, "fabric_colour": "Convoy", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA53345", "com_number": 6, "fabric_colour": "Convoy", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA53345", "com_number": 7, "fabric_colour": "Rot", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA53345", "com_number": 8, "fabric_colour": "Tiefblau", "embroidery_colour": "Orange 1332 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53345", "com_number": 9, "fabric_colour": "Schwarz", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53345", "com_number": 10, "fabric_colour": "Rot", "embroidery_colour": "Blau 3902 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA53345", "com_number": 11, "fabric_colour": "Schwarz", "embroidery_colour": "Orange 1106 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53345", "com_number": 12, "fabric_colour": "Tiefblau", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53345", "com_number": 13, "fabric_colour": "Rot", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA53345", "com_number": 14, "fabric_colour": "Rot", "embroidery_colour": "Grau 3971 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA53345", "com_number": 15, "fabric_colour": "Convoy", "embroidery_colour": "Grau 3971 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA53345", "com_number": 16, "fabric_colour": "Warngelb", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53345", "com_number": 17, "fabric_colour": "Rot", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Gelb 6010 I"},
    {"ma_number": "MA53345", "com_number": 18, "fabric_colour": "Royalblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53345", "com_number": 19, "fabric_colour": "Schwarz", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53345", "com_number": 20, "fabric_colour": "Weiß", "embroidery_colour": "Blau 3544 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50338", "com_number": 1, "fabric_colour": "Schwarz", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA50338", "com_number": 2, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA50338", "com_number": 3, "fabric_colour": "Tiefblau", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA50338", "com_number": 4, "fabric_colour": "Schwarz", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Orange 1106 I"},
    {"ma_number": "MA50338", "com_number": 5, "fabric_colour": "Convoy", "embroidery_colour": "Grau 0138 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50338", "com_number": 6, "fabric_colour": "Weiß", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53344", "com_number": 1, "fabric_colour": "Convoy", "embroidery_colour": "Grau 3971 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA53344", "com_number": 2, "fabric_colour": "Schwarz", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53344", "com_number": 3, "fabric_colour": "Convoy", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA53344", "com_number": 4, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53344", "com_number": 5, "fabric_colour": "Convoy", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA53344", "com_number": 6, "fabric_colour": "Weiß", "embroidery_colour": "Schwarz 0020 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA53344", "com_number": 7, "fabric_colour": "Tiefblau", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA53344", "com_number": 8, "fabric_colour": "Rot", "embroidery_colour": "Rot 1903 I", "frame_colour": "Schwarz 0020 I"},
    {"ma_number": "MA53344", "com_number": 9, "fabric_colour": "Rot", "embroidery_colour": "Rot 1903 I", "frame_colour": "Weiß 0010 I"},
    {"ma_number": "MA50254", "com_number": 1, "fabric_colour": "Rot", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA50254", "com_number": 2, "fabric_colour": "Rot", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Rot 1903 I"},
    {"ma_number": "MA50254", "com_number": 3, "fabric_colour": "Tiefblau", "embroidery_colour": "Orange 1106 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA50254", "com_number": 6, "fabric_colour": "Rot", "embroidery_colour": "Gelb 6010 I", "frame_colour": "Rot 1903I"},
    {"ma_number": "MA50254", "com_number": 8, "fabric_colour": "Tiefblau", "embroidery_colour": "Weiß 0010 I", "frame_colour": "Blau 3554 I"},
    {"ma_number": "MA55887", "com_number": 1, "fabric_colour": "Warngelb", "embroidery_colour": "Rot 1903 I", "frame_colour": "Gelb 6010 I"},
    {"ma_number": "MA55451", "com_number": 1, "fabric_colour": "Convoy", "embroidery_colour": "Grau 3971 I", "frame_colour": "Grau 0138 I"},
    {"ma_number": "MA55451", "com_number": 2, "fabric_colour": "Weiß", "embroidery_colour": "Rot 1903 I", "frame_colour": "Weiß 0010 I"},
]


def seed_reference_if_empty():
    """Load default MA + COM reference on startup if tables are empty.
    Does NOT overwrite existing user data."""
    from api.database import (
        get_ma_reference, upsert_ma_reference,
        get_com_reference, upsert_com_reference,
    )

    if not get_ma_reference():
        upsert_ma_reference(DEFAULT_MA_REFERENCE)
        print(f"[seed] Loaded {len(DEFAULT_MA_REFERENCE)} default MA mappings")

    if not get_com_reference():
        upsert_com_reference(DEFAULT_COM_REFERENCE)
        print(f"[seed] Loaded {len(DEFAULT_COM_REFERENCE)} default COM entries")
