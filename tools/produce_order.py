"""Presentation-only ordering for the complete binder's Produce sheet."""
import re


# More specific names precede their broader constituents.
PRODUCTS = (
    ("SWEET POTATO", r"sweet potato(?:es)?|red.fleshed sweet potato"),
    ("LEMON GRASS", r"lemon grass"),
    ("GRANADILLA", r"granadilla"),
    ("GOOSEBERRIES", r"gooseberries"),
    ("LOGAN BERRIES", r"logan berries"),
    ("SASKATOON BERRIES", r"saskatoon"),
    ("TAY BERRIES", r"tay berries"),
    ("PEPPER", r"peppers?"),
    ("SQUASH", r"squash|kabocha"),
    ("WATERMELON", r"watermelon"),
    ("CANTALOUPE", r"cantaloupe"),
    ("HONEYDEW", r"honeydew"),
    ("BANANA", r"bananas?"),
    ("APPLE", r"apples?"),
    ("ANISE FENNEL", r"anise fennel"),
    ("APRICOT", r"apricots?"),
    ("ARTICHOKE", r"artichokes?"),
    ("ARUGULA", r"arugula"),
    ("ASPARAGUS", r"asparagus"),
    ("AVOCADO", r"avocados?"),
    ("BASIL", r"basil"),
    ("BEAN SPROUTS", r"bean sprouts?"),
    ("BEANS", r"beans?"),
    ("BEETS", r"beets?"),
    ("BLACKBERRIES", r"blackberries"),
    ("BLUEBERRIES", r"blueberries"),
    ("BOYSENBERRIES", r"boysenberries"),
    ("RASPBERRIES", r"raspberries"),
    ("STRAWBERRIES", r"strawberr(?:y|ies)"),
    ("BOK CHOY", r"bok choy"),
    ("BRUSSEL SPROUTS", r"brussel sprouts"),
    ("BROCCOLINI", r"broccolini"),
    ("BROCCOLI", r"broccoli"),
    ("BUDDHA'S HAND", r"buddha.s hand"),
    ("CABBAGE", r"cabbage"),
    ("CARROTS", r"carrott?s?"),
    ("CAULIFLOWER", r"cauliflower"),
    ("CELERY ROOT", r"celery root"),
    ("CELERY", r"celery"),
    ("CHARD", r"(?:swiss )?chard"),
    ("CHERRIES", r"cherries"),
    ("CHIVES", r"chives"),
    ("CILANTRO", r"cilantro"),
    ("CLEMENTINE", r"clementine"),
    ("COCONUT", r"coconut"),
    ("COLLARD GREENS", r"collard greens"),
    ("CORN", r"corn"),
    ("CUCUMBER", r"cucumbers?"),
    ("CURRANTS", r"currents|currants"),
    ("DAIKON", r"daikon"),
    ("DATES", r"dates?"),
    ("DILL", r"dill"),
    ("DRAGON FRUIT", r"dragon fruit|pitahaya"),
    ("EGGPLANT", r"eggplant"),
    ("ENDIVE", r"endive"),
    ("FEIJOA", r"feijoa"),
    ("FIGS", r"figs?"),
    ("GARLIC", r"garlic"),
    ("GINGER", r"ginger"),
    ("GAI CHOY", r"gai choy"),
    ("GAI LAN", r"gai lan"),
    ("GRAPEFRUIT", r"grapefruit|grapefurit|grapefuit"),
    ("GRAPES", r"grapes?"),
    ("HERBS", r"herbs?"),
    ("HORSERADISH", r"horseradish"),
    ("GUAVA", r"guava"),
    ("JICAMA", r"jicama"),
    ("KALE", r"kale"),
    ("KIWI", r"kiwi"),
    ("KUMQUAT", r"kumquat"),
    ("LEEKS", r"leeks?"),
    ("LEMON", r"lemons?"),
    ("LETTUCE", r"lettuce|lettuc"),
    ("LIMEQUATS", r"limequats?"),
    ("LIME", r"limes?"),
    ("LYCHEES", r"lychees?"),
    ("MANDARIN", r"mandarins?"),
    ("ORANGE", r"oranges?"),
    ("MANGO", r"mango(?:es|s)?"),
    ("MELON", r"melons?"),
    ("MINT", r"mint"),
    ("MUSHROOM", r"mushrooms?"),
    ("MUSTARD GREENS", r"mustard greens"),
    ("NECTARINE", r"nectarines?"),
    ("OKRA", r"okra"),
    ("ON CHOY", r"on choy"),
    ("ONION", r"onions?"),
    ("PAPAYA", r"papaya"),
    ("PARSLEY", r"parsley"),
    ("PARSNIP", r"parsnip"),
    ("PASSION FRUIT", r"passion fruit"),
    ("PEACH", r"peaches|peach"),
    ("PEAR", r"pears?"),
    ("PEAS", r"peas"),
    ("PERSIMMON", r"persimmons?"),
    ("PINEAPPLE", r"pineapple"),
    ("PLUM", r"plums?"),
    ("POMEGRANATE", r"pomegranate|pomegramate|pomegrante"),
    ("POMELO", r"pomelo|pummelo|pommelo|shaddock"),
    ("POTATO", r"potato(?:es)?"),
    ("PRUNES", r"prunes?"),
    ("PUMPKIN", r"pumpkin"),
    ("ORNAMENTAL GOURD", r"ornamental gourd"),
    ("RADICCHIO", r"radicchio"),
    ("RADISH", r"radish"),
    ("RAMBOUTAN", r"ramboutan"),
    ("RAPINI", r"rapini"),
    ("RHUBARB", r"rhubarb"),
    ("ROMANESCO", r"romanesco"),
    ("ROSEMARY", r"rosemary"),
    ("RUTABAGA", r"rutabaga"),
    ("SHER LI HON", r"sher li hon"),
    ("SING QUA", r"sing qua"),
    ("SPINACH", r"spinach"),
    ("STARFRUIT", r"starfruit"),
    ("TANGELO", r"tangelos?"),
    ("TANGERINE", r"tangerines?"),
    ("TOMATILLO", r"tomatillos?"),
    ("TOMATO", r"tomatoes|tomatoe?s?"),
    ("TURNIP", r"turnip"),
    ("WATER CHESTNUTS", r"water chestnuts"),
    ("WATERCRESS", r"watercress"),
    ("YAM", r"yams?"),
    ("YU CHOY", r"yu choy"),
    ("YUCCA", r"yucca"),
    ("ZUCCHINI", r"zucchini"),
)

MODIFIERS = re.compile(
    r"\b(?:organic|bc|local|bulk|case|pint|quart|flat|cut|small|medium|large|"
    r"sm|sml|lrg|xl|baby|mini|jumbo|bunch|half|by|the|pounds?|lbs?)\b|\b\d+\b",
    re.I,
)


def product_key(name):
    """Return a main-item key and variety key without changing the source name."""
    matches = []
    for product, pattern in PRODUCTS:
        match = re.search(r"\b(?:" + pattern + r")\b", name, re.I)
        if match:
            matches.append((match.start(), -len(match.group()), product, pattern, match))
    specific = [match for match in matches
                if match[2] in {"WATERMELON", "CANTALOUPE", "HONEYDEW", "GRANADILLA"}]
    if specific:
        _, _, product, pattern, match = min(specific)
    elif matches:
        _, _, product, pattern, match = min(matches)
    else:
        remainder = name
        words = re.findall(r"[^\W_]+", MODIFIERS.sub(" ", name), re.UNICODE)
        if not words:
            return "", "", False
        product = words[0].upper()
        remainder = re.sub(r"\b" + re.escape(words[0]) + r"\b", " ", remainder,
                           count=1, flags=re.I)
        return product, " ".join(re.findall(r"[^\W_]+", MODIFIERS.sub(" ", remainder).upper(),
                                            re.UNICODE)), False
    remainder = name[:match.start()] + " " + name[match.end():]
    if product in {"WATERMELON", "CANTALOUPE", "HONEYDEW"}:
        remainder = re.sub(r"\bmelons?\b", " ", remainder, flags=re.I)
    if product == "SQUASH":
        remainder = re.sub(r"\b(?:squash|kabocha)\b", " ", remainder, flags=re.I)
        if re.search(r"\bkabocha\b", name, re.I):
            remainder += " KABOCHA"
    variety = " ".join(re.findall(r"[^\W_]+", MODIFIERS.sub(" ", remainder).upper(), re.UNICODE))
    return product, variety, True


def order_produce(rows):
    """Sort display copies; report unrecognized main items for review."""
    result = []
    for row in rows:
        item = dict(row)
        product, variety, recognized = product_key(item["name"])
        if not product or not product[0].isalpha():
            raise ValueError(f"{item['source_id']}: cannot determine produce letter")
        item.update(display_section=product[0], display_name=item["name"].upper(),
                    sort_product=product, sort_variety=variety,
                    sort_status="RECOGNIZED" if recognized else "MAIN_ITEM_REVIEW_REQUIRED")
        result.append(item)
    by_id = {row["source_id"]: row for row in result}
    for row in result:
        counterparts = row.get("counterpart_source_ids", "").split(";")
        counterpart = by_id.get(counterparts[0])
        if row["organic"] and counterpart:
            row["sort_product"] = counterpart["sort_product"]
            row["sort_variety"] = counterpart["sort_variety"]
            row["display_section"] = counterpart["display_section"]
    return sorted(result, key=lambda row: (
        row["sort_product"], row["sort_variety"], row["organic"],
        bool(MODIFIERS.search(row["name"])),
        row["name"].upper(), row["plu"], row["source_id"]))


def display_unique_rows(rows):
    """Suppress only exact display duplicates while retaining their identities."""
    displayed, duplicates, seen = [], [], {}
    for row in rows:
        key = (row["plu"], row["name"])
        if key in seen:
            duplicates.append(dict(
                source_id=row["source_id"], source_ids=";".join(row["source_ids"]),
                plu=row["plu"], name=row["name"],
                displayed_source_id=seen[key]["source_id"],
                displayed_source_ids=";".join(seen[key]["source_ids"])))
        else:
            seen[key] = row
            displayed.append(row)
    return displayed, duplicates
