"""
Built-in calendar events for the SCB library (holidays, fandom days, releases, industry events, Carlin
milestones). Used by scb_library.py (Excel) and build_site.py (website).

Every event is a dict:
    key     stable id (so a user's "Show = No" choice survives rebuilds)
    cat     one of CATEGORIES
    title   text; "{n}" is replaced by the number of years since the start year (birthdays, anniversaries)
    start   datetime.date            end   datetime.date or None (multi-day)
    repeat  True = happens every year on the same month/day (from start's year onward)
    notes   free text / link         kw    keywords used to find related videos (movies, trailers)

Release dates are US dates. They were assembled from public sources and the SCB upload record;
please correct anything that looks off (add your own row and hide the built-in one).
"""
import csv
from datetime import date, timedelta

CATEGORIES = [  # name, emoji, colour (hex without #)
    ("Holiday", "🎉", "D1495B"),
    ("Fandom", "🌟", "8E5BB5"),
    ("Movie / trailer", "🎬", "3B6FB6"),
    ("Game release", "🎮", "2E9E6B"),
    ("Carlin life", "💛", "E08A1E"),
    ("SCB milestone", "🏆", "C9A227"),
    ("Other", "📌", "5C6B7A"),
]
CAT_NAMES = [c[0] for c in CATEGORIES]
EMOJI = {c[0]: c[1] for c in CATEGORIES}


def d(s):
    y, m, dd = (int(x) for x in s.split("-"))
    return date(y, m, dd)


# ----------------------------------------------------------------------------- date rules
def nth_weekday(year, month, weekday, n):
    """weekday: Mon=0..Sun=6; n=1..4, or -1 for the last one."""
    if n > 0:
        first = date(year, month, 1)
        return first + timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))
    last = date(year + (month == 12), month % 12 + 1, 1) - timedelta(days=1)
    return last - timedelta(days=(last.weekday() - weekday) % 7)


def easter(year):
    a, b, c = year % 19, year // 100, year % 100
    dd, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - dd - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = (h + l - 7 * m + 114) % 31 + 1
    return date(year, month, day)


# ----------------------------------------------------------------------------- holidays
FIXED_HOLIDAYS = [  # month, day, title, first year
    (1, 1, "New Year's Day", 2012), (2, 14, "Valentine's Day", 2012), (3, 17, "St. Patrick's Day", 2012),
    (4, 1, "April Fools' Day", 2012), (6, 19, "Juneteenth", 2021), (7, 4, "Independence Day (4th of July)", 2012),
    (10, 31, "Halloween", 2012), (11, 11, "Veterans Day", 2012), (12, 24, "Christmas Eve", 2012),
    (12, 25, "Christmas Day", 2012), (12, 31, "New Year's Eve", 2012),
]
SUPER_BOWL = ["2012-02-05", "2013-02-03", "2014-02-02", "2015-02-01", "2016-02-07", "2017-02-05", "2018-02-04",
              "2019-02-03", "2020-02-02", "2021-02-07", "2022-02-13", "2023-02-12", "2024-02-11", "2025-02-09",
              "2026-02-08", "2027-02-14"]
OSCARS = ["2012-02-26", "2013-02-24", "2014-03-02", "2015-02-22", "2016-02-28", "2017-02-26", "2018-03-04",
          "2019-02-24", "2020-02-09", "2021-04-25", "2022-03-27", "2023-03-12", "2024-03-10", "2025-03-02",
          "2026-03-15"]


def holidays(y0, y1):
    out = []
    for m, day, title, fy in FIXED_HOLIDAYS:
        out.append(dict(key=f"hol:{m:02d}{day:02d}", cat="Holiday", title=title, start=date(max(fy, y0), m, day),
                        end=None, repeat=True, notes="", kw=[]))
    rules = [("mlk", "Martin Luther King Jr. Day", lambda y: nth_weekday(y, 1, 0, 3)),
             ("pres", "Presidents' Day", lambda y: nth_weekday(y, 2, 0, 3)),
             ("easter", "Easter Sunday", easter),
             ("mothers", "Mother's Day", lambda y: nth_weekday(y, 5, 6, 2)),
             ("memorial", "Memorial Day", lambda y: nth_weekday(y, 5, 0, -1)),
             ("fathers", "Father's Day", lambda y: nth_weekday(y, 6, 6, 3)),
             ("labor", "Labor Day", lambda y: nth_weekday(y, 9, 0, 1)),
             ("thanks", "Thanksgiving", lambda y: nth_weekday(y, 11, 3, 4)),
             ("blackfri", "Black Friday", lambda y: nth_weekday(y, 11, 3, 4) + timedelta(days=1))]
    for y in range(y0, y1 + 1):
        for k, title, fn in rules:
            out.append(dict(key=f"hol:{k}:{y}", cat="Holiday", title=title, start=fn(y), end=None, repeat=False, notes="", kw=[]))
    for s in SUPER_BOWL:
        if y0 <= d(s).year <= y1:
            out.append(dict(key=f"hol:sb:{s[:4]}", cat="Holiday", title="Super Bowl Sunday", start=d(s), end=None, repeat=False, notes="", kw=[]))
    for s in OSCARS:
        if y0 <= d(s).year <= y1:
            out.append(dict(key=f"hol:oscars:{s[:4]}", cat="Holiday", title="The Oscars", start=d(s), end=None, repeat=False,
                            notes="Academy Awards ceremony", kw=["oscar", "academy award"]))
    return out


# ----------------------------------------------------------------------------- fandom days (repeat yearly)
FANDOM = [  # month, day, title, first year, notes
    (2, 27, "Pokémon Day", 2012, "Anniversary of the first Pokémon games (1996)"),
    (3, 10, "Mario Day (MAR10)", 2012, ""),
    (5, 2, "Battle of Hogwarts anniversary", 2012, ""),
    (5, 4, "Star Wars Day (May the 4th)", 2012, ""),
    (5, 25, "Star Wars anniversary (A New Hope, 1977)", 2012, "Original Star Wars released May 25, 1977"),
    (6, 26, "Harry Potter book anniversary (1997)", 2012, "Philosopher's Stone first published June 26, 1997"),
    (7, 31, "Harry Potter's birthday (and J.K. Rowling's)", 2012, ""),
    (9, 1, "Back to Hogwarts Day", 2012, "The Hogwarts Express leaves King's Cross"),
    (9, 19, "Hermione Granger's birthday", 2012, "Fan-canon date"),
    (11, 18, "Mickey Mouse's birthday (Steamboat Willie, 1928)", 2012, ""),
    (12, 5, "Walt Disney's birthday (1901)", 2012, ""),
]
MEMORIAM = [
    ("2016-01-14", "Alan Rickman passes away (Professor Snape)", ["rickman", "snape"]),
    ("2016-12-27", "Carrie Fisher passes away (Princess Leia)", ["carrie fisher", "leia"]),
    ("2018-11-12", "Stan Lee passes away", ["stan lee"]),
    ("2020-08-28", "Chadwick Boseman passes away (Black Panther)", ["boseman", "black panther"]),
    ("2021-04-16", "Helen McCrory passes away (Narcissa Malfoy)", ["mccrory", "narcissa"]),
    ("2022-10-14", "Robbie Coltrane passes away (Hagrid)", ["coltrane", "hagrid"]),
    ("2023-09-27", "Michael Gambon passes away (Dumbledore)", ["gambon", "dumbledore"]),
    ("2024-09-27", "Maggie Smith passes away (Professor McGonagall)", ["maggie smith", "mcgonagall"]),
]


def fandom(y0, y1):
    out = [dict(key=f"fan:{m:02d}{day:02d}", cat="Fandom", title=t, start=date(max(fy, y0), m, day), end=None, repeat=True,
                notes=n, kw=[]) for m, day, t, fy, n in FANDOM]
    for s, t, kw in MEMORIAM:
        out.append(dict(key=f"fan:mem:{s}", cat="Fandom", title=t, start=d(s), end=None, repeat=False, notes="In memoriam", kw=kw))
    return out


# ----------------------------------------------------------------------------- conventions & industry
COMIC_CON = [("2012-07-12", "2012-07-15"), ("2013-07-18", "2013-07-21"), ("2014-07-24", "2014-07-27"), ("2015-07-09", "2015-07-12"),
             ("2016-07-21", "2016-07-24"), ("2017-07-20", "2017-07-23"), ("2018-07-19", "2018-07-22"), ("2019-07-18", "2019-07-21"),
             ("2021-07-23", "2021-07-25"), ("2022-07-21", "2022-07-24"), ("2023-07-20", "2023-07-23"), ("2024-07-25", "2024-07-28"),
             ("2025-07-24", "2025-07-27"), ("2026-07-23", "2026-07-26")]
D23 = [("2013-08-09", "2013-08-11"), ("2015-08-14", "2015-08-16"), ("2017-07-14", "2017-07-16"), ("2019-08-23", "2019-08-25"),
       ("2022-09-09", "2022-09-11"), ("2024-08-09", "2024-08-11"), ("2026-08-14", "2026-08-16")]
SW_CELEBRATION = [("2015-04-16", "2015-04-19", "Anaheim"), ("2016-07-15", "2016-07-17", "London"), ("2017-04-13", "2017-04-16", "Orlando"),
                  ("2019-04-11", "2019-04-14", "Chicago"), ("2023-04-07", "2023-04-11", "London"), ("2025-04-18", "2025-04-21", "Japan")]
INDUSTRY = [
    ("2012-10-30", None, "Disney announces it is buying Lucasfilm", "Star Wars sequels announced", ["lucasfilm"]),
    ("2019-03-20", None, "Disney completes its acquisition of 21st Century Fox", "", ["fox"]),
    ("2019-11-12", None, "Disney+ launches", "", ["disney+", "disney plus"]),
    ("2020-03-11", None, "WHO declares COVID-19 a pandemic", "Lockdowns follow within days", ["covid", "quarantine"]),
    ("2023-05-02", "2023-09-27", "Writers Guild (WGA) strike", "Hollywood production and promotion slowed", []),
    ("2023-07-14", "2023-11-09", "SAG-AFTRA actors' strike", "Actors could not promote films", []),
]


def conventions():
    out = []
    for a, b in COMIC_CON:
        out.append(dict(key=f"con:sdcc:{a[:4]}", cat="Movie / trailer", title="San Diego Comic-Con", start=d(a), end=d(b), repeat=False,
                        notes="Big trailer and announcement reveals (Marvel, Star Wars, Disney, Harry Potter)", kw=["comic-con", "comic con", "sdcc"]))
    for a, b in D23:
        out.append(dict(key=f"con:d23:{a[:4]}", cat="Movie / trailer", title="D23 Expo (Disney fan event)", start=d(a), end=d(b), repeat=False,
                        notes="Disney, Pixar, Marvel and Lucasfilm announcements", kw=["d23"]))
    for a, b, city in SW_CELEBRATION:
        out.append(dict(key=f"con:swc:{a[:4]}", cat="Movie / trailer", title=f"Star Wars Celebration ({city})", start=d(a), end=d(b), repeat=False,
                        notes="Lucasfilm fan convention - trailers and reveals", kw=["celebration"]))
    for a, b, t, n, kw in INDUSTRY:
        cat = "Other" if "strike" in t or "COVID" in t else "Movie / trailer"
        out.append(dict(key=f"ind:{a}", cat=cat, title=t, start=d(a), end=d(b) if b else None, repeat=False, notes=n, kw=kw))
    return out


# ----------------------------------------------------------------------------- releases (US)
# (date, title, franchise, keywords, kind)   kind: "film" | "series" | "trailer"
MOVIES = [
    # --- Marvel
    ("2012-05-04", "The Avengers", "Marvel", ["avengers"], "film"),
    ("2013-05-03", "Iron Man 3", "Marvel", ["iron man"], "film"),
    ("2013-11-08", "Thor: The Dark World", "Marvel", ["thor"], "film"),
    ("2014-04-04", "Captain America: The Winter Soldier", "Marvel", ["winter soldier", "captain america"], "film"),
    ("2014-08-01", "Guardians of the Galaxy", "Marvel", ["guardians"], "film"),
    ("2015-05-01", "Avengers: Age of Ultron", "Marvel", ["age of ultron", "avengers"], "film"),
    ("2015-07-17", "Ant-Man", "Marvel", ["ant-man", "ant man"], "film"),
    ("2016-05-06", "Captain America: Civil War", "Marvel", ["civil war", "captain america"], "film"),
    ("2016-11-04", "Doctor Strange", "Marvel", ["doctor strange"], "film"),
    ("2017-05-05", "Guardians of the Galaxy Vol. 2", "Marvel", ["guardians"], "film"),
    ("2017-07-07", "Spider-Man: Homecoming", "Marvel", ["spider-man", "spiderman", "spider man"], "film"),
    ("2017-11-03", "Thor: Ragnarok", "Marvel", ["ragnarok", "thor"], "film"),
    ("2018-02-16", "Black Panther", "Marvel", ["black panther"], "film"),
    ("2018-04-27", "Avengers: Infinity War", "Marvel", ["infinity war"], "film"),
    ("2018-07-06", "Ant-Man and the Wasp", "Marvel", ["ant-man", "wasp"], "film"),
    ("2019-03-08", "Captain Marvel", "Marvel", ["captain marvel"], "film"),
    ("2019-04-26", "Avengers: Endgame", "Marvel", ["endgame"], "film"),
    ("2019-07-02", "Spider-Man: Far From Home", "Marvel", ["far from home", "spider-man"], "film"),
    ("2021-01-15", "WandaVision (Disney+)", "Marvel", ["wandavision"], "series"),
    ("2021-03-19", "The Falcon and the Winter Soldier (Disney+)", "Marvel", ["falcon"], "series"),
    ("2021-06-09", "Loki (Disney+)", "Marvel", ["loki"], "series"),
    ("2021-07-09", "Black Widow", "Marvel", ["black widow"], "film"),
    ("2021-08-11", "What If...? (Disney+)", "Marvel", ["what if"], "series"),
    ("2021-09-03", "Shang-Chi and the Legend of the Ten Rings", "Marvel", ["shang-chi", "shang chi"], "film"),
    ("2021-11-05", "Eternals", "Marvel", ["eternals"], "film"),
    ("2021-11-24", "Hawkeye (Disney+)", "Marvel", ["hawkeye"], "series"),
    ("2021-12-17", "Spider-Man: No Way Home", "Marvel", ["no way home", "spider-man"], "film"),
    ("2022-05-06", "Doctor Strange in the Multiverse of Madness", "Marvel", ["multiverse of madness", "doctor strange"], "film"),
    ("2022-07-08", "Thor: Love and Thunder", "Marvel", ["love and thunder", "thor"], "film"),
    ("2022-11-11", "Black Panther: Wakanda Forever", "Marvel", ["wakanda", "black panther"], "film"),
    ("2023-02-17", "Ant-Man and the Wasp: Quantumania", "Marvel", ["quantumania", "ant-man"], "film"),
    ("2023-05-05", "Guardians of the Galaxy Vol. 3", "Marvel", ["guardians"], "film"),
    ("2023-10-06", "Loki season 2 (Disney+)", "Marvel", ["loki"], "series"),
    ("2023-11-10", "The Marvels", "Marvel", ["the marvels"], "film"),
    ("2024-07-26", "Deadpool & Wolverine", "Marvel", ["deadpool"], "film"),
    ("2025-02-14", "Captain America: Brave New World", "Marvel", ["brave new world", "captain america"], "film"),
    ("2025-03-04", "Daredevil: Born Again (Disney+)", "Marvel", ["daredevil"], "series"),
    ("2025-05-02", "Thunderbolts*", "Marvel", ["thunderbolts"], "film"),
    ("2025-07-25", "The Fantastic Four: First Steps", "Marvel", ["fantastic four"], "film"),
    ("2026-07-31", "Spider-Man: Brand New Day", "Marvel", ["brand new day", "spider-man"], "film"),
    ("2026-10-14", "VisionQuest (Disney+)", "Marvel", ["visionquest", "vision"], "series"),
    ("2026-12-18", "Avengers: Doomsday", "Marvel", ["doomsday", "avengers"], "film"),
    ("2027-12-17", "Avengers: Secret Wars", "Marvel", ["secret wars", "avengers"], "film"),
    # --- Pixar
    ("2012-06-22", "Brave", "Pixar", ["brave"], "film"),
    ("2013-06-21", "Monsters University", "Pixar", ["monsters university", "monsters"], "film"),
    ("2015-06-19", "Inside Out", "Pixar", ["inside out"], "film"),
    ("2015-11-25", "The Good Dinosaur", "Pixar", ["good dinosaur"], "film"),
    ("2016-06-17", "Finding Dory", "Pixar", ["dory"], "film"),
    ("2017-06-16", "Cars 3", "Pixar", ["cars"], "film"),
    ("2017-11-22", "Coco", "Pixar", ["coco"], "film"),
    ("2018-06-15", "Incredibles 2", "Pixar", ["incredibles"], "film"),
    ("2019-06-21", "Toy Story 4", "Pixar", ["toy story"], "film"),
    ("2020-03-06", "Onward", "Pixar", ["onward"], "film"),
    ("2020-12-25", "Soul (Disney+)", "Pixar", ["soul"], "film"),
    ("2021-06-18", "Luca (Disney+)", "Pixar", ["luca"], "film"),
    ("2022-03-11", "Turning Red (Disney+)", "Pixar", ["turning red"], "film"),
    ("2022-06-17", "Lightyear", "Pixar", ["lightyear"], "film"),
    ("2023-06-16", "Elemental", "Pixar", ["elemental"], "film"),
    ("2024-06-14", "Inside Out 2", "Pixar", ["inside out"], "film"),
    ("2025-06-20", "Elio", "Pixar", ["elio"], "film"),
    ("2026-03-06", "Hoppers", "Pixar", ["hoppers"], "film"),
    ("2026-06-19", "Toy Story 5", "Pixar", ["toy story"], "film"),
    # --- Disney
    ("2012-11-02", "Wreck-It Ralph", "Disney", ["wreck-it ralph", "ralph"], "film"),
    ("2013-11-27", "Frozen", "Disney", ["frozen"], "film"),
    ("2014-11-07", "Big Hero 6", "Disney", ["big hero"], "film"),
    ("2015-03-13", "Cinderella (live-action)", "Disney", ["cinderella"], "film"),
    ("2016-03-04", "Zootopia", "Disney", ["zootopia"], "film"),
    ("2016-04-15", "The Jungle Book (live-action)", "Disney", ["jungle book"], "film"),
    ("2016-11-23", "Moana", "Disney", ["moana"], "film"),
    ("2017-03-17", "Beauty and the Beast (live-action)", "Disney", ["beauty and the beast"], "film"),
    ("2018-11-21", "Ralph Breaks the Internet", "Disney", ["ralph"], "film"),
    ("2019-05-24", "Aladdin (live-action)", "Disney", ["aladdin"], "film"),
    ("2019-07-19", "The Lion King (live-action)", "Disney", ["lion king"], "film"),
    ("2019-11-22", "Frozen II", "Disney", ["frozen"], "film"),
    ("2020-09-04", "Mulan (live-action, Disney+)", "Disney", ["mulan"], "film"),
    ("2021-03-05", "Raya and the Last Dragon", "Disney", ["raya"], "film"),
    ("2021-05-28", "Cruella", "Disney", ["cruella"], "film"),
    ("2021-11-24", "Encanto", "Disney", ["encanto"], "film"),
    ("2022-11-23", "Strange World", "Disney", ["strange world"], "film"),
    ("2023-05-26", "The Little Mermaid (live-action)", "Disney", ["little mermaid"], "film"),
    ("2023-11-22", "Wish", "Disney", ["wish"], "film"),
    ("2024-11-27", "Moana 2", "Disney", ["moana"], "film"),
    ("2024-12-20", "Mufasa: The Lion King", "Disney", ["mufasa", "lion king"], "film"),
    ("2025-03-21", "Snow White (live-action)", "Disney", ["snow white"], "film"),
    ("2025-05-30", "Lilo & Stitch (live-action)", "Disney", ["lilo", "stitch"], "film"),
    ("2025-10-10", "TRON: Ares", "Disney", ["tron"], "film"),
    ("2025-11-26", "Zootopia 2", "Disney", ["zootopia"], "film"),
    ("2026-07-10", "Moana (live-action)", "Disney", ["moana"], "film"),
    ("2026-11-25", "Hexed", "Disney", ["hexed"], "film"),
    ("2027-11-24", "Frozen 3 (November 2027 - date to be confirmed)", "Disney", ["frozen"], "film"),
    # --- Star Wars
    ("2015-12-18", "Star Wars: The Force Awakens", "Star Wars", ["force awakens", "star wars"], "film"),
    ("2016-12-16", "Rogue One: A Star Wars Story", "Star Wars", ["rogue one"], "film"),
    ("2017-12-15", "Star Wars: The Last Jedi", "Star Wars", ["last jedi", "star wars"], "film"),
    ("2018-05-25", "Solo: A Star Wars Story", "Star Wars", ["solo"], "film"),
    ("2019-12-20", "Star Wars: The Rise of Skywalker", "Star Wars", ["rise of skywalker", "skywalker", "star wars"], "film"),
    ("2019-11-12", "The Mandalorian season 1 (Disney+)", "Star Wars", ["mandalorian", "baby yoda"], "series"),
    ("2020-10-30", "The Mandalorian season 2 (Disney+)", "Star Wars", ["mandalorian"], "series"),
    ("2021-12-29", "The Book of Boba Fett (Disney+)", "Star Wars", ["boba fett"], "series"),
    ("2022-05-27", "Obi-Wan Kenobi (Disney+)", "Star Wars", ["obi-wan", "kenobi"], "series"),
    ("2022-09-21", "Andor season 1 (Disney+)", "Star Wars", ["andor"], "series"),
    ("2023-03-01", "The Mandalorian season 3 (Disney+)", "Star Wars", ["mandalorian"], "series"),
    ("2023-08-23", "Ahsoka (Disney+)", "Star Wars", ["ahsoka"], "series"),
    ("2024-06-04", "The Acolyte (Disney+)", "Star Wars", ["acolyte"], "series"),
    ("2024-12-03", "Skeleton Crew (Disney+)", "Star Wars", ["skeleton crew"], "series"),
    ("2025-04-22", "Andor season 2 (Disney+)", "Star Wars", ["andor"], "series"),
    ("2026-05-22", "The Mandalorian & Grogu", "Star Wars", ["mandalorian", "grogu"], "film"),
    ("2027-01-20", "Ahsoka season 2 (Disney+)", "Star Wars", ["ahsoka"], "series"),
    ("2027-05-28", "Star Wars: Starfighter", "Star Wars", ["starfighter", "star wars"], "film"),
    # --- Harry Potter / Wizarding World
    ("2016-07-31", "Harry Potter and the Cursed Child - script book released", "Harry Potter", ["cursed child"], "film"),
    ("2016-11-18", "Fantastic Beasts and Where to Find Them", "Harry Potter", ["fantastic beasts"], "film"),
    ("2018-11-16", "Fantastic Beasts: The Crimes of Grindelwald", "Harry Potter", ["grindelwald", "fantastic beasts"], "film"),
    ("2022-04-15", "Fantastic Beasts: The Secrets of Dumbledore", "Harry Potter", ["secrets of dumbledore", "fantastic beasts"], "film"),
    ("2023-02-10", "Hogwarts Legacy (game) released", "Harry Potter", ["hogwarts legacy"], "film"),
    ("2026-09-02", "HBO Harry Potter series - second trailer", "Harry Potter", ["hbo", "harry potter series"], "trailer"),
    ("2026-12-25", "HBO's Harry Potter series premieres (HBO Max)", "Harry Potter", ["hbo", "harry potter series"], "series"),
    # --- other franchises they cover
    ("2012-03-23", "The Hunger Games", "Hunger Games", ["hunger games"], "film"),
    ("2013-11-22", "The Hunger Games: Catching Fire", "Hunger Games", ["catching fire", "hunger games"], "film"),
    ("2014-11-21", "The Hunger Games: Mockingjay - Part 1", "Hunger Games", ["mockingjay", "hunger games"], "film"),
    ("2015-11-20", "The Hunger Games: Mockingjay - Part 2", "Hunger Games", ["mockingjay", "hunger games"], "film"),
    ("2023-11-17", "The Hunger Games: The Ballad of Songbirds & Snakes", "Hunger Games", ["songbirds", "hunger games"], "film"),
    ("2024-11-22", "Wicked", "Wicked", ["wicked"], "film"),
    ("2025-11-21", "Wicked: For Good", "Wicked", ["wicked"], "film"),
    # --- trailers / reveals
    ("2014-11-28", "Star Wars: The Force Awakens - first teaser trailer", "Star Wars", ["force awakens", "star wars"], "trailer"),
    ("2015-04-16", "The Force Awakens trailer #2 (Star Wars Celebration)", "Star Wars", ["force awakens", "star wars"], "trailer"),
    ("2015-10-19", "The Force Awakens - final trailer", "Star Wars", ["force awakens", "star wars"], "trailer"),
    ("2016-04-07", "Rogue One - first teaser trailer", "Star Wars", ["rogue one"], "trailer"),
    ("2017-04-14", "The Last Jedi - first teaser trailer", "Star Wars", ["last jedi", "star wars"], "trailer"),
    ("2019-04-12", "The Rise of Skywalker - first teaser trailer", "Star Wars", ["skywalker", "star wars"], "trailer"),
    ("2019-08-23", "The Mandalorian - first trailer (D23)", "Star Wars", ["mandalorian"], "trailer"),
    ("2017-11-29", "Avengers: Infinity War - first trailer", "Marvel", ["infinity war"], "trailer"),
    ("2018-12-07", "Avengers: Endgame - first trailer", "Marvel", ["endgame"], "trailer"),
    ("2019-03-14", "Avengers: Endgame - second trailer", "Marvel", ["endgame"], "trailer"),
    ("2021-08-24", "Spider-Man: No Way Home - first trailer", "Marvel", ["no way home", "spider-man"], "trailer"),
    ("2018-11-15", "Toy Story 4 - first trailer", "Pixar", ["toy story"], "trailer"),
    ("2018-11-20", "The Lion King (2019) - first teaser", "Disney", ["lion king"], "trailer"),
    ("2019-02-12", "Frozen II - first teaser trailer", "Disney", ["frozen"], "trailer"),
]
GAMES = [
    ("2013-10-12", "Pokémon X and Y released", ["pokemon x", "pokémon x", "pokemon y"]),
    ("2014-05-30", "Mario Kart 8 released (Wii U)", ["mario kart"]),
    ("2014-09-02", "The Sims 4 released", ["sims"]),
    ("2014-11-21", "Pokémon Omega Ruby / Alpha Sapphire released", ["omega ruby", "alpha sapphire"]),
    ("2014-11-21", "Super Smash Bros. for Wii U released", ["smash"]),
    ("2015-07-07", "Rocket League released", ["rocket league"]),
    ("2016-07-06", "Pokémon GO released (US)", ["pokemon go", "pokémon go"]),
    ("2016-11-18", "Pokémon Sun and Moon released", ["sun and moon", "pokemon sun", "pokémon sun"]),
    ("2017-03-03", "Nintendo Switch and Zelda: Breath of the Wild released", ["switch", "breath of the wild", "zelda"]),
    ("2017-04-28", "Mario Kart 8 Deluxe released", ["mario kart"]),
    ("2017-12-06", "Getting Over It with Bennett Foddy released", ["getting over it"]),
    ("2018-11-16", "Pokémon Let's Go, Pikachu! / Eevee! released", ["let's go"]),
    ("2018-12-07", "Super Smash Bros. Ultimate released", ["smash"]),
    ("2019-11-15", "Pokémon Sword and Shield released", ["sword and shield", "pokemon sword", "pokémon sword"]),
    ("2020-09-23", "Rocket League goes free-to-play", ["rocket league"]),
    ("2021-11-19", "Pokémon Brilliant Diamond / Shining Pearl released", ["brilliant diamond", "shining pearl"]),
    ("2022-01-28", "Pokémon Legends: Arceus released", ["arceus"]),
    ("2022-11-18", "Pokémon Scarlet and Violet released", ["scarlet", "violet"]),
    ("2023-04-05", "The Super Mario Bros. Movie released", ["mario movie", "super mario bros"]),
    ("2023-05-12", "The Legend of Zelda: Tears of the Kingdom released", ["tears of the kingdom", "zelda"]),
    ("2025-06-05", "Nintendo Switch 2 and Mario Kart World released", ["switch 2", "mario kart"]),
    ("2025-10-16", "Pokémon Legends: Z-A released", ["legends z-a", "z-a"]),
]
# Carlin milestones that are public in the channels' own uploads (date = upload date of the announcing video)
CARLIN_LIFE = [
    ("2014-12-10", "Carlin life", "J & Beth announce their engagement", "Announced in the video 'WE GOT ENGAGED!!'", ["engaged"]),
    ("2017-12-18", "Carlin life", "J & Beth welcome their first child (announced)", "Announced in the video 'MY SON IS BORN'", []),
    ("2019-03-11", "Carlin life", "Ben & Alyce announce their engagement", "Announced in the video 'BEN & ALYCE GET ENGAGED!!'", ["engaged"]),
    ("2020-02-20", "Carlin life", "Ben & Alyce's Disney wedding (video)", "Video: 'OUR FAIRY TALE DISNEY WEDDING'", ["wedding"]),
    ("2020-03-09", "Carlin life", "J & Beth welcome twins (announced)", "Announced in the video 'The Twins Are Born'", []),
    ("2021-11-12", "Carlin life", "Ben & Alyce welcome their first child; 'With the Carlins' channel begins", "Video: 'Welcome to Life with the Carlins'", []),
]
MILESTONES = [
    ("2015-06-17", "SCB reaches 100,000 subscribers", "Video: 'The Utah Teapot and 100,000 Subscribers!'"),
    ("2019-03-02", "SCB Meet Ups announced", ""),
    ("2019-07-23", "SCB reaches 2 million subscribers", "Video: 'A Week Of Milestones | 2-Million Subscribers!'"),
    ("2019-09-20", "Carlin Brothers Coffee 1-year anniversary stream", ""),
    ("2022-06-01", "J & Ben's first live show", "Video: 'Our First LIVE SHOW'"),
    ("2022-06-25", "SCB 10-year celebration livestream", ""),
    ("2024-03-12", "'We're going on tour!' - Through the Griffin Door tour announced", ""),
    ("2024-06-18", "Through the Griffin Tour - West Coast announced", ""),
    ("2024-10-22", "Through the Griffin Tour - Southeast announced", ""),
    ("2025-04-23", "Through the Griffin Tour - Midwest announced", ""),
    ("2025-11-05", "Through the Griffin Tour - Southwest announced", ""),
    ("2026-04-08", "Through the Griffin Tour - Northeast 2026 announced", ""),
]
BIRTHDAYS = [  # month/day/year from public bio sites - flagged for verification
    (1988, 2, 4, "J's birthday (turns {n})"),
    (1989, 10, 25, "Ben's birthday (turns {n})"),
]


def _label(title, kind):
    low = title.lower()
    if kind == "trailer" or any(w in low for w in ("released", "premieres", "trailer", "to be confirmed")):
        return title
    return f"{title} premieres" if kind == "series" else f"{title} released"


def releases():
    out = []
    for s, title, franchise, kw, kind in MOVIES:
        out.append(dict(key=f"rel:{s}:{title[:24]}", cat="Movie / trailer", title=_label(title, kind), start=d(s), end=None, repeat=False,
                        notes=f"{franchise} - US date" + (" (trailer)" if kind == "trailer" else ""), kw=kw))
    for s, title, kw in GAMES:
        out.append(dict(key=f"game:{s}:{title[:24]}", cat="Game release", title=title, start=d(s), end=None, repeat=False, notes="US date", kw=kw))
    return out


def carlin():
    out = []
    for s, cat, title, notes, kw in CARLIN_LIFE:
        out.append(dict(key=f"life:{s}", cat=cat, title=title, start=d(s), end=None, repeat=False, notes=notes, kw=kw))
    for y, m, day, title in BIRTHDAYS:
        out.append(dict(key=f"life:bday:{m:02d}{day:02d}", cat="Carlin life", title=title, start=date(y, m, day), end=None, repeat=True,
                        notes="Birth date from public bio sites - please verify", kw=["birthday"]))
    for s, title, notes in MILESTONES:
        out.append(dict(key=f"ms:{s}", cat="SCB milestone", title=title, start=d(s), end=None, repeat=False, notes=notes, kw=[]))
    return out


def milestones_from_data(videos, channels):
    """Auto-generated: each channel's first video, its 100th/500th/1,000th..., and its most-viewed video."""
    out = []
    by = {}
    for v in videos:
        by.setdefault(v["channel_id"], []).append(v)
    for c in channels:
        a = sorted(by.get(c["id"], []), key=lambda v: v["published_utc"])
        if not a:
            continue
        name = c["name"]
        out.append(dict(key=f"auto:first:{c['id']}", cat="SCB milestone", title=f"{name}: first video in the library", start=a[0]["day"].date(),
                        end=None, repeat=False, notes=a[0]["title"], kw=[]))
        for n in (100, 250, 500, 1000, 1500, 2000):
            if len(a) >= n:
                out.append(dict(key=f"auto:n{n}:{c['id']}", cat="SCB milestone", title=f"{name}: video #{n:,}", start=a[n - 1]["day"].date(),
                                end=None, repeat=False, notes=a[n - 1]["title"], kw=[]))
        top = max(a, key=lambda v: v.get("views") or 0)
        if (top.get("views") or 0) > 0:
            out.append(dict(key=f"auto:top:{c['id']}", cat="SCB milestone", title=f"{name}: most-viewed video uploaded", start=top["day"].date(),
                            end=None, repeat=False, notes=top["title"], kw=[]))
    return out


def all_events(y0, y1, videos=None, channels=None, extra=()):
    """Every built-in event. y0..y1 = years the variable-date holidays are expanded for."""
    ev = holidays(y0, y1) + fandom(y0, y1) + conventions() + releases() + carlin()
    if videos and channels:
        ev += milestones_from_data(videos, channels)
    ev += list(extra)
    return ev


def load_csv(path):
    """Read a site owner's extra events (my_events.csv). Columns: date,end,title,category,repeat,notes"""
    out = []
    try:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for i, r in enumerate(csv.DictReader(fh)):
                if not (r.get("date") and r.get("title")):
                    continue
                cat = r.get("category") or "Other"
                out.append(dict(key=f"csv:{i}:{r['date']}", cat=cat if cat in CAT_NAMES else "Other", title=r["title"].strip(),
                                start=d(r["date"].strip()), end=d(r["end"].strip()) if (r.get("end") or "").strip() else None,
                                repeat=(r.get("repeat") or "").strip().lower() in ("yes", "y", "true", "1"),
                                notes=(r.get("notes") or "").strip(), kw=[]))
    except FileNotFoundError:
        pass
    return out


def occurrences(ev, y_max, y_min=None):
    """Expand one event into concrete (start, end) date pairs up to year y_max (for the website)."""
    s, e = ev["start"], ev.get("end") or ev["start"]
    span = (e - s).days
    if not ev["repeat"]:
        return [(s, e, None)]
    out = []
    for y in range(max(s.year, y_min or s.year), y_max + 1):
        try:
            st = date(y, s.month, s.day)
        except ValueError:      # Feb 29
            st = date(y, 3, 1)
        out.append((st, st + timedelta(days=span), y - s.year))
    return out
