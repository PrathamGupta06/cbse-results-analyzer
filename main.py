CBSE_CLASS_12_SUBJECT_CODES = {
    # Language Subjects
    "001": "ENGLISH ELECTIVE",
    "301": "ENGLISH CORE",
    "002": "HINDI ELECTIVE",
    "302": "HINDI CORE",
    "003": "URDU ELECTIVE",
    "303": "URDU CORE",
    "022": "SANSKRIT ELECTIVE",
    "322": "SANSKRIT CORE",
    "104": "PUNJABI",
    "105": "BENGALI",
    "106": "TAMIL",
    "107": "TELUGU",
    "108": "SINDHI",
    "109": "MARATHI",
    "110": "GUJARATI",
    "111": "MANIPURI",
    "112": "MALAYALAM",
    "113": "ODIA",
    "114": "ASSAMESE",
    "115": "KANNADA",
    "116": "ARABIC",
    "117": "TIBETAN",
    "118": "FRENCH",
    "120": "GERMAN",
    "121": "RUSSIAN",
    "123": "PERSIAN",
    "124": "NEPALI",
    "125": "LIMBOO",
    "126": "LEPCHA",
    "189": "TELUGU TELANGANA",
    "192": "BODO",
    "193": "TANGKHUL",
    "194": "JAPANESE",
    "195": "BHUTIA",
    "196": "SPANISH",
    "197": "KASHMIRI",
    "198": "MIZO",
    "199": "BAHASA MELAYU",

    # Academic Subjects
    "027": "HISTORY",
    "028": "POLITICAL SCIENCE",
    "029": "GEOGRAPHY",
    "030": "ECONOMICS",
    "031": "CARNATIC MUSIC (VOCAL)",
    "032": "CARNATIC MUSIC( MELODIC INSTRUMENTS).",
    "033": "CARNATIC MUSIC ( PERCUSSION INSTRUMENTS - MRIDANGAM)",
    "034": "HINDUSTANI MUSIC (VOCAL)",
    "035": "HINDUSTANI MUSIC ( MELODIC INSTRUMENTS).",
    "036": "HINDUSTANI MUSIC ( PERCUSSION INSTRUMENTS)",
    "037": "PSYCHOLOGY",
    "039": "SOCIOLOGY",
    "041": "MATHEMATICS",
    "042": "PHYSICS",
    "043": "CHEMISTRY",
    "044": "BIOLOGY",
    "045": "BIOTECHNOLOGY",
    "046": "ENGINEERING GRAPHICS",
    "048": "PHYSICAL EDUCATION",
    "049": "PAINTING",
    "050": "GRAPHICS",
    "051": "SCULPTURE",
    "052": "APPLIED/ COMMERCIAL ART",
    "054": "BUSINESS STUDIES",
    "055": "ACCOUNTANCY",
    "056": "KATHAK - DANCE",
    "057": "BHARATNATYAM - DANCE",
    "058": "KUCHIPUDI-DANCE",
    "059": "ODISSI - DANCE",
    "060": "MANIPURI - DANCE",
    "061": "KATHAKALI - DANCE",
    "064": "HOME SCIENCE",
    "265": "INFORMATICS PRACTICES (OLD)(Only for XII)",
    "065": "INFORMATICS PRACTICES (NEW)",
    "283": "COMPUTER SCIENCE (OLD) (Only for XII)",
    "083": "COMPUTER SCIENCE (NEW)",
    "066": "ENTREPRENEURSHIP",
    "073": "KNOWLEDGE TRADITION & PRACTICES OF INDIA",
    "074": "LEGAL STUDIES",
    "076": "NATIONAL CADET CORPS (NCC)",
}

CBSE_CLASS_10_SUBJECT_CODES = {
    # Language Subjects
    "002": "HINDI COURSE-A",
    "085": "HINDI COURSE-B",
    "184": "ENGLISH LANG & LIT.",
    "003": "URDU COURSE-A",
    "303": "URDU COURSE-B",
    "004": "PUNJABI",
    "005": "BENGALI",
    "006": "TAMIL",
    "007": "TELUGU",
    "008": "SINDHI",
    "009": "MARATHI",
    "010": "GUJARATI",
    "011": "MANIPURI",
    "012": "MALAYALAM",
    "013": "ODIA",
    "014": "ASSAMESE",
    "015": "KANNADA",
    "016": "ARABIC",
    "017": "TIBETAN",
    "018": "FRENCH",
    "020": "GERMAN",
    "021": "RUSSIAN",
    "023": "PERSIAN",
    "024": "NEPALI",
    "025": "LIMBOO",
    "026": "LEPCHA",
    "089": "TELUGU TELANGANA",
    "092": "BODO",
    "093": "TANGKHUL",
    "094": "JAPANESE",
    "095": "BHUTIA",
    "096": "SPANISH",
    "097": "KASHMIRI",
    "098": "MIZO",
    "099": "BAH ASA MELAYU",
    "122": "SANSKRIT",
    "131": "RAI",
    "132": "GURUNG",
    "133": "TAMANG",
    "134": "SHERPA",
    "136": "THAI",

    # Academic Subjects
    "041": "MATHEMATICS -STANDARD",
    "241": "MATHEMATICS -BASIC",
    "086": "SCIENCE",
    "087": "SOCIAL SCIENCE",

    # Additional Academic Subjects
    "031": "CARNATIC MUSIC (VOCAL)",
    "032": "CARNATIC MUSIC (MELODIC INSTRUMENTS)",
    "033": "CARNATIC MUSIC (PERCUSSION INSTRUMENTS)",
    "034": "HINDUSTANI MUSIC (VOCAL)",
    "035": "HINDUSTANI MUSIC (MELODIC INSTRUMENTS)",
    "036": "HINDUSTANI MUSIC (PERCUSSION INSTRUMENTS)",
    "049": "PAINTING",
    "064": "HOME SCIENCE",
    "076": "NATIONAL CADET CORPS (NCC)",
    "165": "COMPUTER APPLICATIONS",
    "154": "ELEMENTS OF BUSINESS",
    "254": "ELEMENTS OF BOOK KEEPING & ACCOUNTANCY",

    # Skill Subjects
    "401": "RETAILING",
    "402": "INFORMATION TECHNOLOGY",
    "403": "SECURITY",
    "404": "AUTOMOTIVE",
    "405": "INTRODUCTION TO FINANCIAL MARKETS",
    "406": "INTRODUCTION TO TOURISM",
    "407": "BEAUTY & WELLNESS",
    "408": "AGRICULTURE",
    "409": "FOOD PRODUCTION",
    "410": "FRONT OFFICE OPERATIONS",
    "411": "BANKING & INSURANCE",
    "412": "MARKETING & SALES",
    "413": "HEALTH CARE",
    "414": "APPAREL",
    "415": "MEDIA",
    "416": "MULTI SKILL FOUNDATION COURSE",
    "417": "ARTIFICIAL INTELLIGENCE",

}

ROLL_NUMBER_LENGTH = 8
NAME_STARTING = 13
NAME_ENDING = 65
SUBJECT_STARTING = 64
SUBJECT_ENDING = 113
GRADE_STARTING = 64
GRADE_ENDING = 113
RESULT_STARTING = 114
RESULT_ENDING = 121


def is_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def convert_subject_code_to_name(subject_codes_list):
    sub_names = []
    for subjectCode in subject_codes_list:
        if subjectCode in CBSE_CLASS_10_SUBJECT_CODES:
            sub_names.append(CBSE_CLASS_10_SUBJECT_CODES[subjectCode])
        else:
            sub_names.append(subjectCode)
    return sub_names


def get_subject_headings(file_name):
    heading_subjects = []
    heading_codes = []
    total_students = 0

    file = open(file_name, "r")
    lines = file.readlines()
    i = 0
    while i <= len(lines) - 2:
        name_line = lines[i]
        if is_integer(name_line[:ROLL_NUMBER_LENGTH]):
            total_students += 1
            sub_codes = name_line[SUBJECT_STARTING:SUBJECT_ENDING].split()
            sub_names = convert_subject_code_to_name(sub_codes)

            for sub in sub_names:
                if sub not in heading_subjects:
                    heading_subjects.append(sub)

            for sub_code in sub_codes:
                if sub_code not in heading_codes:
                    heading_codes.append(sub_code)
        i += 1
    print(total_students)
    file.close()
    return heading_subjects, heading_codes


original_file_name = r"input\original files\10th Result.txt"
subject_headings, subject_codes = get_subject_headings(original_file_name)
heading = ["Roll Number", "Name"] + subject_headings
print(heading)
print(subject_codes)

