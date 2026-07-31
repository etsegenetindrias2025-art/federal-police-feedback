# utils/translations.py

TRANSLATIONS = {
    "am": {
        "hdrTitle": "የኢትዮጵያ ፌዴራል ፖሊስ",
        "hdrSub": "የዜጎች አስተያየት መስጫ ስርዓት",
        "mainHeading": "በተሰጠዎት አገልግሎት ምን ያህል ረክተዋል?",
        "mainDesc": "እባክዎ የተሰማዎትን እርካታ ከታች ከሚገኙት ስሜቶች አንዱን በመምረጥ በዝግታ ይግለጹን።",
        "l1": "በጣም አላረኩም", "l2": "አላረኩም", "l3": "መካከለኛ", "l4": "ረክቻለሁ", "l5": "በጣም ረክቻለሁ",
        "lblComment": "አስተያየትዎን በጽሁፍ ይጻፉ (አማራጭ):",
        "commentPlaceholder": "ስለተሰጠው አገልግሎትዎ ሃሳብ ወይም አስተያየት ይጻፉ...",
        "recStatus": "ድምጽ ለመቅረጽ ማይኩን ይጫኑ",
        "recActive": "ድምጽ በመቅረጽ ላይ ነው...",
        "recDone": "ድምጽዎ ተቀምጧል!",
        "btnBack": "ተመለስ",
        "btnSubmit": "አስገባና ጨርስ",
        "errRating": "እባክዎ ከመቀጠልዎ በፊት የሬቲንግ ስሜት (Rating) ይምረጡ!",
        "successTitle": "አመሰግናለሁ! 😊",
        "successDesc": "አስተያየትዎ በተሳካ ሁኔታ ተመዝግቧል ✅"
    },
    "en": {
        "hdrTitle": "Ethiopian Federal Police",
        "hdrSub": "Citizen Feedback System",
        "mainHeading": "How satisfied are you with the service provided?",
        "mainDesc": "Please gently select your satisfaction level from the options below.",
        "l1": "Very Dissatisfied", "l2": "Dissatisfied", "l3": "Neutral", "l4": "Satisfied", "l5": "Very Satisfied",
        "lblComment": "Write your comment (Optional):",
        "commentPlaceholder": "Write your thoughts or feedback about the service...",
        "recStatus": "Click mic to record voice",
        "recActive": "Recording voice...",
        "recDone": "Voice recorded!",
        "btnBack": "Back",
        "btnSubmit": "Submit & Finish",
        "errRating": "Please select a rating before proceeding!",
        "successTitle": "thank you! 😊",
        "successDesc": "your feedback has been submitted successfully ✅"
    },
    "om": {
        "hdrTitle": "Poolisii Federaalaa Itoophiyaa",
        "hdrSub": "Sirna Yaada Kennitoota Tajaajilaa",
        "mainHeading": "Tajaajila kennameen hammam quufteetta?",
        "mainDesc": "Mee miira kee armaan gadii keessaa filachuun nuuf ibsi.",
        "l1": "Baay'ee hin quufne", "l2": "Hin quufne", "l3": "Giddu-galeessa", "l4": "Quufeera", "l5": "Baay'ee quufeera",
        "lblComment": "Yaada kee barreeffamaan barreessi (Filannoo):",
        "commentPlaceholder": "Tajaajila kenname irratti yaada kee barreessi...",
        "recStatus": "Sagalee galchuuf maayikiinii tuqaa",
        "recActive": "Sagalee galchaa jira...",
        "recDone": "Sagaleen kee kuufameera!",
        "btnBack": "Duubatti",
        "btnSubmit": "Galchi Xumuri",
        "errRating": "Mee osoo hin darbiin dura sadarkaa (Rating) filadhu!",
        "successTitle": "galatoomaa! 😊",
        "successDesc": "yaadni kee milkaa'inaan galmeeffameera ✅"
    }
}

def get_translation(lang):
    return TRANSLATIONS.get(lang, TRANSLATIONS['en'])