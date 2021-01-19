from collections import Counter
import PyPDF2
import re
import pprint
import xml.etree.ElementTree as ElemTree
import glob
import matplotlib.pyplot as plt
import pandas as pd
import os

TIER3_TARGETS = {'seal': ['seal', 'seald', 'seale', 'sealed', 'sealeing', 'seales', 'sealest', 'sealeth',
                          'sealid', 'sealing', 'sealinge', 'sealinges', 'sealings', 'seall', 'sealle', 'sealled',
                          'sealles', 'sealling', 'seallyng', 'seals', 'sealyd', 'sealynge', 'sealynges'],
                 'ink': ['inck', 'incke', 'inke', 'inked', 'inkes', 'inking', 'inks', 'inkyng', 'ynke'],
                 'wax': ['wax', 'vvax', 'vvaxe', 'vvaxing', 'vvaxinge', 'waxe', 'waxes', 'waxing', 'waxinge', 'wayyng',
                         'waxynge', 'wox', 'woxe', 'woxen'],
                 'paper': ['paper', 'papered', 'papering', 'papers', 'papyr'],
                 'pen': ['pen', 'penne', 'penned', 'pennes', 'penneth', 'penning', 'penninge', 'penninges', 'pennings',
                         'penns', 'pennyng', 'pennynge', 'pens', 'pensed', 'pent', 'pented', 'pent'],
                 'letter': ['letter', 'leters', 'lettered', 'letteres', 'lettering', 'letteris', 'letterred', 'letters',
                            'letteryd', 'lettred', 'lettres', 'lettrid', 'letttres'],
                 'hand': ['hand', 'hande', 'handst', 'haund', 'hond', 'honde'],
                 'character': ['character', 'caracter', 'caracters', 'carracters', 'charactars', 'characters',
                               'charactors', 'chararacters', 'charecters', 'charracter', 'charracters', 'charrecters'],
                 'book': ['boke', 'boked', 'bokes', 'bokis', 'booke', 'bookes', 'bookis', 'books', 'bookys', 'boooks',
                          'buk', 'buke', 'bukes'],
                 'read': ['read', 'readde', 'reade', 'readding', 'readed', 'readen', 'reades', 'readest', 'readeste',
                          'readeth', 'readethe', 'reading', 'readinge', 'readinges', 'readings', 'reads', 'readst',
                          'readyng', 'readynge', 'readynges', 'readyngs'],
                 'write': ['write', 'vvrite', 'vvriteing', 'vvriten', 'vvrites', 'vvritest', 'vvriteth', 'vvritethe',
                           'vvriting', 'vvritinge', 'vvritinges', 'vvritings', 'vvritten', 'vvrittes', 'vvritting',
                           'vvrittinge', 'vvrittynges', 'vvrote', 'vvryte', 'vvryten', 'vvryteth', 'vvryting',
                           'vvrytinges', 'wryten', 'wrytest', 'wryteth', 'wryting', 'wrytinge', 'vvrytten', 'writed',
                           'writeing', 'writeinge', 'writeings', 'writen', 'writest', 'writeth', 'writing', 'writinge',
                           'writinges', 'writingis', 'writings', 'writis', 'writiug', 'writst', 'writteinge',
                           'writteinges', 'written', 'writtes', 'writteth', 'writteynge', 'writing', 'writtinge',
                           'writtinges', 'writtings', 'writtten', 'writtyng', 'writtynge', 'writtynges', 'writyng',
                           'writynge', 'writynges', 'writyngs', 'wrot', 'wrote', 'wroted', 'wrttings', 'vvrytings',
                           'wrytings', 'wrytten', 'wrytyng', 'wrytynge', 'wrytynges', 'wrytyngs', 'wryte'],
                 'blush': ['blush', 'blushd', 'blushe', 'blushinges', 'blushes', 'blushings', 'blussh', 'blusshe',
                           'blvsh'],
                 'brow': ['brow', 'brovv', 'brovve', 'browe'],
                 'counterfeit': ['counterfeit', 'conterfeit', 'conterfeyt', 'countefeit', 'counterfait', 'counterfaite',
                                 'counterfaited', 'counterfaites', 'counterfaiting', 'counterfaitinge', 'counterfaits',
                                 'counterfayte', 'counterfayted', 'counterfaytinge', 'counterfayts', 'counterfeiteth',
                                 'counterfeiting', 'counterfeitinge', 'counterfeitings', 'counterfeitly',
                                 'counterfeits',
                                 'counterfeitted', 'counterfeityng', 'counterfet', 'counterfeted', 'counterfeteth',
                                 'counterfeting', 'counterfetlie', 'counterfetly', 'counterfeyeth', 'counterfeyting',
                                 'counterfeytinge', 'counterfeyts', 'counterfeytyng', 'counterfeytynge', 'counterfiets',
                                 'counterfit', 'counterfited', 'counterfiting', 'counterfits', 'counterseit',
                                 'counterseiting', 'covnterfaite'],
                 'forgery': ['forgery', 'forgerie', 'forgeries', 'forgerye', 'forgeryes'],
                 'secretary': ['secretary', 'secretarie', 'secretaries', 'secretarye', 'secretaryes'],
                 'sign': ['sign', 'seigned', 'signd', 'signde', 'signe', 'signed', 'signeing', 'signes', 'signeth',
                          'signing', 'signinge', 'signings', 'signs', 'signyng', 'sygne', 'sygnes', 'sygnyd'],
                 'subscribe': ['subscribe', 'subcribe', 'subcribed', 'subscrib', 'subscribd', 'subscribed',
                               'subscribeing',
                               'subscribes', 'subscribest', 'subscibeth', 'subscribing', 'subscribinge', 'subscribyd',
                               'subscibyng', 'subsciue', 'subscriuing', 'subscrive', 'subscrived', 'subscriving',
                               'subscrybed', 'subscrybyd', 'suscribe', 'suscribed'],
                 'binding': ['binding', 'bindding', 'bindeing', 'bindinge', 'bindinges', 'bindings', 'bindyng',
                             'bindynge',
                             'bound', 'bounden', 'bovnd', 'bovnden', 'byndeing', 'bynding', 'byndinge', 'byndings',
                             'byndyng', 'byndynge'],
                 'instrument': ['instrument', 'instruements', 'instrumens', 'instrumente', 'instrumented',
                                'instrumentes'
                                'instruments', 'instrumentys', 'instrustruments', 'instrvment', 'instrvments',
                                'instument',
                                'instuments', 'iustruments'],
                 'token': ['token', 'toaken', 'toakens', 'tokene', 'tokenes', 'tokeneth', 'tokening', 'tokeninge'
                                                                                                      'tokens',
                           'tokenyng', 'tokenynge', 'tokyn', 'tokyns']
                 }

TIER2_TARGETS = {'impress': ['impress', 'empressed', 'impresse', 'impressed', 'impresses', 'impressest', 'impresseth',
                             'impressid', 'impressing', 'impressyd', 'imprest'],
                 'print': ['print', 'printe', 'printed', 'printes', 'printest', 'printeth', 'printid', 'printing',
                           'printinge', 'printings', 'prints', 'printst', 'printyd', 'printyng', 'printynge', 'prynt',
                           'prynte', 'prynted', 'pryntid', 'prynts'],
                 'imprint': ['imprint', 'emprinted', 'enprint', 'enprynted', 'enpryntyd', 'imprinted', 'imprintes',
                             'imprinteth', 'imprinting', 'imprintinge', 'imprints', 'imprintyng', 'imprintynge',
                             'imprynted', 'impryntyd', 'inprint'],
                 'engrave': ['engrave', 'engraue', 'engraued', 'engrauen', 'engraues', 'engraueth', 'engrauing',
                             'engrauinge', 'engrauings', 'engrauyng', 'engravd', 'engraved', 'engraven', 'engraves',
                             'engraveth', 'engraving', 'engravings', 'ingrauings', 'ingrauyng', 'ingravigs'],
                 'sheet': ['sheet', 'sheete', 'sheeted', 'sheetes', 'sheeting', 'sheets', 'shetes'],
                 'parchment': ['parchment', 'parchement', 'parchments', 'partchment'],
                 'gall': ['gall', 'galle', 'lalles', 'galls'],
                 'inkhorn': ['inkhorn', 'inckehorne', 'inckhorne', 'inkehornes', 'inkhorne'],
                 'lines': ['lines', 'lynes'],
                 'handwriting': ['handwriting', 'handewriting', 'handewritynge', 'handvvriting', 'handwrite',
                                 'handwritings', 'handwrityng'],
                 'blotting': ['blotting', 'blot', 'blott', 'blotte', 'blotted', 'blottest', 'blotteth', 'blottings',
                              'blottyng', 'blottynge'],
                 'dashes': ['dashes', 'dash', 'dashd', 'dashe', 'dashe', 'dashed', 'dashest', 'dasheth', 'dashste',
                            'dasht', 'dashyd', 'dasshe', 'dasshed', 'dasshes', 'dasshest', 'dassheth'],
                 'script': ['script', 'scripts'],
                 'scribe': ['scribe', 'scribed', 'scribes', 'scribing', 'scribyng', 'scriue', 'scriue', 'scrive',
                            'scrybe', 'scrybes'],
                 'pounce': ['pounce', 'pouncd', 'pounced', 'pounces', 'pounceth', 'pouncing', 'pouncinge'],
                 'signature': ['signature', 'signatouris', 'signatures'],
                 'scrivener': ['scrivener', 'scriuener', 'scriueners', 'scriveners', 'scrivners'],
                 'leaf': ['leaf', 'leafe', 'leafed', 'leafes', 'leaffed', 'leafing', 'leafs', 'leaues', 'leaves',
                          'leavs', 'leeues', 'leeves'],
                 'blanks': ['blanks', 'blanckes', 'blancks', 'blankes', 'blanks'],
                 'signet': ['signet', 'fignet', 'signets'],
                 'notary': ['notary', 'notarie', 'notaries', 'notarye', 'notaryes', 'nottar']
                 }

TIER1_TARGETS = {'penknife': ['penknife', 'penknifes', 'penknives'],
                 'quill': ['quill', 'quile', 'quilled', 'quilles', 'quills', 'quyl', 'quyll', 'quylle', 'quyls'],
                 'label': ['label', 'labell', 'labelled', 'labelling', 'labells', 'labels', 'lybels'],
                 'engrossing': ['engrossing', 'ingrossing', 'engrosying'],
                 'orthography': ['orthography', 'orthographie', 'orthographies', 'orthographye', 'ortographie',
                                 'ortography', 'ortographye'],
                 'alphabet': ['alphabet', 'alphabete', 'alphabets'],
                 'copybook': ['copybook', 'copybooks'],
                 'quire': ['quire', 'quier', 'quiers', 'quired', 'quires', 'quireth', 'quiring', 'quyre',
                           'quyred', 'quyreth', 'quyring'],
                 'emboss': ['emboss', 'embosse', 'embossed', 'embossing', 'embost', 'imbossed', 'imbossing', 'imbost'],
                 'autograph': ['autograph'],
                 'manuscript': ['manuscript', 'manuscrips', 'manuscripta', 'manuscriptes', 'manuscripts',
                                'manvscripts'],
                 'calligraphy': ['calligraphy', 'caligraphy'],
                 'catchword': ['catchword', 'catchwords'],
                 'broadside': ['broadside', 'broadsides'],
                 'physiognomy': ['physiognomy', 'phisnomye', 'physiognomie', 'physiognomies', 'physiognomye',
                                 'physnomy']
                 }


# HELPER FUNCTIONS
def xml_to_tuple(xml, source_id='EP'):
    '''Take an xml file and returns a tuple with
        the Title, Author, Dates, and Full Text'''

    # SETUP XML PARSER
    tree = ElemTree.parse(xml)
    root = tree.getroot()

    # GET THE SOURCE OF XML FILE
    if source_id == 'EMED':
        original_text = root.findall(".//{http://www.tei-c.org/ns/1.0\}orig")
        cleaned_text = clean_EMED_xml(original_text)
    elif source_id == 'EP':
        original_text = root.findall(".//{http://www.tei-c.org/ns/1.0}w")
        cleaned_text = " ".join([item.text for item in original_text])
    else:
        original_text = None
        cleaned_text = None

    # GET (all) TITLES, AUTHORS, AND DATES
    titles = root.findall(".//{http://www.tei-c.org/ns/1.0}title")
    authors = root.findall(".//{http://www.tei-c.org/ns/1.0}author")
    dates = root.findall(".//{http://www.tei-c.org/ns/1.0}date")

    try:
        relevant_title = titles[0].text
    except IndexError:
        relevant_title = 'No Title Given'
    try:
        relevant_author = authors[0].text
    except IndexError:
        relevant_author = 'No Author Given'
    # These two functions clean up the text and dates
    relevant_dates = get_rel_dates(dates, source_id)
    return (relevant_title, relevant_author, relevant_dates, cleaned_text)


# Helps xml_to_tuple()
def get_rel_dates(dates, source_id='EP'):
    """The EMED files include publication and creation dates
        while the EP files only include nebulous dates. I've
        listed the dates from the EP files as the creation date?"""

    creation_date = None
    pub_date = None
    if source_id == 'EMED':
        for date in dates:
            date_attribute = date.attrib.get('type')
            if date_attribute == 'creation_date':
                creation_date = date.text
            elif date_attribute == 'publication_date':
                pub_date = date.text
        relevant_dates = [creation_date, pub_date]
    elif source_id == 'EP':
        pub_date = dates[1].text
        relevant_dates = [dates[1].text, None]
    return pub_date


# Helps xml_to_tuple()
def clean_EMED_xml(txt):
    """For some reason, the EMED files had a habit of replacing
        apostrophes with weird characters. This function cleans
        that up."""

    cleaned_text = []
    for item in txt:
        if item.text:
            if chr(8217) in item.text:
                cleaned_text.append(item.text.replace(chr(8217), chr(39)))
            if ' ' in item.text:
                cleaned_text.append(item.text.replace(' ', ''))
            else:
                cleaned_text.append(item.text)
    return " ".join(cleaned_text)

# Independent Helper
def publish_df_to_xl(df, filename):
    """Takes Pandas Dataframe object and filename. Creates XL sheet with filename"""
    try:
        df.to_excel(filename, index=False)
    except PermissionError:
        print('Error: Cannot open XL sheet')
    return

# Helps search_tiers (below)
def search_text_for_keywords(txt, target_words_dict):
    """Takes list of split words and a dictionary with targets words as keys and
        spelling equivalences as values. Returns count of target keywords"""
    keyword_count = dict.fromkeys(target_words_dict.keys(), 0)
    for word in txt:
        for target_equivalences in target_words_dict.items():
            if word in target_equivalences:
                keyword_count[word] += 1
                continue

    to_delete = [k for k, v in keyword_count.items() if v == 0]
    for key in to_delete: del keyword_count[key]
    return keyword_count


def search_tiers(txt_tuple):
    """Takes a list of tuples and searches them for the three tiers of target words.
        Returns a list of tuples with Title Author SplitText KeywordCounts"""
    list_of_works_with_keywords = []

    for work in txt_tuple:
        keywords_tier3 = search_text_for_keywords(work[3], TIER3_TARGETS)
        keywords_tier2 = search_text_for_keywords(work[3], TIER2_TARGETS)
        keywords_tier1 = search_text_for_keywords(work[3], TIER1_TARGETS)

        sorted_keywords_tier3 = {k: v for k, v in reversed(sorted(keywords_tier3.items(), key=lambda item: item[1]))}
        sorted_keywords_tier2 = {k: v for k, v in reversed(sorted(keywords_tier2.items(), key=lambda item: item[1]))}
        sorted_keywords_tier1 = {k: v for k, v in reversed(sorted(keywords_tier1.items(), key=lambda item: item[1]))}

        # *change ' '.join(work[3]) to work[3] for Split Text
        list_of_works_with_keywords.append((work[0], work[1], ' '.join(work[3]),
                                            sorted_keywords_tier3,
                                            sorted_keywords_tier2,
                                            sorted_keywords_tier1))  # title author text keywordcounts
    return list_of_works_with_keywords


def enact_conditions(searched_text):
    """If a text meets *one* of three conditions, it is considered relevant. Here are the conditions:
        1.	Texts with at least 4 different terms from the Tier 3 OR
        2.	Texts with at least 1 term from Tier 2 and at least 2 more terms from the Tier 2 or Tier 3
        3.	Texts with at least 1 term from the Tier 1 and at least one more term from any of the Tiers

        In order to do this, each Tier is assigned a point value. Tier 3 = 1 point, Tier 2 = 2 points, Tier 1 = 3 points
        If a text totals to 4 or more points, it is included. There is one way a text could be included but meet the
        conditions: two words from Tier 2 = 4 points. We will include those also."""
    relevant_works = []
    for work in searched_texts:
        tier3_count = work[3]
        tier2_count = work[4]
        tier1_count = work[5]

        score = len(tier3_count)*1 + len(tier2_count)*2 + len(tier1_count)*3
        total_words = sum(tier3_count.values()) + sum(tier2_count.values()) + sum(tier1_count.values())
        if score >= 4:
            relevant_works.append(tuple([work[0], work[1], work[2], work[3], work[4], work[5], score, total_words]))

    return relevant_works


if __name__ == '__main__':
    folder = r'C:\Users\richa\Python_Projects\My_Projects\TextProcessor\\'

    # For now, I am only using EP files
    EP_files = glob.glob(f'{folder}Texts\EP_xml\\*.xml')

    # List of Tuples. get a list of tuples with Title, Author, Dates, and Full Text
    full_texts_db = [xml_to_tuple(file, 'EP') for file in EP_files]

    # List of Tuples. change the Full Text to Split Text (create list of words)
    split_texts_db = [(item[0], item[1], item[2], item[3].split(' '))
                      for item in full_texts_db]

    full_texts_df = pd.DataFrame(full_texts_db,
                                 columns=['title', 'author', 'publication_date', 'full_text'])

    # I *think* that the split texts df is useless but here it is anyway
    split_texts_df = pd.DataFrame(full_texts_db,
                                 columns=['title', 'author', 'publication_date', 'split_text'])

    publish_df_to_xl(full_texts_df, f'{folder}EP_main.xlsx')

    # List of Tuples with Title, Author, Full Text*, Tier3 Dict, Tier2 Dict, Tier1 Dict
    # (can be Split Text, see note in search_tiers())
    searched_texts = search_tiers(split_texts_db)
    searched_texts_df = pd.DataFrame(searched_texts,
                                     columns=['title', 'author', 'text', 'tier3_words', 'tier2_words', 'tier1_words'])
    publish_df_to_xl(searched_texts_df, f'{folder}EP_keywords.xlsx')

    # List of Tuples. Here are the texts that meet the three conditions.
    # They have the same columns as the searched texts
    relevant_texts = enact_conditions(searched_texts)
    #print(relevant_texts[6])
    relevant_texts_df = pd.DataFrame(relevant_texts,
                                     columns=['title', 'author', 'text', 'tier3_words', 'tier2_words',
                                              'tier1_words', 'score', 'total_relevant_words'])
    publish_df_to_xl(relevant_texts_df, f'{folder}EP_relevant.xlsx')