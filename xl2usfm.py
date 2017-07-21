import codecs
import openpyxl
import pdb

def getBookCode(bkID):
    bk = {"17": "TIT"}
    global usfmString
    global curBook
    bknm = bk.get(bkID, "UNKNOWN")
    if(bkID > curBook):
       curBook = bkID
       usfmString += "\\id " + bknm 
    return bknm

def getBookInfo(vID):
    # 170101
    vID = str(vID)
    bkID = vID[:2]
    ch = vID[2:4]
    vs = vID[4:]
    return bkID, getBookCode(bkID), int(ch), int(vs)
    
wb = openpyxl.load_workbook("malOV.xlsx")
ws = wb.active

usfmString = ""
curChapter = curVerse = 0
curBook = '0'
pre_w = inf_w = ""
pre_lx = inf_lx = ""
pre_syn = inf_syn = ""
pre_morph = inf_morph = ""
pre_loc = inf_loc = ""

for i, row in enumerate(ws):
    if i > 0:
        MID = row[0].value
        ORDER = row[1].value
        VERSE = row[2].value
        WORD = row[3].value
        UWORD = row[4].value
        UMEDIEVAL = row[5].value
        LEXEME = row[6].value
        LEMMA = row[7].value
        MAL_ORDER = row[8].value
        MAL_VERSE = row[9].value
        MAL_TYPE = row[10].value
        MAL_GLOSS = row[11].value
        GLOSS = row[12].value
        ULEMMA = row[13].value
        SYN = row[14].value
        MORPH = row[15].value
        LEX = row[16].value
        PUNC = row[17].value

        bid, bnm, c, v = getBookInfo(MAL_VERSE)
        if(c > curChapter):
            curChapter = c
            usfmString += "\n\\c " + str(curChapter)
        if(v > curVerse):
            curVerse = v
            usfmString += "\n\\v " + str(curVerse) + " "
        srcloc = "MAL10OV:" + str(bid) + "." + str(c) + "." + str(v) + "." + str(MAL_ORDER)
        if MAL_TYPE == 'r':
            try:
                usfmString += "\\w " + MAL_GLOSS + '|strongs = "G' + str(LEXEME) + ":" + SYN + MORPH + '" srcloc = "' + srcloc + '"\\w* '
            except:
                usfmString += "\\add [" + MAL_GLOSS +  '] srcloc = "' + srcloc + '"\\add* '
        elif MAL_TYPE == 'ig' or MAL_TYPE == 'ab':
            usfmString += "\\del " + WORD + '|strongs = "G' + str(LEXEME) + ":" + SYN + MORPH + "\\del* "
        elif MAL_TYPE == 'pre':
            pre_w = MAL_GLOSS
            pre_lx = "G" + str(LEXEME)
            pre_syn = SYN
            pre_morph = MORPH + ", "
            pre_loc = srcloc + ", "
        elif MAL_TYPE == 'in':
            inf_w = MAL_GLOSS
            inf_lx = "G" + str(LEXEME)
            inf_syn = SYN
            inf_morph = MORPH  + ", "
            inf_loc = srcloc + ", "
        elif MAL_TYPE == 'su':
            usfmString += "\\w " + pre_w + inf_w + MAL_GLOSS + '|strongs = "' + pre_lx + ":" + pre_syn + pre_morph + inf_lx + ":" + inf_syn + inf_morph + " G" + str(LEXEME) + ":" + SYN + MORPH +  '" srcloc = "' + pre_loc + inf_loc + srcloc + '"\\w* '
            pre_w = inf_w = ""
            pre_lx = inf_lx = ""
            pre_syn = inf_syn = ""
            pre_morph = inf_morph = ""
            pre_loc = inf_loc = ""
        elif MAL_TYPE == 'im':
            usfmString += "\\imp " + WORD + '|strongs = "G' + str(LEXEME) + ":" + SYN + MORPH + "\\imp* "
        else:
            usfmString += "UNKNOWN WORD TYPE "
if i != 0:
    o = codecs.open(str(bid) + bnm + "MOV.usfm", mode = "w", encoding="utf-8")
    o.write(usfmString)
o.close()
