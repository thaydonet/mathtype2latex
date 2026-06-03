class RecordType:
    END = 0
    LINE = 1
    CHAR = 2
    TMPL = 3
    PILE = 4
    MATRIX = 5
    EMBELL = 6
    RULER = 7
    FONT_STYLE_DEF = 8
    SIZE = 9
    FULL = 10
    SUB = 11
    SUB2 = 12
    SYM = 13
    SUBSYM = 14
    COLOR = 15
    COLOR_DEF = 16
    FONT_DEF = 17
    EQN_PREFS = 18
    ENCODING_DEF = 19
    FUTURE = 100
    ROOT = 255


class OptionType:
    MtefOptNudge = 0x08
    MtefOptCharEmbell = 0x01
    MtefOptCharFuncStart = 0x02
    MtefOptCharEncChar8 = 0x04
    MtefOptCharEncChar16 = 0x10
    MtefOptCharEncNoMtcode = 0x20
    MtefOptLineNull = 0x01
    mtefOPT_LP_RULER = 0x02
    MtefOptLineLspace = 0x04
    MtefOptLpRuler = 0x02
    MtefColorCmyk = 0x01
    MtefColorSpot = 0x02
    MtefColorName = 0x04
    mtefCOLOR_CMYK = 0x01
    mtefCOLOR_SPOT = 0x02
    mtefCOLOR_NAME = 0x04


class CharTypeface:
    fnTEXT = 1
    fnFUNCTION = 2
    fnVARIABLE = 3
    fnLCGREEK = 4
    fnUCGREEK = 5
    fnSYMBOL = 6
    fnVECTOR = 7
    fnNUMBER = 8
    fnUSER1 = 9
    fnUSER2 = 10
    fnMTEXTRA = 11
    fnTEXT_FE = 12
    fnEXPAND = 22
    fnMARKER = 23
    fnSPACE = 24


class MtLine:
    def __init__(self):
        self.nudgeX = 0
        self.nudgeY = 0
        self.lineSpace = 0
        self.null = False


class MtChar:
    def __init__(self):
        self.nudgeX = 0
        self.nudgeY = 0
        self.options = 0
        self.typeface = 0
        self.mtcode = 0
        self.bits8 = 0
        self.bits16 = 0


class MtTmpl:
    def __init__(self):
        self.nudgeX = 0
        self.nudgeY = 0
        self.selector = 0
        self.variation = 0
        self.options = 0


class MtPile:
    def __init__(self):
        self.nudgeX = 0
        self.nudgeY = 0
        self.halign = 0
        self.valign = 0


class MtMatrix:
    def __init__(self):
        self.nudgeX = 0
        self.nudgeY = 0
        self.valign = 0
        self.h_just = 0
        self.v_just = 0
        self.rows = 0
        self.cols = 0


class MtEmbellRd:
    def __init__(self):
        self.options = 0
        self.nudgeX = 0
        self.nudgeY = 0
        self.embellType = 0


class MtSize:
    def __init__(self):
        self.lsize = 0
        self.dsize = 0


class MtfontStyleDef:
    def __init__(self):
        self.fontDefIndex = 0
        self.name = ""


class MtfontDef:
    def __init__(self):
        self.encDefIndex = 0
        self.name = ""


class MtColorDefIndex:
    def __init__(self):
        self.index = 0


class MtColorDef:
    def __init__(self):
        self.values = []
        self.name = ""


class MtEqnPrefs:
    def __init__(self):
        self.sizes = []
        self.spaces = []
        self.styles = []


class MtAST:
    def __init__(self, tag=0, value=None, children=None):
        self.tag = tag
        self.value = value
        self.children = children if children is not None else []


class SelectorType:
    tmANGLE = 0
    tmPAREN = 1
    tmBRACE = 2
    tmBRACK = 3
    tmBAR = 4
    tmDBAR = 5
    tmFLOOR = 6
    tmCEILING = 7
    tmOBRACK = 8
    tmINTERVAL = 9
    tmROOT = 10
    tmFRACT = 11
    tmUBAR = 12
    tmOBAR = 13
    tmARROW = 14
    tmINTEG = 15
    tmSUM = 16
    tmPROD = 17
    tmCOPROD = 18
    tmUNION = 19
    tmINTER = 20
    tmINTOP = 21
    tmSUMOP = 22
    tmLIM = 23
    tmHBRACE = 24
    tmHBRACK = 25
    tmLDIV = 26
    tmSUB = 27
    tmSUP = 28
    tmSUBSUP = 29
    tmDIRAC = 30
    tmVEC = 31
    tmTILDE = 32
    tmHAT = 33
    tmARC = 34
    tmJSTATUS = 35
    tmSTRIKE = 36
    tmBOX = 37


class EmbellType:
    emb1DOT = 2
    emb2DOT = 3
    emb3DOT = 4
    emb1PRIME = 5
    emb2PRIME = 6
    embBPRIME = 7
    embTILDE = 8
    embHAT = 9
    embNOT = 10
    embRARROW = 11
    embLARROW = 12
    embBARROW = 13
    embR1ARROW = 14
    embL1ARROW = 15
    embMBAR = 16
    embOBAR = 17
    emb3PRIME = 18
    embFROWN = 19
    embSMILE = 20
    embX_BARS = 21
    embUP_BAR = 22
    embDOWN_BAR = 23
    emb4DOT = 24
    embU_1DOT = 25
    embU_2DOT = 26
    embU_3DOT = 27
    embU_4DOT = 28
    embU_BAR = 29
    embU_TILDE = 30
    embU_FROWN = 31
    embU_SMILE = 32
    embU_RARROW = 33
    embU_LARROW = 34
    embU_BARROW = 35
    embU_R1ARROW = 36
    embU_L1ARROW = 37
