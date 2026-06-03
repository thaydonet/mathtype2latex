from io import BytesIO
from .ole_util.helper import Helper
from .ole_util.ole import Ole
from .record import (
    MtLine,
    MtChar,
    MtTmpl,
    MtPile,
    MtMatrix,
    MtEmbellRd,
    MtfontStyleDef,
    MtSize,
    MtfontDef,
    MtColorDefIndex,
    MtColorDef,
    MtEqnPrefs,
    RecordType,
    OptionType,
    CharTypeface,
    SelectorType,
    EmbellType,
    MtAST,
)
from .chars import Chars, SpecialChar

oleCbHdr = 28


class MTEF:
    def __init__(self):
        self.mMtefVer = 0
        self.mPlatform = 0
        self.mProduct = 0
        self.mVersion = 0
        self.mVersionSub = 0
        self.mApplication = ""
        self.mInline = 0
        self.reader = None
        self.ast = None
        self.nodes = []
        self.Valid = False

    def _read1(self):
        b = self.reader.read(1)
        if b is None or len(b) == 0:
            return None
        return b[0]

    def readRecord(self):
        self.Valid = True
        self.mMtefVer = self._read1()
        if self.mMtefVer is None:
            return
        self.mPlatform = self._read1()
        self.mProduct = self._read1()
        self.mVersion = self._read1()
        self.mVersionSub = self._read1()
        if self.mMtefVer >= 5:
            self.mApplication, _ = self.readNullTerminatedString()
            self.mInline = self._read1() or 0

        while True:
            record = self._read1()
            if record is None:
                break

            if record >= RecordType.FUTURE:
                skipLen = Helper.bytes2int(self.reader.read(1))
                self.reader.seek(skipLen, 1)
                continue

            if record == RecordType.END:
                self.nodes.append(MtAST(RecordType.END, None, None))
            elif record == RecordType.LINE:
                line = MtLine()
                self.readLine(line)
                self.nodes.append(MtAST(RecordType.LINE, line, None))
            elif record == RecordType.CHAR:
                char = MtChar()
                self.readChar(char)
                self.nodes.append(MtAST(RecordType.CHAR, char, None))
            elif record == RecordType.TMPL:
                tmpl = MtTmpl()
                self.readTMPL(tmpl)
                self.nodes.append(MtAST(RecordType.TMPL, tmpl, None))
            elif record == RecordType.PILE:
                pile = MtPile()
                self.readPile(pile)
                self.nodes.append(MtAST(RecordType.PILE, pile, None))
            elif record == RecordType.MATRIX:
                matrix = MtMatrix()
                self.readMatrix(matrix)
                self.nodes.append(MtAST(RecordType.MATRIX, matrix, None))
            elif record == RecordType.EMBELL:
                embell = MtEmbellRd()
                self.readEmbell(embell)
                self.nodes.append(MtAST(RecordType.EMBELL, embell, None))
            elif record == RecordType.FONT_STYLE_DEF:
                fsDef = MtfontStyleDef()
                fsDef.fontDefIndex = Helper.bytes2int(self.reader.read(1))
                fsDef.name, _ = self.readNullTerminatedString()
            elif record == RecordType.SIZE:
                mtSize = MtSize()
                mtSize.lsize = Helper.bytes2int(self.reader.read(1))
                mtSize.dsize = Helper.bytes2int(self.reader.read(1))
            elif record == RecordType.SUB:
                self.nodes.append(MtAST(RecordType.SUB, None, None))
            elif record == RecordType.SUB2:
                self.nodes.append(MtAST(RecordType.SUB2, None, None))
            elif record == RecordType.SYM:
                self.nodes.append(MtAST(RecordType.SYM, None, None))
            elif record == RecordType.SUBSYM:
                self.nodes.append(MtAST(RecordType.SUBSYM, None, None))
            elif record == RecordType.FONT_DEF:
                fdef = MtfontDef()
                fdef.encDefIndex = Helper.bytes2int(self.reader.read(1))
                fdef.name, _ = self.readNullTerminatedString()
                self.nodes.append(MtAST(RecordType.FONT_DEF, fdef, None))
            elif record == RecordType.COLOR:
                cIndex = MtColorDefIndex()
                cIndex.index = Helper.bytes2int(self.reader.read(1))
            elif record == RecordType.COLOR_DEF:
                cDef = MtColorDef()
                self.readColorDef(cDef)
            elif record == RecordType.FULL:
                self.nodes.append(MtAST(RecordType.FULL, None, None))
            elif record == RecordType.EQN_PREFS:
                prefs = MtEqnPrefs()
                self.readEqnPrefs(prefs)
                self.nodes.append(MtAST(RecordType.EQN_PREFS, prefs, None))
            elif record == RecordType.ENCODING_DEF:
                enc, _ = self.readNullTerminatedString()
                self.nodes.append(MtAST(RecordType.ENCODING_DEF, enc, None))
            else:
                self.Valid = False
                break

        return None

    def readNullTerminatedString(self):
        buf = []
        while True:
            p = self.reader.read(1)
            if p is None or len(p) != 1:
                break
            if p[0] == 0:
                break
            buf.append(p)
        return b"".join(buf), None

    def readLine(self, line):
        options = self._read1()
        if options is None:
            return
        if OptionType.MtefOptNudge == (OptionType.MtefOptNudge & options):
            line.nudgeX, line.nudgeY, _ = self.readNudge()
        if OptionType.MtefOptLineLspace == (OptionType.MtefOptLineLspace & options):
            line.lineSpace = self._read1() or 0
        if OptionType.mtefOPT_LP_RULER == (OptionType.mtefOPT_LP_RULER & options):
            nStops = self._read1() or 0
            for _ in range(nStops):
                self._read1()
                self._read2()
        if OptionType.MtefOptLineNull == (OptionType.MtefOptLineNull & options):
            line.null = True

    def _read2(self):
        b = self.reader.read(2)
        if b is None or len(b) < 2:
            return 0
        return int.from_bytes(b, byteorder="little")

    def readChar(self, char):
        options = self._read1()
        if options is None:
            return
        if OptionType.MtefOptNudge == (OptionType.MtefOptNudge & options):
            char.nudgeX, char.nudgeY, _ = self.readNudge()
        char.typeface = self._read1() or 0
        if OptionType.MtefOptCharEncNoMtcode != (
            OptionType.MtefOptCharEncNoMtcode & options
        ):
            char.mtcode = self._read2()
        if OptionType.MtefOptCharEncChar8 == (OptionType.MtefOptCharEncChar8 & options):
            char.bits8 = self._read1() or 0
        if OptionType.MtefOptCharEncChar16 == (
            OptionType.MtefOptCharEncChar16 & options
        ):
            char.bits16 = self._read2()

    def readNudge(self):
        b1 = self._read2()
        b2 = self._read2()
        if b1 == 128 or b2 == 128:
            nudgeX = self._read2()
            nudgeY = self._read2()
            return nudgeX, nudgeY, None
        return b1, b2, None

    def readTMPL(self, tmpl):
        options = self._read1()
        if options is None:
            return
        if OptionType.MtefOptNudge == (OptionType.MtefOptNudge & options):
            tmpl.nudgeX, tmpl.nudgeY, _ = self.readNudge()
        tmpl.selector = self._read1() or 0
        byte1 = self._read1()
        if byte1 is None:
            return
        if 0x80 == (byte1 & 0x80):
            byte2 = self._read1() or 0
            tmpl.variation = (byte1 & 0x7F) | (byte2 << 8)
        else:
            tmpl.variation = byte1
        tmpl.options = self._read1() or 0

    def readPile(self, pile):
        options = self._read1()
        if options is None:
            return
        if OptionType.MtefOptNudge == (OptionType.MtefOptNudge & options):
            pile.nudgeX, pile.nudgeY, _ = self.readNudge()
        pile.halign = self._read1() or 0
        pile.valign = self._read1() or 0

    def readMatrix(self, matrix):
        options = self._read1()
        if options is None:
            return
        if OptionType.MtefOptNudge == (OptionType.MtefOptNudge & options):
            matrix.nudgeX, matrix.nudgeY, _ = self.readNudge()
        matrix.valign = self._read1() or 0
        matrix.h_just = self._read1() or 0
        matrix.v_just = self._read1() or 0
        matrix.rows = self._read1() or 0
        matrix.cols = self._read1() or 0

    def readEmbell(self, embell):
        options = self._read1()
        if options is None:
            return
        if OptionType.MtefOptNudge == (OptionType.MtefOptNudge & options):
            embell.nudgeX, embell.nudgeY, _ = self.readNudge()
        embell.embellType = self._read1() or 0

    def readColorDef(self, colorDef):
        options = self._read1()
        if options is None:
            return
        if OptionType.mtefCOLOR_CMYK == (OptionType.mtefCOLOR_CMYK & options):
            for _ in range(4):
                colorDef.values.append(self._read2())
        else:
            for _ in range(3):
                colorDef.values.append(self._read2())
        if OptionType.mtefCOLOR_NAME == (OptionType.mtefCOLOR_NAME & options):
            colorDef.name, _ = self.readNullTerminatedString()

    def readEqnPrefs(self, eqnPrefs):
        self._read1()
        size = self._read1() or 0
        eqnPrefs.sizes, _ = self.readDimensionArrays(size)
        size = self._read1() or 0
        eqnPrefs.spaces, _ = self.readDimensionArrays(size)
        size = self._read1() or 0
        styles = []
        for _ in range(size):
            c = self._read1() or 0
            if c != 0:
                c = self._read1() or 0
            styles.append(c)
        eqnPrefs.styles = styles

    def readDimensionArrays(self, size):
        shareData = {"flag": True, "tmpStr": "", "count": 0, "array": []}

        def fx(x):
            if shareData["flag"]:
                shareData["flag"] = False
                if x == 0x00:
                    shareData["tmpStr"] += "in"
                elif x == 0x01:
                    shareData["tmpStr"] += "cm"
                elif x == 0x02:
                    shareData["tmpStr"] += "pt"
                elif x == 0x03:
                    shareData["tmpStr"] += "pc"
                elif x == 0x04:
                    shareData["tmpStr"] += "%"
            else:
                if x == 0x0A:
                    shareData["tmpStr"] += "."
                elif x == 0x0B:
                    shareData["tmpStr"] += "-"
                elif x == 0x0F:
                    shareData["flag"] = True
                    shareData["count"] += 1
                    shareData["array"].append(shareData["tmpStr"])
                    shareData["tmpStr"] = ""
                elif x <= 9:
                    shareData["tmpStr"] += str(x)

        while True:
            if shareData["count"] >= size:
                break
            ch = self._read1()
            if ch is None:
                break
            hi = (ch & 0xF0) // 16
            lo = ch & 0x0F
            fx(hi)
            fx(lo)
        return shareData["array"], None

    def Translate(self):
        latexStr, _ = self.makeLatex(self.ast)
        if self.Valid:
            return latexStr
        else:
            return ""

    def makeAST(self):
        ast = MtAST()
        ast.tag = 0xFF
        ast.value = None
        self.ast = ast
        stack = [ast]

        for node in self.nodes:
            if node.tag == RecordType.LINE:
                if stack:
                    parent = stack[-1]
                    parent.children.append(node)
                if not node.value.null:
                    stack.append(node)
            elif node.tag == RecordType.TMPL:
                if stack:
                    stack[-1].children.append(node)
                stack.append(node)
            elif node.tag == RecordType.PILE:
                if stack:
                    stack[-1].children.append(node)
                stack.append(node)
            elif node.tag == RecordType.MATRIX:
                if stack:
                    stack[-1].children.append(node)
                stack.append(node)
            elif node.tag == RecordType.END:
                if stack:
                    stack.pop()
            elif node.tag == RecordType.CHAR:
                if stack:
                    stack[-1].children.append(node)
            elif node.tag == RecordType.EMBELL:
                if stack:
                    parent = stack[-1]
                    parent.children.append(node)
                    embellType = node.value.embellType
                    if embellType in (
                        EmbellType.emb1DOT,
                        EmbellType.embHAT,
                        EmbellType.embOBAR,
                    ):
                        if len(parent.children) >= 2:
                            embellData = parent.children[-1]
                            charData = parent.children[-2]
                            parent.children = parent.children[:-2]
                            parent.children.append(embellData)
                            parent.children.append(charData)
                stack.append(node)
            elif node.tag in (
                RecordType.FONT_STYLE_DEF,
                RecordType.SIZE,
                RecordType.COLOR,
                RecordType.COLOR_DEF,
                RecordType.FONT_DEF,
                RecordType.EQN_PREFS,
                RecordType.ENCODING_DEF,
                RecordType.FULL,
                RecordType.SUB,
                RecordType.SUB2,
                RecordType.SYM,
                RecordType.SUBSYM,
            ):
                pass

    def _embell_command(self, embellType):
        if embellType == EmbellType.emb1DOT:
            return "\\dot"
        if embellType == EmbellType.embHAT:
            return "\\hat"
        if embellType == EmbellType.embOBAR:
            return "\\overline"
        if embellType == EmbellType.embTILDE:
            return "\\tilde"
        return None

    def _embell_text(self, embellType):
        command = self._embell_command(embellType)
        if command:
            return f" {command} "
        if embellType == EmbellType.emb1PRIME:
            return "'"
        if embellType == EmbellType.emb2PRIME:
            return "''"
        if embellType == EmbellType.emb3PRIME:
            return "'''"
        return ""

    def _children_latex(self, children):
        buf = ""
        idx = 0
        while idx < len(children):
            child = children[idx]
            if child.tag == RecordType.EMBELL and idx + 1 < len(children):
                command = self._embell_command(child.value.embellType)
                if command:
                    target, _ = self.makeLatex(children[idx + 1])
                    buf += f" {command}{{{target.strip()}}}"
                    idx += 2
                    continue
            _latex, _ = self.makeLatex(child)
            buf += _latex
            idx += 1
        return buf

    def makeLatex(self, ast):
        buf = ""
        if ast is None:
            return buf, None

        if ast.tag == 0xFF:
            buf += self._children_latex(ast.children)
            return buf, None

        if ast.tag == RecordType.CHAR:
            mtcode = ast.value.mtcode
            typeface = ast.value.typeface
            char = chr(mtcode) if mtcode < 0x10000 else "?"
            hexExtend = ""
            typefaceFmt = ""
            if (typeface - 128) == CharTypeface.fnMTEXTRA:
                hexExtend = "/mathmode"
            elif (typeface - 128) == CharTypeface.fnSPACE:
                hexExtend = "/mathmode"
            elif (typeface - 128) == CharTypeface.fnTEXT:
                typefaceFmt = "{ \\rm{ %s } }"
            hexCode = "%04x" % mtcode
            hexKey = "char/0x%s%s" % (hexCode, hexExtend)
            sChar = Chars.get(hexKey)
            used_mathmode_fallback = False
            if not sChar and not hexExtend:
                sChar = Chars.get("char/0x%s/mathmode" % hexCode)
                used_mathmode_fallback = sChar is not None
            if sChar:
                char = sChar
            else:
                sChar = SpecialChar.get(char)
                if sChar:
                    char = sChar
            if typefaceFmt != "" and not used_mathmode_fallback:
                char = typefaceFmt % char
            buf += char
            return buf, None

        elif ast.tag == RecordType.TMPL:
            tmpl = ast.value
            sel = tmpl.selector

            if sel in (
                SelectorType.tmANGLE,
                SelectorType.tmPAREN,
                SelectorType.tmBAR,
                SelectorType.tmINTERVAL,
            ):
                mainSlot = leftSlot = rightSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    leftSlot, _ = self.makeLatex(ast.children[1])
                if len(ast.children) > 2:
                    rightSlot, _ = self.makeLatex(ast.children[2])

                # Check if mainSlot already has array/matrix
                stripped_main = mainSlot.strip()
                if stripped_main.startswith(
                    "\\begin{array}"
                ) or stripped_main.startswith("\\begin{matrix}"):
                    # Already has array → just add left/right
                    mainStr = mainSlot
                    leftStr = "\\left %s" % leftSlot if leftSlot else ""
                    rightStr = "\\right %s" % rightSlot if rightSlot else ""
                    buf += "%s %s %s" % (leftStr, mainStr, rightStr)
                else:
                    # No array → keep original behavior with braces
                    mainStr = "{ %s }" % mainSlot if mainSlot else ""
                    leftStr = "\\left %s" % leftSlot if leftSlot else ""
                    rightStr = "\\right %s" % rightSlot if rightSlot else ""
                    buf += "%s %s %s" % (leftStr, mainStr, rightStr)
                return buf, None

            elif sel == SelectorType.tmBRACE:
                mainSlot = leftSlot = rightSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    leftSlot, _ = self.makeLatex(ast.children[1])
                if len(ast.children) > 2:
                    rightSlot, _ = self.makeLatex(ast.children[2])
                if rightSlot == "":
                    rightSlot = "."
                else:
                    rightSlot = " " + rightSlot

                # If mainSlot already contains an array or matrix, do not wrap it redundantly
                stripped_main = mainSlot.strip()
                if stripped_main.startswith(
                    "\\begin{array}"
                ) or stripped_main.startswith("\\begin{matrix}"):
                    buf += "\\left %s %s \\right%s" % (
                        leftSlot,
                        mainSlot,
                        rightSlot,
                    )
                else:
                    buf += "\\left %s \\begin{array}{l} %s \\end{array} \\right%s" % (
                        leftSlot,
                        mainSlot,
                        rightSlot,
                    )
                return buf, None

            elif sel == SelectorType.tmBRACK:
                mainSlot = leftSlot = rightSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    leftSlot, _ = self.makeLatex(ast.children[1])
                if len(ast.children) > 2:
                    rightSlot, _ = self.makeLatex(ast.children[2])
                if mainSlot == "":
                    mainSlot = "\\space"
                if leftSlot == "":
                    leftSlot = "."
                if rightSlot == "":
                    rightSlot = "."

                # Same logic as tmBRACE: check if mainSlot already has array/matrix
                stripped_main = mainSlot.strip()
                if stripped_main.startswith(
                    "\\begin{array}"
                ) or stripped_main.startswith("\\begin{matrix}"):
                    # Already has array/matrix → don't wrap
                    buf += "\\left%s %s \\right%s" % (leftSlot, mainSlot, rightSlot)
                else:
                    # No array → wrap with array{l}
                    buf += "\\left%s \\begin{array}{l} %s \\end{array} \\right%s" % (
                        leftSlot,
                        mainSlot,
                        rightSlot,
                    )
                return buf, None

            elif sel == SelectorType.tmROOT:
                mainSlot = radiSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    radiSlot, _ = self.makeLatex(ast.children[1])
                buf += "\\sqrt[%s] { %s }" % (radiSlot, mainSlot)
                return buf, None

            elif sel == SelectorType.tmFRACT:
                numSlot = denSlot = ""
                if len(ast.children) > 0:
                    numSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    denSlot, _ = self.makeLatex(ast.children[1])
                buf += "\\frac { %s } { %s }" % (numSlot, denSlot)
                return buf, None

            elif sel == SelectorType.tmUBAR:
                mainSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                buf += " {\\underline{ %s }} " % mainSlot
                return buf, None

            elif sel == SelectorType.tmOBAR:
                mainSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                buf += " \\overline{%s} " % mainSlot.strip()
                return buf, None

            elif sel in (
                SelectorType.tmSUM,
                SelectorType.tmPROD,
                SelectorType.tmCOPROD,
                SelectorType.tmUNION,
                SelectorType.tmINTER,
                SelectorType.tmINTOP,
                SelectorType.tmSUMOP,
            ):
                mainSlot = upperSlot = lowerSlot = operatorSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    lowerSlot, _ = self.makeLatex(ast.children[1])
                if len(ast.children) > 2:
                    upperSlot, _ = self.makeLatex(ast.children[2])
                if len(ast.children) > 3:
                    operatorSlot, _ = self.makeLatex(ast.children[3])
                mainStr = "{ %s }" % mainSlot if mainSlot else ""
                lowerStr = "\\limits_{ %s }" % lowerSlot if lowerSlot else ""
                upperStr = "^ %s" % upperSlot if upperSlot else ""
                buf += "%s %s %s %s" % (operatorSlot, lowerStr, upperStr, mainStr)
                return buf, None

            elif sel == SelectorType.tmINTEG:
                mainSlot = upperSlot = lowerSlot = operatorSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    lowerSlot, _ = self.makeLatex(ast.children[1])
                if len(ast.children) > 2:
                    upperSlot, _ = self.makeLatex(ast.children[2])
                if len(ast.children) > 3:
                    operatorSlot, _ = self.makeLatex(ast.children[3])
                mainStr = "{ %s }" % mainSlot if mainSlot else ""
                lowerStr = "\\limits_{ %s }" % lowerSlot if lowerSlot else ""
                upperStr = "^ %s" % upperSlot if upperSlot else ""
                intOp = "\\int"
                var = tmpl.variation
                if var & 0x0002:
                    intOp = "\\iint"
                if var & 0x0003 == 0x0003:
                    intOp = "\\iiint"
                if var & 0x0004:
                    intOp = "\\oint"
                if operatorSlot:
                    intOp = operatorSlot
                buf += "%s %s %s %s" % (intOp, lowerStr, upperStr, mainStr)
                return buf, None

            elif sel == SelectorType.tmLIM:
                mainSlot = lowerSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    lowerSlot, _ = self.makeLatex(ast.children[1])
                mainStr = "\\mathop { %s }" % mainSlot if mainSlot else ""
                lowerStr = "\\limits_{ %s }" % lowerSlot if lowerSlot else ""
                buf += "%s %s" % (mainStr, lowerStr)
                return buf, None

            elif sel == SelectorType.tmSUP:
                subSlot = supSlot = ""
                if len(ast.children) > 0:
                    subSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    supSlot, _ = self.makeLatex(ast.children[1])
                buf += " ^ { %s } " % supSlot
                if subSlot:
                    buf += " { %s } " % subSlot
                return buf, None

            elif sel == SelectorType.tmSUB:
                subSlot = supSlot = ""
                if len(ast.children) > 0:
                    subSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    supSlot, _ = self.makeLatex(ast.children[1])
                subFmt = "_{ %s }" % subSlot if subSlot else ""
                supFmt = "^{ %s }" % supSlot if supSlot else ""
                buf += "%s %s" % (subFmt, supFmt)
                return buf, None

            elif sel == SelectorType.tmSUBSUP:
                subSlot = supSlot = ""
                if len(ast.children) > 0:
                    subSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    supSlot, _ = self.makeLatex(ast.children[1])
                subFmt = "_{ %s }" % subSlot if subSlot else ""
                supFmt = "^{ %s }" % supSlot if supSlot else ""
                buf += "%s %s" % (subFmt, supFmt)
                return buf, None

            elif sel == SelectorType.tmVEC:
                mainSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                mainStr = "{ %s }" % mainSlot if mainSlot else ""
                variationsMap = {
                    0x0001: "left",
                    0x0002: "right",
                    0x0004: "tvVE_UNDER",
                    0x0008: "harpoonup",
                }
                topStr = "\\overset\\"
                for vCode in [0x0001, 0x0002, 0x0004, 0x0008]:
                    if vCode & tmpl.variation:
                        topStr += variationsMap[vCode]
                if tmpl.variation < 8:
                    topStr += "arrow"
                buf += "%s %s" % (topStr, mainStr)
                return buf, None

            elif sel == SelectorType.tmHAT:
                mainSlot = topSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    topSlot, _ = self.makeLatex(ast.children[1])
                mainStr = "{ %s }" % mainSlot if mainSlot else ""
                topStr = " %s " % topSlot if topSlot else ""
                buf += "%s %s" % (topStr, mainStr)
                return buf, None

            elif sel == SelectorType.tmARC:
                mainSlot = topSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    topSlot, _ = self.makeLatex(ast.children[1])
                mainStr = "{ %s }" % mainSlot if mainSlot else ""
                topStr = "\\overset %s" % topSlot if topSlot else ""
                buf += "%s %s" % (topStr, mainStr)
                return buf, None

            elif sel == SelectorType.tmTILDE:
                mainSlot = topSlot = ""
                if len(ast.children) > 0:
                    mainSlot, _ = self.makeLatex(ast.children[0])
                if len(ast.children) > 1:
                    topSlot, _ = self.makeLatex(ast.children[1])
                mainStr = "{ %s }" % mainSlot if mainSlot else ""
                topStr = "\\widetilde " if not topSlot else "%s " % topSlot
                buf += "%s %s" % (topStr, mainStr)
                return buf, None

            else:
                for _ast in ast.children:
                    _latex, _ = self.makeLatex(_ast)
                    buf += _latex
                return buf, None

        elif ast.tag == RecordType.PILE:
            idx = 0
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)
                if idx > 0:
                    buf += " \\\\ "
                buf += _latex
                idx += 1
            return buf, None

        elif ast.tag == RecordType.MATRIX:
            matrixCol = int(ast.value.cols)
            idx = 0
            for _ast in ast.children:
                _latex, _ = self.makeLatex(_ast)
                if idx == 0:
                    buf += " \\begin{array} {l} "
                    buf += _latex
                    idx += 1
                    continue
                if idx % matrixCol == 0:
                    buf += " \\\\ "
                else:
                    buf += " & "
                buf += _latex
                idx += 1
            buf += " \\end{array} "
            return buf, None

        elif ast.tag == RecordType.LINE:
            buf += self._children_latex(ast.children)
            return buf, None

        elif ast.tag == RecordType.EMBELL:
            buf += self._embell_text(ast.value.embellType)
            return buf, None

        return "", None

    @classmethod
    def OpenBytes(cls, bts):
        return cls.Open(BytesIO(bts))

    @classmethod
    def Open(cls, reader):
        ole, err = Ole.Open(reader)
        if err is not None:
            return None, err

        dir_list, err = ole.ListDir()
        if err is not None:
            return None, err

        for file in dir_list:
            if "Equation Native" == file.Name():
                root = dir_list[0]
                stream_reader = ole.OpenFile(file, root)
                hdrBuffer = stream_reader.read(oleCbHdr)
                if hdrBuffer is not None and len(hdrBuffer) == oleCbHdr:
                    hdrReader = BytesIO(hdrBuffer)
                    cbHdr = Helper.bytes2int(hdrReader.read(2))
                    if cbHdr is None or cbHdr != oleCbHdr:
                        return None, "MTEF.Open: read byte error"
                    hdrReader.seek(4 + 2, 1)
                    cbSize = Helper.bytes2int(hdrReader.read(4))
                    stream_reader.seek(cbHdr, 0)
                    eqnBody = stream_reader.read(cbSize)
                    eqn = MTEF()
                    eqn.reader = BytesIO(eqnBody)
                    eqn.readRecord()
                    eqn.makeAST()
                    return eqn, None
                return None, "MTEF.Open: read byte error"

        return None, "Equation Native stream not found"
