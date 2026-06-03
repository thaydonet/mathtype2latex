from io import BytesIO
from .helper import Helper


class Header:
    def __init__(self):
        self.Id = []
        self.Clid = []
        self.Verminor = 0
        self.Verdll = 0
        self.Byteorder = 0
        self.Lsectorb = 0
        self.Lssectorb = 0
        self.Cfat = 0
        self.Dirstart = 0
        self.Sectorcutoff = 0
        self.Sfatstart = 0
        self.Csfat = 0
        self.Difstart = 0
        self.Cdif = 0
        self.Msat = []

    def fromBytes(self, bts):
        reader = BytesIO(bts)
        self.Id = []
        for _ in range(2):
            self.Id.append(Helper.bytes2int(reader.read(4)))
        self.Clid = []
        for _ in range(4):
            self.Clid.append(Helper.bytes2int(reader.read(4)))
        self.Verminor = Helper.bytes2int(reader.read(2))
        self.Verdll = Helper.bytes2int(reader.read(2))
        self.Byteorder = Helper.bytes2int(reader.read(2))
        self.Lsectorb = Helper.bytes2int(reader.read(2))
        self.Lssectorb = Helper.bytes2int(reader.read(2))
        reader.read(2)
        reader.read(8)
        self.Cfat = Helper.bytes2int(reader.read(4))
        self.Dirstart = Helper.bytes2int(reader.read(4))
        reader.read(4)
        self.Sectorcutoff = Helper.bytes2int(reader.read(4))
        self.Sfatstart = Helper.bytes2int(reader.read(4))
        self.Csfat = Helper.bytes2int(reader.read(4))
        self.Difstart = Helper.bytes2int(reader.read(4))
        self.Cdif = Helper.bytes2int(reader.read(4))
        self.Msat = []
        for _ in range(109):
            self.Msat.append(Helper.bytes2int(reader.read(4)))

    @classmethod
    def parseHeader(cls, bts):
        header = Header()
        header.fromBytes(bts)
        if (
            header.Id[0] != 0xE011CFD0
            or header.Id[1] != 0xE11AB1A1
            or header.Byteorder != 0xFFFE
        ):
            return None, "not an OLE file"
        return header, None
