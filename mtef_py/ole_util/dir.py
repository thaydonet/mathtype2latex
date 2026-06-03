from io import BytesIO
from .helper import Helper


class FileType:
    EMPTY = 0
    USERSTORAGE = 1
    USERSTREAM = 2
    LOCKBYTES = 3
    PROPERTY = 4
    ROOT = 5


class File:
    def __init__(self):
        self.NameBts = []
        self.Bsize = 0
        self.Type = FileType.EMPTY
        self.Flag = 0
        self.Left = 0
        self.Right = 0
        self.Child = 0
        self.Guid = []
        self.Userflags = 0
        self.Time = []
        self.Sstart = 0
        self.Size = 0
        self.Proptype = 0

    def fromBytes(self, bts):
        reader = BytesIO(bts)
        self.NameBts = []
        for _ in range(32):
            self.NameBts.append(reader.read(2))
        self.Bsize = Helper.bytes2int(reader.read(2))
        self.Type = Helper.bytes2int(reader.read(1))
        self.Flag = Helper.bytes2int(reader.read(1))
        self.Left = Helper.bytes2int(reader.read(4))
        self.Right = Helper.bytes2int(reader.read(4))
        self.Child = Helper.bytes2int(reader.read(4))
        self.Guid = []
        for _ in range(8):
            self.Guid.append(Helper.bytes2int(reader.read(2)))
        self.Userflags = Helper.bytes2int(reader.read(4))
        self.Time = []
        for _ in range(2):
            self.Time.append(Helper.bytes2int(reader.read(8)))
        self.Sstart = Helper.bytes2int(reader.read(4))
        self.Size = Helper.bytes2int(reader.read(4))
        self.Proptype = Helper.bytes2int(reader.read(4))

    def Name(self):
        return b"".join(self.NameBts[: int(self.Bsize / 2 - 1)]).decode("utf-16-le")
