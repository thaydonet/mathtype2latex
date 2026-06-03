from .helper import Helper

DEBUG = False


class StreamReader:
    def __init__(
        self,
        sat=[],
        start=0,
        reader=None,
        offset_of_sector=0,
        offset_in_sector=0,
        size_sector=0,
        size=0,
        offset=0,
        sector_pos=None,
    ):
        self.sat = sat
        self.start = start
        self.reader = reader
        self.offset_of_sector = offset_of_sector
        self.offset_in_sector = offset_in_sector
        self.size_sector = size_sector
        self.size = size
        self.offset = offset
        self.sector_pos = sector_pos

    def read(self, read_size=None):
        if self.offset_of_sector == Helper.ENDOFCHAIN:
            return b""
        ans = []
        pos = (
            self.sector_pos(self.offset_of_sector, self.size_sector)
            + self.offset_in_sector
        )
        self.reader.seek(pos, 0)
        readed = 0
        reaminLen = read_size - readed
        while reaminLen > self.size_sector - self.offset_in_sector:
            to_read_size = self.size_sector - self.offset_in_sector
            read_bytes = self.reader.read(self.size_sector - self.offset_in_sector)
            if read_bytes:
                ans.append(read_bytes)
            if read_bytes is None or len(read_bytes) != to_read_size:
                return b"".join(ans)
            else:
                readed += len(read_bytes)
                self.offset_in_sector = 0
                if self.offset_of_sector >= len(self.sat):
                    return b"".join(ans)
                else:
                    self.offset_of_sector = self.sat[self.offset_of_sector]
                if self.offset_of_sector == Helper.ENDOFCHAIN:
                    return b"".join(ans)
                pos = (
                    self.sector_pos(self.offset_of_sector, self.size_sector)
                    + self.offset_in_sector
                )
                self.reader.seek(pos, 0)
            reaminLen = read_size - readed

        read_bytes = self.reader.read(read_size - readed)
        if read_bytes:
            ans.append(read_bytes)
        if read_bytes and len(read_bytes) == read_size - readed:
            self.offset_in_sector += len(read_bytes)
            return b"".join(ans)
        else:
            return b"".join(ans)

    def seek(self, offset, whence=0):
        if whence == 0:
            self.offset_of_sector = self.start
            self.offset_in_sector = 0
            self.offset = offset
        else:
            self.offset += offset

        if self.offset_of_sector == Helper.ENDOFCHAIN:
            return self.offset

        need_go_to = False
        while offset >= self.size_sector - self.offset_in_sector:
            self.offset_of_sector = self.sat[self.offset_of_sector]
            offset -= self.size_sector - self.offset_in_sector
            self.offset_in_sector = 0
            if self.offset_of_sector == Helper.ENDOFCHAIN:
                need_go_to = True
                break

        if not need_go_to:
            if self.size <= self.offset:
                self.offset = self.size
            else:
                self.offset_in_sector += offset

        return self.offset
