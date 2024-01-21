from collections import defaultdict
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from struct import unpack
from base64 import b64decode, b64encode
from quopri import decodestring
from email import message_from_string
import re

# helper functions
def hex_str(n):
    return hex(n)[2:].upper().rjust(2, '0')

def hex_str_arr(bin_array, delim=''):
    return delim.join(map(hex_str, bin_array))

def dt_winminutes(m, tz='UTC'):
    # Windows timestamp: minutes since 1601-01-01
    try:
        dt = datetime(1601, 1, 1) + timedelta(seconds=m * 60)
        return dt.replace(tzinfo=ZoneInfo(tz))
    except OverflowError:
        print('datetime overflow: ' + str(m))
        return datetime.max

def dt_macabsolute(s, tz='UTC'):
    # Apple Mac Absolute timestamp: seconds since 2001-01-01
    dt = datetime(2001, 1, 1) + timedelta(seconds=s)
    return dt.replace(tzinfo=ZoneInfo(tz))

def localtime(dt, tz='America/New_York'):
    # Return local time in Y-M-D H:M:S format
    return dt.astimezone(ZoneInfo(tz)).strftime('%Y-%m-%d %H:%M:%S')

def ol_days_of_week(byte):
    # Return day of week list from byte
    dows = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA']
    b = filter(lambda i: i[1], zip(dows, map(int, reversed(bin(byte)[2:]))))
    return [i[0] for i in b]

def ol_color(byte):
    # Two formats - RRBBGG and 0R0B0G, but we can parse them the same
    rbg = unpack('<xBxBxB', byte)
    # Then return an RBG hex code
    rbg = [hex(i)[2:].rjust(2, '0') for i in rbg]
    return '#' + ''.join(rbg)

def ol_type_code(byte):
    # Four character type codes, stored backwards. Return null if 00000000
    if byte == b'\x00\x00\x00\x00':
        return None
    else:
        return byte.decode()[::-1]

def parse_long_list(chunk):
    # This is a list of N long long numbers, first four are the length
    chunk = chunk[4:]
    fmt = '<' + str(int(len(chunk) / 8)) + 'q'
    return list(unpack(fmt, chunk))

def parse_int_list(chunk):
    # This is a list of N ints, first four are the length
    chunk = chunk[4:]
    fmt = '<' + str(int(len(chunk) / 4)) + 'i'
    return list(unpack(fmt, chunk))

def parse_date_list(chunk):
    # This is a list of ints, each representing a date in minutes
    fmt = '<' + str(int(len(chunk) / 4)) + 'i'
    return [dt_winminutes(b).date() for b in unpack(fmt, chunk)]

def a_b_d_y(s):
    if s:
        return datetime.strptime(s, '%a, %b %d, %Y')
    else:
        return None

def get_first(items, key):
    for item in items:
        if key in item.data:
            print(item.RecordID, item.data[key])
            break

def stats(items, key):
    types = sorted(set(type(i) for i in items), key=lambda t: str(t))
    for t in types:
        a = len([i for i in items if type(i) is t])
        b = len([i for i in items if type(i) is t and key in i.data])
        c = len([i for i in items if type(i) is t and i.data.get(key)])
        if b > 0:
            print(t)
            print('\t', str(b) + '/' + str(a) + ' (' + str(c) + ')')
            print('\n')

def diff(a, b):
    nota = list()
    notb = list()
    diff = list()
    for k in set(list(a.keys()) + list(b.keys())):
        if k not in a:
            nota.append((k, b[k]))
        elif k not in b:
            notb.append((k, a[k]))
        elif a[k] != b[k]:
            diff.append((k, (a[k], b[k])))
    return (nota, notb, diff)

def correlate(items, key):
    keys = defaultdict(dict)
    unmap = unmapped_keys(items)
    for i in items:
        val = i.data.get(key, None)
        base = -1 if val is None else int(bool(val))
        for k, v in i.__dict__.items():
            if k == 'data':
                continue
            comp = -1 if v is None else int(bool(v))
            keys['m:' + k][(base, comp)] = keys['m:' + k].get((base, comp), 0) + 1
        for k in unmap:
            if k == key:
                continue
            v = i.data.get(k, None)
            comp = -1 if v is None else int(bool(v))
            keys['u:' + k][(base, comp)] = keys['u:' + k].get((base, comp), 0) + 1
    return keys

def unmapped_keys(items):
    return sorted(set(sum([list(i.data.keys()) for i in items], start=list())))

def json_serializer(obj):
    if isinstance(obj, datetime):
        return obj.strftime('%Y-%m-%d %H:%M:%S.%f %z')
    elif isinstance(obj, date):
        return obj.strftime('%Y-%m-%d')
    else:
        raise TypeError ("Type %s not serializable" % type(obj))

def fix_attachment_encoding(attachment):
    msg = message_from_string(attachment)
    detect_encoding(msg)
    return msg.as_string()

def detect_encoding(msg):
    if msg.is_multipart():
        for part in msg.get_payload():
            detect_encoding(part)
    else:
        # Decode filename
        fn = msg.get_filename()
        if fn:
            fn = encoded_words_to_text(fn)
            if msg.get_param('name'):
                msg.set_param('name', fn)
            if msg.get_param('filename', header='content-disposition'):
                msg.set_param('filename', fn, header='content-disposition')

        # Detect content-type
        if msg.get_content_type() == 'application/octet-stream':
            p = msg.get_payload()
            d = b64decode(p)
            # Common image formats
            if d[6:10] in (b'JFIF', b'Exif') or d[:4] == b'\xff\xd8\xff\xdb':
                ct = 'image/jpeg'
            elif d[:6] in (b'GIF87a', b'GIF89a'):
                ct = 'image/gif'
            elif d[:2] in (b'MM', b'II'):
                ct = 'image/tiff'
            elif d.startswith(b'BM'):
                ct = 'image/bmp'
            # Maybe this is an email?
            elif 'content-type: text' in p.lower():
                ct = 'multipart/related'
            else:
                ct = 'application/octet-stream'
            msg.set_type(ct)

def get_ext(content_type):
    if content_type == 'image/jpeg':
        return 'jpg'
    elif content_type == 'application/pdf':
        return 'pdf'
    elif content_type == 'application/ics':
        return 'ics'
    elif content_type == 'multipart/related':
        return 'eml'
    elif content_type == 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
        return 'pptx'
    elif content_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        return 'docx'
    elif content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        return 'xlsx'
    else:
        raise ValueError(content_type)

def encoded_words_to_text(encoded_words):
    decoded_word = ''
    for word in encoded_words.split():
        encoded_word_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='
        charset, encoding, encoded_text = re.match(encoded_word_regex, word).groups()
        if encoding == 'B':
            byte_string = b64decode(encoded_text)
        elif encoding == 'Q':
            byte_string = decodestring(encoded_text)
        decoded_word += byte_string.decode(charset)
    return decoded_word
