# Excel Math : Python2
# From : https://stackoverflow.com/questions/48983939/
'''
1 =A
27 = AA
'''

import string
from functools import reduce


def divmod_excel(n):
    a, b = divmod(n, 26)
    if b == 0:
        return a - 1, b + 26
    return a, b


def to_excel(num):
    chars = []
    while num > 0:
        num, d = divmod_excel(num)
        chars.append(string.ascii_uppercase[d - 1])
    return ''.join(reversed(chars))


def from_excel(chars):
    return reduce(lambda r, x: r * 26 + x + 1,
                  map(string.ascii_uppercase.index, chars), 0)
