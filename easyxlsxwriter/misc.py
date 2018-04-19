def extension_chk(filename, ext):
    # Append a dot if not present
    if '.' not in ext:
        ext = '.' + ext
    return ext in filename[-(len(ext)):]


# Utility to quickly to generate range selectors
# Eg =Sheet1!$A$2:$A$7
def sheetrange(sheet, pt1, pt2=None):
    prefix = '=' + sheet + '!'
    pt1str = '$' + pt1.x + '$' + str(pt1.y)
    pt2str = ''
    if pt2 is not None:
        pt2str = ':$' + pt2.x + '$' + str(pt2.y)
    return prefix + pt1str + pt2str
