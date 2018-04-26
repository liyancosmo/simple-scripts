import xlrd

cut = 10
substitution = "-9999"

filename = "/Users/cosmozhang/Downloads/FrontiersPS-Tree-Ring Data.xlsx"
outfilename = "/Users/cosmozhang/Downloads/FrontiersPS-Tree-Ring Data.rwl"

namemap = {
  "Wutaishan": "WT",
  "Hengshan": "HS"
}

def safeInt(val):
  ret = None
  try:
    ret = int(val)
    return ret
  except Exception, e:
    return None
  return ret
def safeFloat(val):
  ret = None
  try:
    ret = float(val)
    return ret
  except Exception, e:
    return None
  return ret

def parseColumn(name, years, data):
  retlines = []
  datalen = len(data)
  ys = years[0]
  ye = years[0]
  for y in years:
    if y < ys:
      ys = y
    if y > ye:
      ye = y
  seqy = range(ys,ye+2)
  seqd = [None for i in seqy]
  for i in xrange(datalen):
    d = data[i]
    if not safeFloat(d) is None:
      y = years[i]
      seqd[y-ys] = d
  # begin formatting
  line = None
  for y in seqy:
    d = seqd[y-ys]
    hasline = not line is None
    if (y % cut) == 0 and not line is None:
      retlines.append(line)
      line = None
    if d is None:
      if hasline:
        if line is None:
          line = [name, ("%04d" % y), substitution]
        else:
          line.append(substitution)
        retlines.append(line)
        line = None
        continue
    else:
      if line is None:
        line = [name, ("%04d" % y)]
      line.append(d)
  if not line is None:
    retlines.append(line)
    line = None
  return retlines

def parseSheet(sheet):
  sheetname = sheet.name
  nrows = sheet.nrows
  ncols = sheet.ncols
  if nrows < 2 or ncols < 2:
    return None
  years = sheet.col_values(0)[1:]
  years = [safeInt(y) for y in years]
  names = sheet.row_values(0)[1:]
  names = [("%s%02d" % (namemap[sheetname], safeInt(n))) for n in names]
  retlines = []
  for i in xrange(ncols-1):
    coldata = sheet.col_values(i+1)[1:]
    retlines += parseColumn(names[i], years, coldata)
  return retlines

def parseWorkbook(workbook):
  retlines = []
  for sheet in workbook.sheets():
    retlines += parseSheet(sheet)
  retstr = ""
  for line in retlines:
    retstr += "%-5s% 4s" % (line[0], line[1])
    for d in line[2:]:
      retstr += "% 8s" % d
    retstr += "\r\n"
  return retstr

workbook = xlrd.open_workbook(filename)

outfile = open(outfilename, "w")
outfile.write(parseWorkbook(workbook))

