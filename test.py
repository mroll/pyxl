from pyxl import Pyxl

xlfile = Pyxl('TestFile.xlsx')
sheet1 = xlfile[0]

print sheet1[0]
print sheet1[1,2]
print sheet1["Sequence (5'-3')"]
