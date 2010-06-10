NB. Excel using vtable oleautomation
NB. only works if ms office excel has been installed
NB. (but I got no ms excel to test, use at your own risk.)

'require'~'tables/wdooo'

cocurrent 'base'

NB. MS Office Excel
NB. some differences from VBA
NB. cell position is (col,row) and 1-base
NB. use olemethod and oleget/oleset
test=: 3 : 0
PATH=. jpath '~addons/tables/wdooo'
f1=. 'file:///', PATH, '/test1.xls'
smoutput f1
p=. '' conew 'wdooo'
try.
  'me base temp'=: olecreate__p 'Excel.Application'
  oleget__p me ; base ; 'workbooks'
  wb=: oleid__p me
  olemethod__p me ; wb ; 'open' ; f1
  oleget__p me ; base ; 'activeworkbook'
  awb=: oleid__p me
  oleget__p me ; awb ; 'worksheets' ; 1     NB. sheet 1-base
  aws=: oleid__p me
NB. write string into a cell
  oleget__p me ; aws ; 'range' ; xlcell 3 10
  range=. oleid__p me
  oleput__p me ; range ; 'value' ; 'Ms Excel'
  oleput__p me ; range ; 'numberformat' ; '@'
  olerelease__p range
NB. write number into a cell
  oleget__p me ; aws ; 'range' ; xlcell 4 10
  range=. oleid__p me
  oleput__p me ; range ; 'value' ; 123
  oleput__p me ; range ; 'HorizontalAlignment' ; _4152    NB. xlRight  0xffffefc8
  olerelease__p range
NB. write date into a cell
  oleget__p me ; aws ; 'range' ; xlcell 5 10
  range=. oleid__p me
  oleput__p me ; range ; 'value' ; (-&36522) 76533        NB. 2007-2-28
  oleput__p me ; range ; 'numberformat' ; 'yyyy-mmm-dd'
  olerelease__p range
NB. save and cleanup
  (olerelease__p ::0:) aws
  VT_BOOL (oleput__p ::0:) me ; base ; 'DisplayAlerts' ; 0
  (olemethod__p ::0:) me ; awb ; 'save'
  (olemethod__p ::0:) me ; awb ; 'close'
  (olerelease__p ::0:) awb
  (olerelease__p ::0:) wb
  (olemethod__p ::0:) me ; base ; 'quit'
  VT_BOOL (oleput__p ::0:) me ; base ; 'DisplayAlerts' ; 1
  (oledestroy__p ::0:) me
  smoutput 'success'
catch.
  smoutput oleqer__p ''
end.
destroy__p ''
)

xlcell=: 3 : 0
'c r'=. y
c1=. <.c%26 [ c2=. 26|c
' '-.~(c1{' ABCDEFGHI'), (c2{'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), ": 1+r
)

