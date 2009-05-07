NB. example
cocurrent 'base'

NB. OpenOffice.org calc
NB. some differences from VBA
NB. cell position is (col,row) and 0-base
NB. usually use olemethod instead of oleget/oleset
NB. file names use URL format eg. file:///C:/test.xls (always forward slash)
NB. might not coerce 1 to VT_BOOL TRUE so need to specify VT_... as left argument
NB. use array argument (discuss later)
test=: 3 : 0
f1=. 'file:///', '/' I.@('\'&=)@]} jpath 'Addons/tables/wdooo/test1.xls'
p=: conew 'wdooo'
try.
  'base temp'=: olecreate__p 'com.sun.star.ServiceManager'
  olemethod__p base ; 'createInstance' ; 'com.sun.star.frame.Desktop'
  desktop=: oleid__p temp
  (VT_BSTR, VT_BSTR, VT_I4, VT_UNKNOWN+VT_ARRAY) olemethod__p desktop ; 'loadComponentFromURL' ; f1 ; '_blank' ; 0 ; VT_UNKNOWN olevector__p ''
  doc=: oleid__p temp
  olemethod__p doc ; 'getSheets'
  olemethod__p temp ; 'getByIndex' ; 0
  sheet=: oleid__p temp
  olemethod__p sheet ; 'getCellByPosition' ; 3 ; 9
  olemethod__p temp ; 'SetString' ; 'OOo Calc'
  olemethod__p sheet ; 'getCellByPosition' ; 4 ; 9
  olemethod__p temp ; 'SetValue' ; 123
  olemethod__p sheet ; 'getCellByPosition' ; 5 ; 9
  olemethod__p temp ; 'SetFormula' ; '=DATE(2007;2;28)'
  olemethod__p doc ; 'store'
  VT_BOOL olemethod__p doc ; 'close' ; 1
  olemethod__p desktop ; 'terminate'
NB. clean up
  olerelease__p sheet
  olerelease__p doc
  olerelease__p desktop
  1!:2&2 'success'
catch.
  1!:2&2 oleqer__p ''
end.
destroy__p ''
)

