NB. build.ijs

writesourcex_jp_ '~Addons/tables/wdooo/source';'~Addons/tables/wdooo/wdooo.ijs'

f=. 3 : 0
(jpath '~addons/tables/wdooo/',y) fcopynew jpath '~Addons/tables/wdooo/',y
)

mkdir_j_ jpath '~addons/tables/wdooo'
f 'wdooo.ijs'
f 'msexcel.ijs'
f 'oocalc.ijs'
f 'test1.xls'
