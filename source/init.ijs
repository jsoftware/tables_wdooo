NB. init

'require'~'dll'

NB.*symdat v pointer to J name data, used in structs
NB.   symdat symget <'name'
symdat_z_=: 3 : 0   NB.
had=. {.memr y,(IF64{4 8),1,JPTR
had+{.memr had,0,1,JINT
)

