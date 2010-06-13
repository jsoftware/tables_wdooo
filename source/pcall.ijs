NB. =========================================================
NB. ------- include files from oleg pcall start -------------

coclass 'olegpcall'

NB.*acall c address call using function pointer
NB.   func=: FuncAddr acall 'i x *w ...'
NB.   res=. func arg1;arg2;...
acall=: 2 : '(''0 '',(":m),'' > '',n)&(15!:0)'

NB.*icall c interface call using v-table
NB.   iuQueryInterface=: IU_QueryInterface icall 'i x  *c *i' @ ;
NB.   res=. iObj iuQueryInterface arg1;arg2;...
icall=: 2 : '(''1 '',(":m),'' > '',n)&(15!:0)'

NB.*idef d interface index table definition
NB.   'IU_'idef IUnknown=: ;:'QueryInterface AddRef Release'
idef=: 4 : '((x&,)&.>y)=: i.#y'

coclass 'olegpole32'
coinsert 'olegpcall'

CoInitializeEx=:     'ole32 CoInitializeEx   > x x i'&cd
CLSIDFromProgID=: 'ole32 CLSIDFromProgID  > i *w *c'&cd
CLSIDFromString=: 'ole32 CLSIDFromString  > i *w *c'&cd
CoCreateInstance=: 'ole32 CoCreateInstance > i *c i i *c *x'&cd
CoGetObject=: 'ole32 CoGetObject      > i *w i *c *x'&cd

VariantClear=: 'oleaut32 VariantClear > i *x'&cd
VariantChangeType=: 'oleaut32 VariantChangeType > i *i *i i s'&cd
SysFreeString=: 'oleaut32 SysFreeString > n i'&cd
SysAllocStringLen=: 'oleaut32 SysAllocStringLen > i *w i'&cd
CoGetActiveObject=: 'oleaut32 GetActiveObject     > i *c i *x'&cd

GUID=: 'WWWWXXYYZZZZZZZZ'
GUID_NULL=: (#GUID) # 0{a.
VAR1=: 'VtR1R2R3Valu'
VAR2=: 'VtR1R2R3Val1Val2'
DISPPARAMS=: 'ArgsNamdCArgCNmd'
TYPEATTR=: GUID,'LcidReseCtorDtorScheSinsTpknCfCvCtSvBaTfMjMnAliaIdld'
TYPEDESC=: 'LptdVt00'
PARAMDESC=: 'LppdPf00'
ELEMDESC=: TYPEDESC,PARAMDESC
FUNCDESC=: 'MbidScodParmFunkInvkCalcCpCoOvCs',ELEMDESC,'Ff00'
IID_IUnknown=: '{00000000-0000-0000-C000-000000000046}'
IID_IDispatch=: '{00020400-0000-0000-C000-000000000046}'

'CLSCTX_INPROC_SERVER CLSCTX_LOCAL_SERVER'=: 16b0001 16b0004
CTX=: CLSCTX_INPROC_SERVER+CLSCTX_LOCAL_SERVER

'COINIT_APARTMENTTHREADED COINIT_MULTITHREADED' =: 2 0

'VT_EMPTY VT_NULL VT_I2 VT_I4  VT_R4 VT_R8 VT_CY VT_DATE'=: i.8
'VT_BSTR VT_DISPATCH VT_ERROR VT_BOOL'=: 8+i.4
'VT_VARIANT VT_UNKNOWN VT_DECIMAL'=: 12+i.3
'VT_PTR VT_SAFEARRAY VT_CARRAY VT_USERDEFINED'=: 26+i.4
'VT_VECTOR VT_ARRAY VT_BYREF VT_TYPEMASK'=: 16b1000 16b2000 16b4000 16b0fff

'DISP_METH DISP_GET DISP_PUT DISP_SET'=: 1 2 4 8

'IU_'idef IUnknown=: ;:'QueryInterface AddRef Release'
'ID_'idef IDispatch=: IUnknown,;:;<;._2(0 : 0)
  GetTypeInfoCount GetTypeInfo GetIDsOfNames Invoke
)
'IT_'idef ITypeInfo=: IUnknown,;:;<;._2(0 : 0)
  GetTypeAttr GetTypeComp GetFuncDesc GetVarDesc GetNames
  GetRefTypeOfImplType GetImplTypeFlags GetIDsOfNames Invoke
  GetDocumentation GetDllEntry GetRefTypeInfo AddressOfMember
  CreateInstance GetMops GetContainingTypeLib ReleaseTypeAttr
  ReleaseFuncDesc ReleaseVarDesc
)

iuQueryInterface=: IU_QueryInterface icall 'i x  *c *x' @ ;
iuAddRef=: IU_AddRef icall 'i x' @ [
iuRelease=: IU_Release icall 'i x' @ [
idGetIDsOfNames=: ID_GetIDsOfNames icall 'i x  *c *x x x *x' @ ;
idGetTypeInfo=: ID_GetTypeInfo icall 'i x  x x *x' @ ;
idInvoke=: ID_Invoke icall 'i x  x *c x s *x *x x x' @ ;
itGetTypeAttr=: IT_GetTypeAttr icall 'i x  *x' @ ;
itReleaseTypeAttr=: IT_ReleaseTypeAttr icall 'i x  x' @ ;
itGetFuncDesc=: IT_GetFuncDesc icall 'i x  x *x' @ ;
itReleaseFuncDesc=: IT_ReleaseFuncDesc icall 'i x  x' @ ;
itGetNames=: IT_GetNames icall 'i x  x *x x *x' @ ;
itGetDocumentation=: IT_GetDocumentation icall 'i x  x *x *x x x' @ ;
itGetRefTypeInfo=: IT_GetRefTypeInfo icall 'i x  x *x' @ ;

hex8=: ,~ '00000000' }.~ #
hfd8=: '0x' , hex8@hfd
herr=: hfd8 assert 0 <: ]

CreateObject=: 3 : 0
IID_IDispatch CreateObject y
:
herr CoCreateInstance (GetGuid y) ; 0 ; CTX ; (GetGuid x) ; p=. ,_2
{.p
)

GetObject=: 3 : 0
IID_IDispatch GetObject y
:
herr CoGetObject y ; 0 ; (GetGuid x) ; p=. ,_2
{.p
)

GetActiveObject=: 3 : 0
IID_IDispatch GetActiveObject y
:
herr CoGetActiveObject (GetGuid y) ; 0 ; p=. ,_2
{.p
)

GetGuid=: 3 : 0
f=. CLSIDFromProgID`CLSIDFromString@.('{'={.y)
herr f y ; guid=. 1#GUID
guid
)

h=: ([: ;:^:_1"1 [: <"1 hfd)@(([: , _4 (_2&(3!:4))@|.\ ])^:(2=3!:0))
si=: I.@E.~   NB. TYPEATTR si 'Cfun'
us=: 0&(3!:4)
mi=: [: {.@memr ,&(0 1,JINT)
mc=: ,&0@] memr@, ,&JCHAR@[
mI=: 4 : '{.memr y,x,1,JINT'
mS=: 4 : '{.us memr y,x,2,JCHAR'
and=: 17 b.

GetStr=: 3 : 0
if. 0=y do. ''return. end.
len=. mi _4+y    NB. BSTR length
val=. len mc y   NB. BSTR char pairs value
8 u: 6 u: val
)

GetStrSafeFree=: 3 : 0
if. 0=y=. {.y do. '' return. end.
r=. GetStr y
SysFreeString y
r
)

AllocStr=: 3 : 0
SysAllocStringLen y;#y
)

VariantAlloc=: 3 : '(2-2)#~4%~#VAR2'
VariantStr=: GetStr@(2&{)

BoolVar=: 3 : 'VT_BOOL,0,(_1 0{~0-:{.y),0'
IntVar=: 3 : 'VT_I4,0,({.y),0'
PtrVar=: 3 : 'VT_UNKNOWN,0,({.y),0'
FloatVar=: 3 : 'VT_R8,0,_2(3!:4)2(3!:5){.y'
StrVar=: 3 : 'VT_BSTR,0,(AllocStr y),0'
EmptyVar=: 3 : '({.y,VT_EMPTY),0,0,_1'

JVar=: 3 : 0
if. y-:a: do. EmptyVar'' return. end.
if. y-:<0 do. EmptyVar VT_NULL return. end.
if. 0<L.y do. PtrVar >y return. end.
select. 3!:0 y
case. 1;4 do. IntVar y
case. 8 do. FloatVar y
case. do. StrVar ,":,y
end.
)

VarJ=: 3 : 0
select. VT_TYPEMASK and {.y
case. VT_EMPTY do. a:
case. VT_NULL do. <0
case. VT_I2 do. _1(3!:4)1(3!:4) 2{y
case. VT_I4 do. 2{y
case. VT_DISPATCH;VT_UNKNOWN do. <2{y
case. VT_R4 do. _1(3!:5)2(3!:4) 2{y
case. VT_R8 do. _2(3!:4)2(3!:4) 2 3{y
case. VT_BSTR do. VariantStr y
case. VT_BOOL do. 0~:2{y
case. do. VariantStr y [ VariantChangeType y;y;0;VT_BSTR
end.
)

VTSTR=: ; <@cut;._2 (0 : 0)
void null short long float double CURRENCY DATE
BSTR IDispatch* SCODE boolean VARIANT IUnknown* WCHAR .
char BYTE WORD DWORD int64 uint64 int UINT
void HRESULT PTR SAFEARRAY CARRAY USERDEFINED LPSTR LPWSTR
. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
FILETIME BLOB STREAM STORAGE STREAMED_OBJECT STORED_OBJECT BLOB_OBJECT
CF CLSID BAD_TYPE
)

VtStr=: 3 : 'VTSTR >@{~ (<:#VTSTR) <. y and 16bfff'

TypeDesc=: 0&$: : (4 : 0)"0
if. 0=y do. 'void' return. end.
select. vt=. 16bfff and 4 mS y
case. VT_PTR do. '*',~x TypeDesc 0 mI y return.
case. VT_USERDEFINED do. x RefDesc 0 mI y return.
case. do. VtStr vt
end.
)

RefDesc=: 4 : 0
herr x itGetRefTypeInfo y ; rt=. ,_2
rt=. {.rt
r=. >@{. rt GetDoc _1
rt iuRelease ''
r
)

FuncDesc=: 4 : 0"0
herr x itGetFuncDesc y ; pfd=. ,_2
mid=. 0 mI pfd
if. c=. (FUNCDESC si 'Cp') mS pfd do.    NB. cParams, params count
  p=. (FUNCDESC si 'Parm') mI pfd        NB. Params
  r=. x <@TypeDesc p+(#ELEMDESC)*i.c
else. r=. '' end.
r=. r,~<x TypeDesc {.(FUNCDESC si 'Lptd') + pfd  NB. return type
herr x itReleaseFuncDesc {.pfd

res=. (c+1)#_1
herr x itGetNames mid ; res ; (#res) ; ,_1
res=. 0 (I.res=_1)}res
r (, ' '&,)&.> <@GetStrSafeFree"0 res    NB. names of arguments
)

GetDoc=: 4 : 0"0
herr x itGetDocumentation y ; (name=. ,_2) ; (doc=. ,_2) ; 0 ; 0
<@GetStrSafeFree"0 name,doc           NB. name;doc
)

FuncDoc=: 4 : 0"0
herr x itGetFuncDesc y ; pfd=. ,_2
mid=. 0 mI pfd
herr x itReleaseFuncDesc {.pfd
x GetDoc mid
)


NB. Global Interface Table
NB. Marchalling interface pointers between appartments and threads

CLSID_StdGlobalInterfaceTable=: '{00000323-0000-0000-C000-000000000046}'
IID_IGlobalInterfaceTable=: '{00000146-0000-0000-C000-000000000046}'

'GIT_'idef IGlobalInterfaceTable=: IUnknown,;:;<;._2(0 : 0)
  RegisterInterfaceInGlobal RevokeInterfaceFromGlobal GetInterfaceFromGlobal
)

gitRegisterInterfaceInGlobal=: GIT_RegisterInterfaceInGlobal icall 'i x  x *c *x' @ ;
gitRevokeInterfaceFromGlobal=: GIT_RevokeInterfaceFromGlobal icall 'i x  x' @ ;
gitGetInterfaceFromGlobal=: GIT_GetInterfaceFromGlobal icall 'i x  x *c *x' @ ;

gitGet=: 3 : 0
IID_IDispatch gitGet y
:
git=. IID_IGlobalInterfaceTable CreateObject CLSID_StdGlobalInterfaceTable
herr CLSIDFromString x ; iid=. 1#GUID
herr git gitGetInterfaceFromGlobal y;iid;p=. ,2-2
git iuRelease ''
{.p
)

CoInitializeEx^:IFCONSOLE 0;COINIT_APARTMENTTHREADED

NB. ------- include files from oleg pcall end ---------------
NB. =========================================================

