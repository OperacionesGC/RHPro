Attribute VB_Name = "MdlClasesGlobales"
Option Explicit

'Global TablaDeSimbolos As New Collection
Global eval As New CEval
Global ErrorEnExpresion As Boolean

' Caches
Global objCache As New CNuevaCache
Global objCache_NovGlobales As New CNuevaCache
Global objCache_Acu_Liq_Monto As New CNuevaCache
Global objCache_Acu_Liq_MontoReal As New CNuevaCache
Global objCache_Acu_Liq_Cantidad As New CNuevaCache

Global objCache_BusquedasGlobales As New CNuevaCache    'EAM (v6.44) - Guada los resultados de búsquedas globales

'FGZ  - 10/02/2004
Global objCache_detliq_Monto As New CNuevaCache
Global objCache_detliq_Cantidad As New CNuevaCache
