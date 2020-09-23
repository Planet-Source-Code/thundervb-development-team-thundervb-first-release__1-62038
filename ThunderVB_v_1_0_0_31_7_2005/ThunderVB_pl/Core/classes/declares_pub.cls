VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "declares_pub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Declares..

Public Type PlugIn_Entry

    used As Boolean                     'Is this entry used?? [false if this slot is empty]
    Loaded As Boolean                   'Is this entry loaded??
    Name As String                      'Name of the plugin
    Version As String                   'Version of teh plugin , string format
    VersionNum() As Single              'Version in x.x.x format
    Desciption As String                'A short Desciption for this plugin
    DesciptionFull As String            'A Desciption for this plugin
    type As pl_type                     'plugin Type
    Speed As pl_Speed_Req               'How much power consuming this plugin is
    Id As Long                          'plugin ID , this should be unique [email to the athors to get one]
    interface As ThunderVB_pl_int_v1_0  'plugin interface class
    dllfile As String                   'name of the dll ..
    
End Type

Public Type PlugIn_List

    plugins() As PlugIn_Entry
    count As Long 'hehe no more that 2 Giga plugins :P :P
    
End Type

Public Type ID_List

    Id() As Long
    count As Long
    
End Type


Public Type Error

    ErrorID As Long     'id of the error.. , 0 means no error, -1 is userdefined(ony text info)
    ErrorString As String 'any data that can be used to identyfy the location of the error
    ErrorInfo As String   'Text explanaition of the error
    
End Type



Public Type ButArray

    count As Long
    items() As clsComBut
    
End Type


Public Type cph_list

    items() As ThunderVB_pl_cph_v1_0
    count As Long
    
End Type


Public Enum ThunVB_Errors

    'File IO error
    tvb_File_Does_Not_Exist = 53
    'GPF error codes
    tvb_GPF_Error = vbObjectError + 512
    'CP hooks error codes..
    tvb_CPH_Before_First_Exists = tvb_GPF_Error + 1
    tvb_CPH_Before_Last_Exists = tvb_CPH_Before_First_Exists + 1
    tvb_CPH_After_First_Exists = tvb_CPH_Before_Last_Exists + 1
    tvb_CPH_After_Last_Exists = tvb_CPH_After_First_Exists + 1
    'Resource System Error
    tvb_Res_NoDataEntrys = tvb_CPH_After_Last_Exists + 1
    tvb_Res_Header_Corupted = tvb_Res_NoDataEntrys + 1
    tvb_Res_No_Languages = tvb_Res_Header_Corupted + 1
    
End Enum
